using System;
using System.Net;
using System.Text;
using System.Xml.Linq;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using UIAutomationClient;
using System.Windows.Forms;
using System.Xml.XPath;
using System.Diagnostics;

namespace InteractiveInspector
{
    public class Program
    {
        static string CurrentTargetName = null;
        static string CurrentProcess = null;
        static string SelectedElementId = null;
        static UiaHelper uiaHelper = new UiaHelper();
        static MsaaHelper msaaHelper = new MsaaHelper();
        static IHelper currentHelper = uiaHelper;
        static RECT CurrentWindowRect;
        static List<(RECT Rect, string Id)> RectIdList = new List<(RECT, string)>();
        static bool ShowAllHighlights = false;
        static XElement finderXml;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        public interface IHelper
        {
            IntPtr ResolveWindow(string windowName);
            XElement DumpXml(IntPtr hwnd);
            string GetElementProperties(string elementId);
            Program.RECT GetBoundingRectangleRECT(string elementId);
        }
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")] static extern short VkKeyScan(char ch);
        [DllImport("user32.dll")]  static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        const uint MOUSEEVENTF_MOVE = 0x0001;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint KEYEVENTF_KEYDOWN = 0x0000;
        const uint KEYEVENTF_KEYUP = 0x0002;

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        static bool PointInRectScreen(int x, int y, RECT r) => x >= r.Left && x <= r.Right && y >= r.Top && y <= r.Bottom;
        static Dictionary<string, string> GetVisibleWindows()
        {
            var windows = new Dictionary<string, string>();

            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                {
                    StringBuilder sb = new StringBuilder(256);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    string title = sb.ToString();

                    if (!string.IsNullOrEmpty(title))
                    {
                        // Получаем имя процесса по hwnd
                        uint processId;
                        GetWindowThreadProcessId(hWnd, out processId);
                        Process process = Process.GetProcessById((int)processId);
                        string processName = process.ProcessName;

                        if (!windows.ContainsKey(processName))
                        {
                            windows.Add(processName, title);
                        }
                    }
                }
                return true;
            }, IntPtr.Zero);

            return windows;
        }

        static Dictionary<string, string> ParseQueryString(string rawUrl)
        {
            var dict = new Dictionary<string, string>();
            int idx = rawUrl.IndexOf('?');
            if (idx < 0) return dict;
            var q = rawUrl.Substring(idx + 1);
            foreach (var pair in q.Split('&'))
            {
                var kv = pair.Split(new[] { '=' }, 2);
                if (kv.Length == 2)
                    dict[WebUtility.UrlDecode(kv[0])] = WebUtility.UrlDecode(kv[1]);
            }
            return dict;
        }

        static string CaptureWindow(IntPtr hwnd, RECT? highlightRect = null, List<RECT> allRects = null)
        {
            if (hwnd == IntPtr.Zero) return "";
            try
            {
                GetWindowRect(hwnd, out RECT rect);
                CurrentWindowRect = rect;
                int width = rect.Right - rect.Left;
                int height = rect.Bottom - rect.Top;
                if (width <= 0 || height <= 0) return "";

                using Bitmap bmp = new Bitmap(width, height, PixelFormat.Format32bppArgb);
                using (Graphics gfx = Graphics.FromImage(bmp))
                {
                    IntPtr hdc = gfx.GetHdc();
                    PrintWindow(hwnd, hdc, 0);
                    gfx.ReleaseHdc(hdc);

                    if (highlightRect.HasValue)
                    {
                        var r = highlightRect.Value;
                        using Pen pen = new Pen(Color.Red, 3);
                        gfx.DrawRectangle(pen, r.Left - rect.Left, r.Top - rect.Top, r.Right - r.Left, r.Bottom - r.Top);
                    }

                    if (ShowAllHighlights && allRects != null)
                    {
                        using Pen pen = new Pen(Color.FromArgb(128, Color.Lime), 2);
                        foreach (var r in allRects)
                        {
                            gfx.DrawRectangle(pen, r.Left - rect.Left, r.Top - rect.Top, r.Right - r.Left, r.Bottom - r.Top);
                        }
                    }
                }

                using MemoryStream ms = new MemoryStream();
                bmp.Save(ms, ImageFormat.Png);
                return "data:image/png;base64," + Convert.ToBase64String(ms.ToArray());
            }
            catch { return ""; }
        }

        static void CollectAllRects(XElement element, bool orderTopDown)
        {
            if (element == null) return;
            string id = element.Attribute("id")?.Value;
            foreach (var child in element.Elements("Element")) CollectAllRects(child, orderTopDown);
            if (!string.IsNullOrEmpty(id))
            {
                try
                {
                    RECT r = currentHelper.GetBoundingRectangleRECT(id);
                    if (orderTopDown) RectIdList.Add((r, id));
                    else RectIdList.Insert(0, (r, id));
                }
                catch { }
            }
        }

        static void ClickAt(int x, int y)
        {
            // Сохраняем текущее положение курсора
            Point originalPos = Cursor.Position;

            // Сохраняем текущее активное окно
            IntPtr foregroundHwnd = GetForegroundWindow();

            IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);
            if (hwnd != IntPtr.Zero) SetForegroundWindow(hwnd);
            Thread.Sleep(50);

            // Перемещаем курсор и кликаем
            Cursor.Position = new Point(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);

            Thread.Sleep(50);

            // Возвращаем курсор на исходное место
            Cursor.Position = originalPos;

            // Возвращаем фокус на предыдущее окно
            if (foregroundHwnd != IntPtr.Zero)
                SetForegroundWindow(foregroundHwnd);
        }

        static void DoubleClickAt(int x, int y)
        {
            ClickAt(x, y);
            Thread.Sleep(100);
            ClickAt(x, y);
        }

        static void SendKeysString(string text)
        {
            Point originalPos = Cursor.Position;
            IntPtr foregroundHwnd = GetForegroundWindow();

            IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);
            if (hwnd != IntPtr.Zero) SetForegroundWindow(hwnd);
            Thread.Sleep(50);

            foreach (char c in text)
            {
                short vkey = VkKeyScan(c);
                byte vk = (byte)(vkey & 0xFF);
                bool shift = (vkey & 0x0100) != 0;

                if (shift) keybd_event(0x10, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero); // SHIFT down
                keybd_event(vk, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                if (shift) keybd_event(0x10, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);   // SHIFT up

                Thread.Sleep(5);
            }

            Cursor.Position = originalPos;
            if (foregroundHwnd != IntPtr.Zero)
                SetForegroundWindow(foregroundHwnd);
        }


        static XElement FindByXPath(string xpath)
        {
            if (finderXml == null || string.IsNullOrWhiteSpace(xpath)) return null;

            try
            {
                // Преобразуем XML в строку с нижним регистром для тегов и атрибутов
                var lowerCaseXml = ToLowerCaseXml(finderXml);

                // Выполняем XPath запрос по преобразованному XML
                var nodes = lowerCaseXml.XPathSelectElements(xpath.ToLowerInvariant());

                // Находим первый элемент
                var value = nodes.FirstOrDefault();

                if (value != null) return value;

                Console.WriteLine("No element found.");
                return null;
            }
            catch (Exception ex)
            {
                // Для отладки можно раскомментировать:
                Console.WriteLine("XPath error: " + ex.Message);
                return null;
            }
        }

        static XElement ToLowerCaseXml(XElement element)
        {
            // Преобразуем имя текущего элемента в нижний регистр
            var lowerCaseElement = new XElement(
                element.Name.ToString().ToLowerInvariant(), // Имя элемента в нижнем регистре
                element.Attributes().Select(a => new XAttribute(a.Name.ToString().ToLowerInvariant(), a.Value)) // Атрибуты в нижнем регистре
            );

            // Рекурсивно обрабатываем дочерние элементы
            foreach (var child in element.Elements())
            {
                lowerCaseElement.Add(ToLowerCaseXml(child)); // Рекурсивный вызов для дочерних элементов
            }

            // Добавляем текстовое содержимое, не меняя его
            foreach (var node in element.Nodes())
            {
                // Добавляем все узлы (включая текстовые узлы) в новый элемент
                lowerCaseElement.Add(node);
            }

            return lowerCaseElement;
        }

        static void ProcessConsoleCommand(string commandLine)
        {
            string[] parts = commandLine.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return;
            string cmd = parts[0].ToLowerInvariant();
            string arg = parts.Length > 1 ? parts[1] : "";

            RECT rect;
            int centerX = -1;
            int centerY = -1;

            if (SelectedElementId == null)
            {
                Console.WriteLine("No element selected.");
            }
            else
            {
                rect = currentHelper.GetBoundingRectangleRECT(SelectedElementId);
                centerX = rect.Left + (rect.Right - rect.Left) / 2;
                centerY = rect.Top + (rect.Bottom - rect.Top) / 2;
            }

            switch (cmd)
            {
                case "selectprocess":
                    // Select the process by its name (process name as an argument)
                    if (string.IsNullOrEmpty(arg))
                    {
                        Console.WriteLine("Usage: selectprocess <process_name>");
                        return;
                    }

                    var windows = GetVisibleWindows();
                    if (windows.ContainsKey(arg))
                    {
                        CurrentProcess = arg;
                        CurrentTargetName = windows[CurrentProcess];
                        Console.WriteLine($"Process selected: {CurrentProcess}");

                        // Теперь получаем XML дерево для выбранного окна
                        IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);
                        XElement xml = currentHelper.DumpXml(hwnd);
                        finderXml = ConvertToXPathXml(xml);

                        // Очищаем старые данные о выделенных элементах
                        RectIdList.Clear();
                        CollectAllRects(xml.Elements("Element").FirstOrDefault(), false);

                        Console.WriteLine("XML tree loaded for selected process.");
                    }
                    else
                    {
                        Console.WriteLine($"Process '{arg}' not found.");
                    }
                    break;


                case "click":
                    if (arg.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase))
                    {
                        string xp = arg.Substring("xpath=".Length).Trim();

                        XElement node = FindByXPath(xp);
                        if (node == null)
                        {
                            Console.WriteLine("Element not found by xpath.");
                            return;
                        }

                        string id = node.Attribute("id")?.Value;
                        if (id == null)
                        {
                            Console.WriteLine("XPath node has no id.");
                            return;
                        }

                        RECT r = currentHelper.GetBoundingRectangleRECT(id);
                        int cx = r.Left + (r.Right - r.Left) / 2;
                        int cy = r.Top + (r.Bottom - r.Top) / 2;

                        ClickAt(cx, cy);
                        Console.WriteLine($"Click by xpath OK: {xp}");
                    }
                    else
                    {
                        ClickAt(centerX, centerY);
                        Console.WriteLine("Click performed.");
                    }
                    break;

                case "dblclick":
                    if (arg.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase))
                    {
                        string xp = arg.Substring("xpath=".Length).Trim();

                        XElement node = FindByXPath(xp);
                        if (node == null)
                        {
                            Console.WriteLine("Element not found by xpath.");
                            return;
                        }

                        string id = node.Attribute("id")?.Value;
                        if (id == null)
                        {
                            Console.WriteLine("XPath node has no id.");
                            return;
                        }

                        RECT r = currentHelper.GetBoundingRectangleRECT(id);
                        int cx = r.Left + (r.Right - r.Left) / 2;
                        int cy = r.Top + (r.Bottom - r.Top) / 2;

                        DoubleClickAt(cx, cy);
                        Console.WriteLine($"DoubleClick by xpath OK: {xp}");
                    }
                    else
                    {
                        DoubleClickAt(centerX, centerY);
                        Console.WriteLine("DoubleClick performed.");
                    }
                    break;

                case "sendkeys":
                    if (arg.StartsWith("xpath=", StringComparison.OrdinalIgnoreCase))
                    {
                        var parts2 = arg.Split(' ', 2);
                        if (parts2.Length < 2)
                        {
                            Console.WriteLine("Usage: sendkeys xpath=... text");
                            return;
                        }

                        string xp = parts2[0].Substring("xpath=".Length);
                        string text = parts2[1];

                        XElement node = FindByXPath(xp);
                        if (node == null)
                        {
                            Console.WriteLine("Element not found by xpath.");
                            return;
                        }

                        string id = node.Attribute("id")?.Value;
                        if (id == null)
                        {
                            Console.WriteLine("XPath node has no id.");
                            return;
                        }

                        RECT r = currentHelper.GetBoundingRectangleRECT(id);
                        int cx = r.Left + (r.Right - r.Left) / 2;
                        int cy = r.Top + (r.Bottom - r.Top) / 2;

                        ClickAt(cx, cy);
                        Thread.Sleep(120);
                        SendKeysString(text);

                        Console.WriteLine($"sendkeys by xpath OK: {xp}");
                    }
                    else
                    {
                        SendKeysString(arg);
                        Console.WriteLine($"Sent keys: {arg}");
                    }
                    break;
                case "getprop":
                    if (string.IsNullOrEmpty(arg) || !arg.Contains("xpath="))
                    {
                        Console.WriteLine("Usage: getprop xpath=<xpath_expression> <property_name>");
                        return;
                    }

                    // Разделяем аргументы
                    string[] args = arg.Split(new[] { ' ' }, 2);
                    if (args.Length < 2)
                    {
                        Console.WriteLine("Usage: getprop xpath=<xpath_expression> <property_name>");
                        return;
                    }

                    string xpath = args[0].Substring("xpath=".Length).Trim();
                    string propertyName = args[1].Trim();

                    // Находим элемент по XPath
                    XElement element = FindByXPath(xpath);
                    if (element == null)
                    {
                        Console.WriteLine("Element not found by xpath.");
                        return;
                    }

                    // Получаем значение указанного свойства
                    string propertyValue = GetSpecificPropertyFromXml(element, propertyName);
                    Console.WriteLine(propertyValue);
                    break;

                default:
                    Console.WriteLine("Unknown command.");
                    break;
            }
        }

        static string GetSpecificPropertyFromXml(XElement element, string propertyName)
        {
            // Ищем свойство в атрибутах элемента по имени
            var attribute = element.Attribute(propertyName);
            if (attribute != null)
            {
                return attribute.Value;
            }
            else
            {
                return $"Property '{propertyName}' not found.";
            }
        }


        private static XElement ConvertToXPathXml(XElement input)
        {
            if (input == null) return null;

            // Берём название тега = controlType
            string tagName = input.Attribute("controlType")?.Value;
            if (string.IsNullOrEmpty(tagName))
                tagName = "Unknown";

            XElement output = new XElement(tagName);

            // Копируем все атрибуты
            foreach (var attr in input.Attributes())
            {
                output.SetAttributeValue(attr.Name, attr.Value);
            }

            // Рекурсивно обрабатываем детей
            foreach (var child in input.Elements())
            {
                output.Add(ConvertToXPathXml(child));
            }

            return output;
        }

        static void Main()
        {
            // HTTP сервер (без изменений)
            new Thread(() =>
            {
                HttpListener listener = new HttpListener();
                listener.Prefixes.Add("http://localhost:5050/");
                listener.Start();
                Console.WriteLine("Inspector running at http://localhost:5050");

                while (true)
                {
                    try
                    {
                        Thread.Sleep(100);

                        var ctx = listener.GetContext();
                        var req = ctx.Request;
                        var resp = ctx.Response;
                        var qs = ParseQueryString(req.Url.ToString());

                        var windows = GetVisibleWindows();
                        if (string.IsNullOrEmpty(CurrentTargetName) && windows.Count > 0)
                            CurrentTargetName = windows.Values.First();
                        if (string.IsNullOrEmpty(CurrentProcess) && windows.Count > 0)
                            windows.Keys.First();

                        if (qs.TryGetValue("name", out string target)) CurrentProcess = target;
                        CurrentTargetName = windows[CurrentProcess];
                        if (qs.TryGetValue("helper", out string helper))
                            currentHelper = (helper == "MSAA") ? (IHelper)msaaHelper : (IHelper)uiaHelper;

                        IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);

                        if (req.Url.AbsolutePath.Equals("/inspect", StringComparison.OrdinalIgnoreCase))
                        {
                            XElement xml = currentHelper.DumpXml(hwnd);
                            finderXml = ConvertToXPathXml(xml);
                            RectIdList.Clear();
                            CollectAllRects(xml.Elements("Element").FirstOrDefault(), false);
                            string treeHtml = "<ul class='tree'>" + PageRenderer.XmlTreeToHtml(xml.Elements("Element").FirstOrDefault(), SelectedElementId) + "</ul>";
                            string screenshotSrc = CaptureWindow(hwnd, null, RectIdList.Select(t => t.Rect).ToList());
                            string html = PageRenderer.RenderPage(windows.Keys.ToList(), windows.FirstOrDefault(x => x.Value == CurrentTargetName).Key, currentHelper, ShowAllHighlights, treeHtml, screenshotSrc);
                            byte[] buf = Encoding.UTF8.GetBytes(html);
                            resp.ContentType = "text/html; charset=utf-8";
                            resp.OutputStream.Write(buf, 0, buf.Length);
                            resp.Close();
                        }
                        else if (req.Url.AbsolutePath.Equals("/props", StringComparison.OrdinalIgnoreCase))
                        {
                            if (qs.TryGetValue("id", out string id))
                            {
                                SelectedElementId = id;
                                string propsHtml = currentHelper.GetElementProperties(id);
                                RECT highlightRect = currentHelper.GetBoundingRectangleRECT(id);
                                string screenshotSrc = CaptureWindow(hwnd, highlightRect, RectIdList.Select(t => t.Rect).ToList());
                                var json = $"{{\"html\":{System.Text.Json.JsonSerializer.Serialize(propsHtml)},\"screenshot\":\"{screenshotSrc}\"}}";
                                byte[] buf2 = Encoding.UTF8.GetBytes(json);
                                resp.ContentType = "application/json";
                                resp.OutputStream.Write(buf2, 0, buf2.Length);
                                resp.Close();
                            }
                            else { resp.StatusCode = 400; resp.Close(); }
                        }
                        else if (req.Url.AbsolutePath.Equals("/screenshot", StringComparison.OrdinalIgnoreCase))
                        {
                            ShowAllHighlights = qs.TryGetValue("all", out string allVal) && allVal == "1";
                            string screenshotSrc = CaptureWindow(hwnd, null, RectIdList.Select(t => t.Rect).ToList());
                            byte[] buf = Encoding.UTF8.GetBytes(screenshotSrc);
                            resp.ContentType = "text/plain";
                            resp.OutputStream.Write(buf, 0, buf.Length);
                            resp.Close();
                        }
                        else if (req.Url.AbsolutePath.Equals("/click", StringComparison.OrdinalIgnoreCase))
                        {
                            if (!qs.TryGetValue("x", out string sx) || !qs.TryGetValue("y", out string sy))
                            {
                                resp.StatusCode = 400;
                                resp.Close();
                            }
                            else
                            {
                                int x = int.Parse(sx);
                                int y = int.Parse(sy);
                                GetWindowRect(hwnd, out RECT winRect);
                                CurrentWindowRect = winRect;

                                string foundId = null;
                                RECT? highlight = null;
                                for (int i = RectIdList.Count - 1; i >= 0; i--)
                                {
                                    var entry = RectIdList[i];
                                    if (PointInRectScreen(x + CurrentWindowRect.Left, y + CurrentWindowRect.Top, entry.Rect))
                                    {
                                        foundId = entry.Id;
                                        highlight = entry.Rect;
                                        SelectedElementId = foundId;
                                        break;
                                    }
                                }

                                string screenshotSrc = CaptureWindow(hwnd, highlight, RectIdList.Select(t => t.Rect).ToList());
                                string json = $"{{\"id\":{(foundId == null ? "null" : System.Text.Json.JsonSerializer.Serialize(foundId))},\"screenshot\":\"{screenshotSrc}\"}}";
                                byte[] buf = Encoding.UTF8.GetBytes(json);
                                resp.ContentType = "application/json";
                                resp.OutputStream.Write(buf, 0, buf.Length);
                                resp.Close();
                            }
                        }
                        else if (req.Url.AbsolutePath.Equals("/console", StringComparison.OrdinalIgnoreCase))
                        {
                            if (qs.TryGetValue("cmd", out string command))
                            {
                                if (!string.IsNullOrEmpty(command))
                                {
                                    ProcessConsoleCommand(command);
                                }
                                resp.StatusCode = 200;
                                byte[] buf = Encoding.UTF8.GetBytes("OK");
                                resp.OutputStream.Write(buf, 0, buf.Length);
                                resp.Close();
                            }
                            else
                            {
                                resp.StatusCode = 400;
                                resp.Close();
                            }
                        }

                        else
                        {
                            if (req.Url.AbsolutePath == "/") { resp.Redirect("/inspect"); resp.Close(); }
                            else { resp.StatusCode = 404; resp.Close(); }
                        }
                    }
                    catch (Exception ex) { Console.WriteLine("Error: " + ex); }
                }
            }).Start();

            // Консольный ввод в отдельном потоке
            new Thread(() =>
            {
                Console.WriteLine("Console commands: click, dblclick, sendkeys <text>");
                while (true)
                {
                    Console.Write("> ");
                    string line = Console.ReadLine();
                    if (!string.IsNullOrEmpty(line)) ProcessConsoleCommand(line);
                }
            }).Start();
        }
    }
}
