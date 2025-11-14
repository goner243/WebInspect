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

namespace InteractiveInspector
{
    public class Program
    {
        static string CurrentTargetName = null;
        static string SelectedElementId = null;
        static UiaHelper uiaHelper = new UiaHelper();
        static MsaaHelper msaaHelper = new MsaaHelper();
        static IHelper currentHelper = uiaHelper;
        static RECT CurrentWindowRect;
        static List<(RECT Rect, string Id)> RectIdList = new List<(RECT, string)>();
        static bool ShowAllHighlights = false;

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT { public int Left, Top, Right, Bottom; }

        public interface IHelper
        {
            IntPtr ResolveWindow(string windowName);
            XElement DumpXml(IntPtr hwnd);
            string GetElementProperties(string elementId);
            Program.RECT GetBoundingRectangleRECT(string elementId);
        }

        #region Win32 imports
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);
        [DllImport("user32.dll")] static extern bool PrintWindow(IntPtr hwnd, IntPtr hDC, uint nFlags);
        [DllImport("user32.dll")] static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);
        [DllImport("user32.dll")] static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);
        [DllImport("user32.dll")] static extern bool IsWindowVisible(IntPtr hWnd);
        [DllImport("user32.dll")] static extern bool SetForegroundWindow(IntPtr hWnd);
        [DllImport("user32.dll")] static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint dwData, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")] static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, UIntPtr dwExtraInfo);
        [DllImport("user32.dll")] static extern short VkKeyScan(char ch);

        const uint MOUSEEVENTF_MOVE = 0x0001;
        const uint MOUSEEVENTF_LEFTDOWN = 0x0002;
        const uint MOUSEEVENTF_LEFTUP = 0x0004;
        const uint KEYEVENTF_KEYDOWN = 0x0000;
        const uint KEYEVENTF_KEYUP = 0x0002;

        delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);
        #endregion

        static bool PointInRectScreen(int x, int y, RECT r) => x >= r.Left && x <= r.Right && y >= r.Top && y <= r.Bottom;

        static List<string> GetVisibleWindows()
        {
            var list = new List<string>();
            EnumWindows((hWnd, lParam) =>
            {
                if (IsWindowVisible(hWnd))
                {
                    StringBuilder sb = new StringBuilder(256);
                    GetWindowText(hWnd, sb, sb.Capacity);
                    string title = sb.ToString();
                    if (!string.IsNullOrEmpty(title)) list.Add(title);
                }
                return true;
            }, IntPtr.Zero);
            return list;
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
                            gfx.DrawRectangle(pen, r.Left - rect.Left, r.Top - rect.Top, r.Right - r.Left, r.Bottom - rect.Top);
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

        #region Input simulation helpers
        static void ClickAt(int x, int y)
        {
            IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);
            if (hwnd != IntPtr.Zero) SetForegroundWindow(hwnd);
            Thread.Sleep(50);
            Cursor.Position = new Point(x, y);
            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        static void DoubleClickAt(int x, int y)
        {
            ClickAt(x, y);
            Thread.Sleep(100);
            ClickAt(x, y);
        }

        static void SendKeysString(string text)
        {
            IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);
            if (hwnd != IntPtr.Zero) SetForegroundWindow(hwnd);
            Thread.Sleep(50); // ждем фокус

            foreach (char c in text)
            {
                short vkey = VkKeyScan(c);
                byte vk = (byte)(vkey & 0xFF);
                bool shift = (vkey & 0x0100) != 0;

                if (shift) keybd_event(0x10, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero); // SHIFT down
                keybd_event(vk, 0, KEYEVENTF_KEYDOWN, UIntPtr.Zero);
                keybd_event(vk, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);
                if (shift) keybd_event(0x10, 0, KEYEVENTF_KEYUP, UIntPtr.Zero);   // SHIFT up

                Thread.Sleep(5); // небольшая задержка между символами
            }
        }

        static void ProcessConsoleCommand(string commandLine)
        {
            string[] parts = commandLine.Split(' ', 2, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length == 0) return;
            string cmd = parts[0].ToLowerInvariant();
            string arg = parts.Length > 1 ? parts[1] : "";

            if (SelectedElementId == null)
            {
                Console.WriteLine("No element selected.");
                return;
            }

            RECT rect = currentHelper.GetBoundingRectangleRECT(SelectedElementId);
            int centerX = rect.Left + (rect.Right - rect.Left) / 2;
            int centerY = rect.Top + (rect.Bottom - rect.Top) / 2;

            switch (cmd)
            {
                case "click":
                    ClickAt(centerX, centerY);
                    Console.WriteLine("Click performed.");
                    break;
                case "dblclick":
                    DoubleClickAt(centerX, centerY);
                    Console.WriteLine("Double click performed.");
                    break;
                case "sendkeys":
                    SendKeysString(arg);
                    Console.WriteLine($"Sent keys: {arg}");
                    break;
                default:
                    Console.WriteLine("Unknown command.");
                    break;
            }
        }
        #endregion

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
                        var ctx = listener.GetContext();
                        var req = ctx.Request;
                        var resp = ctx.Response;
                        var qs = ParseQueryString(req.Url.ToString());

                        var windows = GetVisibleWindows();
                        if (string.IsNullOrEmpty(CurrentTargetName) && windows.Count > 0)
                            CurrentTargetName = windows[0];

                        if (qs.TryGetValue("name", out string target)) CurrentTargetName = target;
                        if (qs.TryGetValue("helper", out string helper))
                            currentHelper = (helper == "MSAA") ? (IHelper)msaaHelper : (IHelper)uiaHelper;

                        IntPtr hwnd = currentHelper.ResolveWindow(CurrentTargetName);

                        if (req.Url.AbsolutePath.Equals("/inspect", StringComparison.OrdinalIgnoreCase))
                        {
                            XElement xml = currentHelper.DumpXml(hwnd);
                            RectIdList.Clear();
                            CollectAllRects(xml.Elements("Element").FirstOrDefault(), false);
                            string treeHtml = "<ul class='tree'>" + PageRenderer.XmlTreeToHtml(xml.Elements("Element").FirstOrDefault(), SelectedElementId) + "</ul>";
                            string screenshotSrc = CaptureWindow(hwnd, null, RectIdList.Select(t => t.Rect).ToList());
                            string html = PageRenderer.RenderPage(windows, CurrentTargetName, currentHelper, ShowAllHighlights, treeHtml, screenshotSrc);
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
