using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Xml.Linq;
using Accessibility;
using static InteractiveInspector.Program;

namespace InteractiveInspector
{
    public class MsaaHelper : IHelper
    {
        public Dictionary<string, IAccessible> ElementMap { get; private set; } = new();
        private int NodeCounter = 0;

        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("oleacc.dll")]
        private static extern int AccessibleObjectFromWindow(IntPtr hwnd, uint id, ref Guid iid, [Out, MarshalAs(UnmanagedType.Interface)] out IAccessible acc);

        private const uint OBJID_CLIENT = 0xFFFFFFFC;
        private static Guid IID_IAccessible = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");

        public IntPtr ResolveWindow(string name = null)
        {
            if (!string.IsNullOrEmpty(name))
            {
                var hwnd = FindWindow(null, name);
                if (hwnd != IntPtr.Zero) return hwnd;
            }
            return GetForegroundWindow();
        }

        private string GetRoleName(int role)
        {
            return role switch
            {
                0x0 => "None",
                0x1 => "TitleBar",
                0x2 => "MenuBar",
                0x3 => "Menu",
                0x4 => "PopupMenu",
                0x5 => "MenuItem",
                0x6 => "ToolTip",
                0x7 => "Application",
                0x8 => "Document",
                0x9 => "Pane",
                0xA => "Chart",
                0xB => "Dialog",
                0xC => "Border",
                0xD => "Grouping",
                0xE => "Separator",
                0xF => "ToolBar",
                0x10 => "StatusBar",
                0x11 => "Table",
                0x12 => "ColumnHeader",
                0x13 => "RowHeader",
                0x14 => "Column",
                0x15 => "Row",
                0x16 => "Cell",
                0x17 => "Link",
                0x18 => "HelpBalloon",
                0x19 => "Character",
                0x1A => "List",
                0x1B => "ListItem",
                0x1C => "Outline",
                0x1D => "OutlineItem",
                0x1E => "PageTab",
                0x1F => "PageTabList",
                0x20 => "PropertyPage",
                0x21 => "Indicator",
                0x22 => "Graphic",
                0x23 => "StaticText",
                0x24 => "Text",
                0x25 => "PushButton",
                0x26 => "CheckButton",
                0x27 => "RadioButton",
                0x28 => "ComboBox",
                0x29 => "DropList",
                0x2A => "ProgressBar",
                0x2B => "Dial",
                0x2C => "HotkeyField",
                0x2D => "Slider",
                0x2E => "SpinButton",
                0x2F => "Canvas",
                0x30 => "Animation",
                0x31 => "Equation",
                0x32 => "ButtonDropDown",
                0x33 => "ButtonMenu",
                0x34 => "ButtonDropDownGrid",
                0x35 => "WhiteSpace",
                0x36 => "PageTabPage",
                0x37 => "Clock",
                0x38 => "Splitter",
                _ => "Unknown",
            };
        }


        private XElement DumpElement(IAccessible element, string parentId = "")
        {
            if (element == null) return null;

            NodeCounter++;
            string currentId = parentId + "_" + NodeCounter;

            string name = "", controlType = "", value = "";
            int childCount = 0;
            try { name = element.get_accName(0); } catch { }
            try
            {
                int r = element.get_accRole(0) is int roleInt ? roleInt : 0;
                controlType = GetRoleName(r); // используем Role как ControlType
            }
            catch { }
            try { value = element.get_accValue(0); } catch { }
            try { childCount = element.accChildCount; } catch { }

            ElementMap[currentId] = element;

            XElement el = new XElement("Element",
                new XAttribute("id", currentId),
                new XAttribute("name", name ?? ""),
                new XAttribute("controlType", controlType ?? ""), // здесь ControlType
                new XAttribute("value", value ?? "")
            );

            try
            {
                for (int i = 1; i <= childCount; i++)
                {
                    object child = element.get_accChild(i);
                    if (child is IAccessible childAcc)
                        el.Add(DumpElement(childAcc, currentId));
                }
            }
            catch { }

            return el;
        }

        public XElement DumpXml(IntPtr hwnd)
        {
            NodeCounter = 0;
            ElementMap.Clear();

            if (AccessibleObjectFromWindow(hwnd, OBJID_CLIENT, ref IID_IAccessible, out var root) != 0 || root == null)
                return new XElement("Root");

            XElement rootXml = new XElement("Root");
            rootXml.Add(DumpElement(root));
            return rootXml;
        }

        public string GetElementProperties(string elementId)
        {
            if (!ElementMap.TryGetValue(elementId, out var element)) return "";

            var sb = new StringBuilder();
            try { sb.AppendLine($"<b>Name:</b> {element.get_accName(0)}<br>"); } catch { }
            try
            {
                int r = element.get_accRole(0) is int roleInt ? roleInt : 0;
                sb.AppendLine($"<b>ControlType:</b> {GetRoleName(r)}<br>"); // Role как ControlType
            }
            catch { }
            try { sb.AppendLine($"<b>Value:</b> {element.get_accValue(0)}<br>"); } catch { }
            try { sb.AppendLine($"<b>State:</b> {element.get_accState(0)}<br>"); } catch { }

            try
            {
                element.accLocation(out int left, out int top, out int width, out int height, 0);
                sb.AppendLine($"<b>BoundingRectangle:</b> X={left},Y={top},Width={width},Height={height}<br>");
            }
            catch { }

            return sb.ToString();
        }

        public (int left, int top, int width, int height) GetBoundingRectangle(string elementId)
        {
            if (!ElementMap.TryGetValue(elementId, out var element)) return (0, 0, 0, 0);
            try
            {
                element.accLocation(out int left, out int top, out int width, out int height, 0);
                return (left, top, width, height);
            }
            catch { return (0, 0, 0, 0); }
        }

        public Program.RECT GetBoundingRectangleRECT(string elementId)
        {
            var rect = GetBoundingRectangle(elementId);
            return new Program.RECT
            {
                Left = rect.left,
                Top = rect.top,
                Right = rect.left + rect.width,
                Bottom = rect.top + rect.height
            };
        }
    }
}
