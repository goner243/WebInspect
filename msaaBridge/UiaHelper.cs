using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Linq;
using UIA = UIAutomationClient;
using System.Runtime.InteropServices;
using System.ComponentModel.Design;
using static InteractiveInspector.Program;

namespace InteractiveInspector
{
    public class UiaHelper : IHelper
    {
        public Dictionary<string, UIA.IUIAutomationElement> ElementMap { get; private set; } = new();
        private int NodeCounter = 0;

        [DllImport("user32.dll")]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        public IntPtr ResolveWindow(string name = null)
        {
            if (!string.IsNullOrEmpty(name))
            {
                var hwnd = FindWindow(null, name);
                if (hwnd != IntPtr.Zero) return hwnd;
            }
            return GetForegroundWindow();
        }

        private string GetControlTypeName(int controlType)
        {
            return controlType switch
            {
                50000 => "Button",
                50001 => "Calendar",
                50002 => "CheckBox",
                50003 => "ComboBox",
                50004 => "Edit",
                50005 => "Hyperlink",
                50006 => "Image",
                50007 => "ListItem",
                50008 => "List",
                50009 => "Menu",
                50010 => "MenuBar",
                50011 => "MenuItem",
                50012 => "ProgressBar",
                50013 => "RadioButton",
                50014 => "ScrollBar",
                50015 => "Slider",
                50016 => "Spinner",
                50017 => "StatusBar",
                50018 => "Tab",
                50019 => "TabItem",
                50020 => "Text",
                50021 => "ToolBar",
                50022 => "ToolTip",
                50023 => "Tree",
                50024 => "TreeItem",
                50025 => "Custom",
                50026 => "Group",
                50027 => "Custom",
                50028 => "Thumb",
                50029 => "DataGrid",
                50030 => "DataItem",
                50031 => "Document",
                50032 => "Window",
                50033 => "Pane",
                50034 => "Header",
                50035 => "HeaderItem",
                50036 => "Table",
                50037 => "TitleBar",
                50038 => "Separator",
                50039 => "SemanticZoom",
                50040 => "AppBar",
                _ => "Unknown",
            };
        }


        private XElement DumpElement(UIA.IUIAutomationElement element, string parentId = "")
        {
            if (element == null) return null;

            NodeCounter++;
            string currentId = parentId + "_" + NodeCounter;

            string name = "", typeName = "", className = "";
            try { name = element.CurrentName; } catch { }
            try { typeName = GetControlTypeName(element.CurrentControlType); } catch { }
            try { className = element.CurrentClassName; } catch { }
            if (string.IsNullOrEmpty(className))
            {
                try { className = element.CurrentAutomationId; } catch { }
            }

            ElementMap[currentId] = element;

            XElement el = new XElement("Element",
                new XAttribute("id", currentId),
                new XAttribute("name", name ?? ""),
                new XAttribute("controlType", typeName ?? ""),
                new XAttribute("className", className ?? "")
            );

            try
            {
                UIA.CUIAutomation uia = new UIA.CUIAutomation();
                var cond = uia.CreateTrueCondition();
                var children = element.FindAll(UIA.TreeScope.TreeScope_Children, cond);
                for (int i = 0; i < children.Length; i++)
                {
                    var child = children.GetElement(i);
                    el.Add(DumpElement(child, currentId));
                }
            }
            catch { }

            return el;
        }

        public XElement DumpXml(IntPtr hwnd)
        {
            UIA.CUIAutomation uia = new UIA.CUIAutomation();
            UIA.IUIAutomationElement rootElement = uia.ElementFromHandle(hwnd);
            NodeCounter = 0;
            ElementMap.Clear();
            XElement root = new XElement("Root");
            root.Add(DumpElement(rootElement));
            return root;
        }

        public string GetElementProperties(string elementId)
        {
            if (!ElementMap.TryGetValue(elementId, out var element)) return "";

            var sb = new StringBuilder();
            try { sb.AppendLine($"<b>Name:</b> {element.CurrentName}<br>"); } catch { }
            try { sb.AppendLine($"<b>AutomationId:</b> {element.CurrentAutomationId}<br>"); } catch { }
            try { sb.AppendLine($"<b>ControlType:</b> {GetControlTypeName(element.CurrentControlType)}<br>"); } catch { }
            try
            {
                string className = element.CurrentClassName;
                if (string.IsNullOrEmpty(className)) className = element.CurrentAutomationId;
                sb.AppendLine($"<b>ClassName:</b> {className}<br>");
            }
            catch { }
            try { sb.AppendLine($"<b>FrameworkId:</b> {element.CurrentFrameworkId}<br>"); } catch { }
            try { sb.AppendLine($"<b>IsEnabled:</b> {element.CurrentIsEnabled}<br>"); } catch { }
            try { sb.AppendLine($"<b>HasKeyboardFocus:</b> {element.CurrentHasKeyboardFocus}<br>"); } catch { }
            try
            {
                var r = element.CurrentBoundingRectangle;
                sb.AppendLine($"<b>BoundingRectangle:</b> X={r.left},Y={r.top},Width={r.right - r.left},Height={r.bottom - r.top}<br>");
            }
            catch { }

            return sb.ToString();
        }

        public UIA.tagRECT GetBoundingRectangle(string elementId)
        {
            if (!ElementMap.TryGetValue(elementId, out var element)) return new UIA.tagRECT();
            try { return element.CurrentBoundingRectangle; } catch { return new UIA.tagRECT(); }
        }

        public Program.RECT GetBoundingRectangleRECT(string elementId)
        {
            var r = GetBoundingRectangle(elementId); // уже реализованный метод
            return new Program.RECT
            {
                Left = r.left,
                Top = r.top,
                Right = r.right,
                Bottom = r.bottom
            };
        }

    }
}
