using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Windows;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using SystemPoint = System.Drawing.Point;

namespace RevitTest.Schedules.Utilities
{
    class Win32Utilities
    {
        public static SystemPoint Revit2Screen(UIDocument uidoc, UIView uiView, XYZ point)
        {
            try
            {
                var rect = uiView.GetWindowRectangle();

                Transform inverse = uidoc.ActiveView.CropBox.Transform.Inverse;
                IList<XYZ> zoomCorners = uiView.GetZoomCorners();
                XYZ xyz = inverse.OfPoint(point);
                XYZ xyz2 = inverse.OfPoint(zoomCorners[0]);
                XYZ xyz3 = inverse.OfPoint(zoomCorners[1]);
                double num5 = (double)(rect.Right - rect.Left) / (xyz3.X - xyz2.X);
                int x = rect.Left + (int)((xyz.X - xyz2.X) * num5);
                int y = rect.Top + (int)((xyz3.Y - xyz.Y) * num5);
                return new SystemPoint(x, y);
            }
            catch
            {
            }
            return SystemPoint.Empty;
        }

        public static void PressEscKeyInRevit(bool pressEsc2Times = false)
        {
            IWin32Window _revit_window = new JtWindowHandle(ComponentManager.ApplicationWindow);
            Press.PostMessage(_revit_window.Handle, (uint)Press.KEYBOARD_MSG.WM_KEYDOWN, (uint)Keys.Escape, 0);
            if (pressEsc2Times)
            {
                Press.PostMessage(_revit_window.Handle,
                (uint)Press.KEYBOARD_MSG.WM_KEYDOWN,
                (uint)Keys.Escape, 0);
            }
        }

        public static void PressMouseLeftButton()
        {
            //IWin32Window _revit_window = new JtWindowHandle(ComponentManager.ApplicationWindow);
            //Press.PostMessage(_revit_window.Handle, (uint)Press.KEYBOARD_MSG.WM_KEYDOWN, (uint)Keys.LButton, 0);

            uint X = (uint)Cursor.Position.X;
            uint Y = (uint)Cursor.Position.Y;
            Press.mouse_event((uint)Press.MOUSEEVENTF.LEFTDOWN, X, Y, 0, 0);
        }
    }

    public class JtWindowHandle : IWin32Window
    {
        IntPtr _hwnd;

        public JtWindowHandle(IntPtr h)
        {
            Debug.Assert(IntPtr.Zero != h,
              "expected non-null window handle");

            _hwnd = h;
        }

        public IntPtr Handle
        {
            get
            {
                return _hwnd;
            }
        }
    }
    public class Press
    {
        [DllImport("USER32.DLL", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);

        public enum MOUSEEVENTF : uint
        {
            LEFTDOWN = 0x02,
            LEFTUP = 0x04,
            RIGHTDOWN = 0x08,
            RIGHTUP = 0x10
        }
        [DllImport("USER32.DLL")]
        public static extern bool PostMessage(
          IntPtr hWnd, uint msg, uint wParam, uint lParam);

        [DllImport("user32.dll")]
        static extern uint MapVirtualKey(
          uint uCode, uint uMapType);

        public enum WH_KEYBOARD_LPARAM : uint
        {
            KEYDOWN = 0x00000001,
            KEYUP = 0xC0000001
        }

        public enum KEYBOARD_MSG : uint
        {
            WM_KEYDOWN = 0x100,
            WM_KEYUP = 0x101
        }

        public enum MVK_MAP_TYPE : uint
        {
            VKEY_TO_SCANCODE = 0,
            SCANCODE_TO_VKEY = 1,
            VKEY_TO_CHAR = 2,
            SCANCODE_TO_LR_VKEY = 3
        }

        /// <summary>
        /// Post one single keystroke.
        /// </summary>
        public static void OneKey(IntPtr handle, char letter)
        {
            uint scanCode = MapVirtualKey(letter,
              (uint)MVK_MAP_TYPE.VKEY_TO_SCANCODE);

            uint keyDownCode = (uint)
              WH_KEYBOARD_LPARAM.KEYDOWN
              | (scanCode << 16);

            uint keyUpCode = (uint)
              WH_KEYBOARD_LPARAM.KEYUP
              | (scanCode << 16);

            PostMessage(handle,
              (uint)KEYBOARD_MSG.WM_KEYDOWN,
              letter, keyDownCode);

            PostMessage(handle,
              (uint)KEYBOARD_MSG.WM_KEYUP,
              letter, keyUpCode);
        }

        /// <summary>
        /// Post a sequence of keystrokes.
        /// </summary>
        public static void Keys(string command)
        {
            IntPtr revitHandle = System.Diagnostics.Process
              .GetCurrentProcess().MainWindowHandle;

            foreach (char letter in command)
            {
                OneKey(revitHandle, letter);
            }
        }
    }

}
