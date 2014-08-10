using System;
using System.Collections.Generic;

using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;

namespace WinAPI
{
    public static class WindowAttrib
    {
        [StructLayout(LayoutKind.Sequential)]
        struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }

            public static implicit operator Point(POINT p)
            {
                return new Point(p.X, p.Y);
            }

            public static implicit operator POINT(Point p)
            {
                return new POINT(p.X, p.Y);
            }
        }

        #region User32.dll

        const int WM_GETTEXT = 13;
        const int GCL_HICONSM = -34;
        const int GCL_HICON = -14;
        const int ICON_SMALL = 0;
        const int ICON_BIG = 1;
        const int ICON_SMALL2 = 2;
        const int WM_GETICON = 0x7F;

        [DllImport("user32.dll")]
        static extern int GetWindowText(int hWnd, StringBuilder text, int count);
        [DllImport("user32.dll")]
        static extern uint RealGetWindowClass(int win, StringBuilder pszType, uint cchType);
        [DllImport("user32.dll", SetLastError = true)]
        static extern int GetWindowThreadProcessId(int hWnd, ref int processId);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool GetWindowRect(int win, out RECT rect);
        [DllImport("user32.dll")]
        static extern bool GetClientRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        public static extern int GetForegroundWindow();
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(int win);
        [DllImport("user32.dll")]
        static extern bool ScreenToClient(int win, ref POINT point);
        [DllImport("user32.dll")]
        static extern bool ClientToScreen(int win, ref POINT point);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = false)]
        static extern IntPtr SendMessage(int hWnd, int Msg, int wParam, StringBuilder lParam);
        [DllImport("user32.dll", EntryPoint = "GetClassLong")]
        static extern uint GetClassLongPtr32(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll", EntryPoint = "GetClassLongPtr")]
        static extern IntPtr GetClassLongPtr64(IntPtr hWnd, int nIndex);

        #endregion

        /// <summary>Повертає заголовок вікна</summary>
        /// <param name="hWnd">Адреса вікна</param>
        public static string GetWindowText(int hWnd)
        {
            StringBuilder buff = new StringBuilder(256);
            GetWindowText(hWnd, buff, 256);

            return buff.ToString();
        }

        public static string GetWindowCaption(int hWnd)
        {
            StringBuilder buff = new StringBuilder(256);
            //Get the text from the active window into the stringbuilder
            SendMessage(hWnd, WM_GETTEXT, 256, buff);
            return buff.ToString();
        }

        /// <summary>Повертає заголовок активного вікна</summary>
        public static string GetForegroundWindowText()
        {
            int win = GetForegroundWindow();
            return GetWindowText(win);
        }

        /// <summary>Повертає назву класу вікна</summary>
        /// <param name="win">Адреса вікна</param>
        public static string GetWindowClass(int win)
        {
            if (win == 0) return "";

            StringBuilder title = new StringBuilder(512);
            RealGetWindowClass(win, title, 512);

            return title.ToString().Trim();
        }

        /// <summary>Повертає ідентифікатор процесу, до якого належить дане вікно</summary>
        /// <param name="win">Адреса вікна</param>
        public static int GetProcessId(int win)
        {
            int id = 0;

            GetWindowThreadProcessId(win, ref id);

            return id;
        }
        
        /// <summary>Показує чи існує задане вікно</summary>
        /// <param name="win">Адреса вікна</param>
        public static bool Exists(int win)
        {
            return GetProcessId(win) != 0;
        }

        /// <summary>Повертає прямокутник заданого вікна</summary>
        /// <param name="win">Адреса вікна</param>
        public static Rectangle GetWindowRect(int win)
        {
            RECT rect = new RECT();

            GetWindowRect(win, out rect);

            Rectangle result = new Rectangle();

            result.X = rect.Left;
            result.Y = rect.Top;
            result.Width = rect.Right - rect.Left;
            result.Height = rect.Bottom - rect.Top;

            return result;
        }

        /// <summary>Повертає прямокутник клієнтської області заданого вікна</summary>
        /// <param name="win">Адреса вікна</param>
        public static Rectangle GetClientRect(int win)
        {
            RECT rect = new RECT();

            GetClientRect(new IntPtr(win), out rect);

            Rectangle result = new Rectangle();
            POINT p = new POINT(0, 0);
            ClientToScreen(win, ref p);

            result.Location = p;
            result.Width = rect.Right - rect.Left;
            result.Height = rect.Bottom - rect.Top;

            return result;
        }

        public static IntPtr GetClassLongPtr(IntPtr hWnd, int nIndex)
        {
            if (IntPtr.Size > 4) return GetClassLongPtr64(hWnd, nIndex);
            else return new IntPtr(GetClassLongPtr32(hWnd, nIndex));
        }

        public static Icon GetAppIcon(IntPtr hwnd)
        {
            IntPtr iconHandle = SendMessage(hwnd.ToInt32(), WM_GETICON, ICON_SMALL2, null);
            if (iconHandle == IntPtr.Zero)
                iconHandle = SendMessage(hwnd.ToInt32(), WM_GETICON, ICON_SMALL, null);
            if (iconHandle == IntPtr.Zero)
                iconHandle = SendMessage(hwnd.ToInt32(), WM_GETICON, ICON_BIG, null);
            if (iconHandle == IntPtr.Zero)
                iconHandle = GetClassLongPtr(hwnd, GCL_HICON);
            if (iconHandle == IntPtr.Zero)
                iconHandle = GetClassLongPtr(hwnd, GCL_HICONSM);

            if (iconHandle == IntPtr.Zero)
                return null;

            Icon icn = Icon.FromHandle(iconHandle);

            return icn;
        }

    }
}
