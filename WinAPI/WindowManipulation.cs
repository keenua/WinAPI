using System;
using System.Collections.Generic;
//using System.Data.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using System.Drawing;

namespace WinAPI
{
    public class WindowManipulation
    {
        const int defaultMaxWaitSec = 2;
        const int defaultSleep = 10;

        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MINIMIZE = 0xF020;
        const int SC_MAXIMIZE = 0xF030;
        const int SC_CLOSE = 0xF060;
        const int SC_RESTORE = 0xF120;
        const int WM_ACTIVATEAPP = 0x001C;
        const int WM_USER = 0x0400;
        const int WM_LBUTTONDBLCLK = 0x0203;
        const int WM_LBUTTONDOWN = 0x0201;
        const int WM_LBUTTONUP = 0x0202;
        const int WM_ACTIVATE = 0x0006;
        const int WA_CLICKACTIVE = 2;

        const int TV_FIRST = 0x1100;
        const int TVGN_CARET = 0x9;
        const int TVM_SELECTITEM = (TV_FIRST + 11);
        const int TVM_GETNEXTITEM = (TV_FIRST + 10);
        const int TVM_GETITEM = (TV_FIRST + 12);
        const int TVM_EXPAND = (TV_FIRST + 2);
        const int TVE_COLLAPSE = 0x0001;
        const int TVE_EXPAND = 0x0002;
        const int TVE_COLLAPSERESET = 0x8000;
        
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SendMessage(IntPtr _WindowHandler, int _WM_USER, int wParam, [Out] StringBuilder windowText);
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr HWnd);
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern IntPtr GetActiveWindow();
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);
        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, uint nCmdShow);
        [DllImport("User32.dll")]
        static extern bool MoveWindow(IntPtr win, int x, int y, int width, int height, bool redraw);

        public static void MaximizeWindow(IntPtr win)
        {
            if (win != IntPtr.Zero)
            {
                SendMessage(win, WM_SYSCOMMAND, SC_MINIMIZE, null);
                SendMessage(win, WM_SYSCOMMAND, SC_MAXIMIZE, null);
            }
        }

        public static void MinimizeWindow(IntPtr win)
        {
            if (win != IntPtr.Zero)
                SendMessage(win, WM_SYSCOMMAND, SC_MINIMIZE, null);
        }

        public static void RestoreWindow(IntPtr win)
        {
            if (win != IntPtr.Zero)
                ShowWindow(win, 1);
        }

        public static void ActivateWindow(IntPtr win)
        {
            if (win != IntPtr.Zero) 
                SendMessage(win, WM_ACTIVATE, WA_CLICKACTIVE, 0);
        }

        public static string SetForegroundWindowByName(string processName, string windowSubName, bool maximize)
        {
            Process process = null;

            Process[] p = Process.GetProcessesByName(processName);
            if (p.Length > 0)
            {
                process = p[0];
            }
            else return "";

            string currentText = WindowAttrib.GetForegroundWindowText();

            IntPtr win = process.MainWindowHandle;

            foreach (Process pr in p) pr.Dispose();

            if (windowSubName != "")
                while (win != IntPtr.Zero)
                {
                    string text = WindowAttrib.GetWindowText(win);
                    if (text.Contains(windowSubName))
                    {
                        if (currentText != text)
                        {
                            if (maximize) MaximizeWindow(win);
                            SetForegroundWindow(win);
                        }
                        return WindowAttrib.GetWindowText(win);
                    }
                    win = GetWindow(win, 2);
                }

            return "";
        }

        public static void SetForegroundWindowAndWait(IntPtr win, int maxWaitSec)
        {
            SetForegroundWindow(win);

            IntPtr curr = IntPtr.Zero;

            DateTime start = DateTime.Now;

            while (curr != win)
            {
                curr = GetForegroundWindow();

                if ((DateTime.Now - start).TotalSeconds > maxWaitSec) break;
            }
        }

        public static void SetForegroundWindowAndWait(IntPtr win)
        {
            SetForegroundWindowAndWait(win, defaultMaxWaitSec);
        }

        public static IntPtr WaitToOpen(string path, int maxWaitSec)
        {
            IntPtr win = IntPtr.Zero;

            DateTime start = DateTime.Now;

            while (win == IntPtr.Zero)
            {
                win = WindowSearch.GetWindowByPath(path);

                Thread.Sleep(defaultSleep);

                if ((DateTime.Now - start).TotalSeconds > maxWaitSec) break;
            }

            return win;
        }

        public static IntPtr WaitToOpen(string path)
        {
            return WaitToOpen(path, defaultMaxWaitSec);
        }

        public static IntPtr WaitToFocus(string caption)
        {
            return WaitToFocus(caption, defaultMaxWaitSec);
        }

        public static IntPtr WaitToFocus(string caption, int maxWaitSec)
        {
            IntPtr win = IntPtr.Zero;

            DateTime start = DateTime.Now;

            while (true)
            {
                win = WindowAttrib.GetForegroundWindow();

                string text = WindowAttrib.GetWindowText(win);

                if (text.Contains(caption)) return win;

                Thread.Sleep(defaultSleep);

                if ((DateTime.Now - start).TotalSeconds > maxWaitSec) break;
            }

            return win;
        }

        public static void CloseWindow(IntPtr wnd)
        {
            if (wnd != IntPtr.Zero) 
                SendMessage(wnd, WM_SYSCOMMAND, SC_CLOSE, null);
        }

        public static void MoveWindow(IntPtr win, int x, int y)
        {
            if (win != IntPtr.Zero)
            {
                Rectangle rect = WindowAttrib.GetWindowRect(win);

                MoveWindow(win, x, y, rect.Width, rect.Height, true);
            }
        }

        public static void ResizeWindow(IntPtr win, int width, int height)
        {
            if (win != IntPtr.Zero)
            {
                Rectangle rect = WindowAttrib.GetWindowRect(win);

                MoveWindow(win, rect.X, rect.Y, width, height, true);
            }
        }

        public static void ResizeWindow(IntPtr win, Size size)
        {
            ResizeWindow(win, size.Width, size.Height);
        }

        public static void ResizeWindow_Hard(IntPtr win, int width, int height)
        {
            ResizeWindow_Hard(win, width, height, 5, 5);
        }

        public static void ResizeWindow_Hard(IntPtr win, int width, int height, int xOffset, int yOffset)
        {
            if (win != IntPtr.Zero)
            {
                SetForegroundWindow(win);

                Rectangle rect = WindowAttrib.GetWindowRect(win);

                int dx = width - rect.Width;
                int dy = height - rect.Height;

                Point p = new Point(rect.Location.X + rect.Width - xOffset, rect.Location.Y + rect.Height - yOffset);

                MouseKeyboard.Drag(p, dx, dy);
            }
        }


        public static void ResizeWindow_Hard(IntPtr win, Size size)
        {
            ResizeWindow_Hard(win, size.Width, size.Height);
        }

        public static void ExpandTreeViewItem(IntPtr tree, IntPtr item)
        {
            SendMessage(tree, TVM_EXPAND, TVE_COLLAPSE | TVE_COLLAPSERESET, (int)item);
            SendMessage(tree, TVM_EXPAND, TVE_EXPAND, (int)item);
        }
    }
}
