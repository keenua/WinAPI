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
    public static class WindowManipulation
    {
        const int defaultMaxWaitSec = 2;
        const int defaultSleep = 10;

        const int WM_SYSCOMMAND = 0x0112;
        const int SC_MINIMIZE = 0xF020;
        const int SC_MAXIMIZE = 0xF030;
        const int SC_CLOSE = 0xF060;

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern int SendMessage(IntPtr _WindowHandler, int _WM_USER, int wParam, [Out] StringBuilder windowText);
        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(int HWnd);
        [DllImport("user32.dll")]
        public static extern int GetForegroundWindow();
        [DllImport("user32.dll")]
        public static extern int GetActiveWindow();
        [DllImport("user32.dll")]
        public static extern int GetWindow(int hWnd, int uCmd);
        [DllImport("user32.dll")]
        static extern bool ShowWindow(int hWnd, uint nCmdShow);
        [DllImport("User32.dll")]
        static extern bool MoveWindow(int win, int x, int y, int width, int height, bool redraw);

        public static void MaximizeWindow(int win)
        {
            SendMessage(new IntPtr(win), WM_SYSCOMMAND, SC_MINIMIZE, null);
            SendMessage(new IntPtr(win), WM_SYSCOMMAND, SC_MAXIMIZE, null);
        }

        public static void MinimizeWindow(int win)
        {
            SendMessage(new IntPtr(win), WM_SYSCOMMAND, SC_MINIMIZE, null);
        }

        public static void RestoreWindow(int win)
        {
            ShowWindow(win, 1);
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

            int win = process.MainWindowHandle.ToInt32();

            foreach (Process pr in p) pr.Dispose();

            if (windowSubName != "")
                while (win > 0)
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

        public static void SetForegroundWindowAndWait(int win, int maxWaitSec)
        {
            SetForegroundWindow(win);

            int curr = 0;

            DateTime start = DateTime.Now;

            while (curr != win)
            {
                curr = GetForegroundWindow();

                if ((DateTime.Now - start).TotalSeconds > maxWaitSec) break;
            }
        }

        public static void SetForegroundWindowAndWait(int win)
        {
            SetForegroundWindowAndWait(win, defaultMaxWaitSec);
        }

        public static int WaitToOpen(string path, int maxWaitSec)
        {
            int win = 0;

            DateTime start = DateTime.Now;

            while (win < 1)
            {
                win = WindowSearch.GetWindowByPath(path);

                Thread.Sleep(defaultSleep);

                if ((DateTime.Now - start).TotalSeconds > maxWaitSec) break;
            }

            return win;
        }

        public static int WaitToOpen(string path)
        {
            return WaitToOpen(path, defaultMaxWaitSec);
        }

        public static int WaitToFocus(string caption)
        {
            return WaitToFocus(caption, defaultMaxWaitSec);
        }


        public static int WaitToFocus(string caption, int maxWaitSec)
        {
            int win = 0;

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

        public static void CloseWindow(int wnd)
        {
            if (wnd > 0) SendMessage(new IntPtr(wnd), WM_SYSCOMMAND, SC_CLOSE, null);
        }

        public static void MoveWindow(int win, int x, int y)
        {
            Rectangle rect = WindowAttrib.GetWindowRect(win);

            MoveWindow(win, x, y, rect.Width, rect.Height, true);
        }

        public static void ResizeWindow(int win, int width, int height)
        {
            Rectangle rect = WindowAttrib.GetWindowRect(win);

            MoveWindow(win, rect.X, rect.Y, width, height, true);
        }

        public static void ResizeWindow(int win, Size size)
        {
            ResizeWindow(win, size.Width, size.Height);
        }

        public static void ResizeWindow_Hard(int win, int width, int height)
        {
            ResizeWindow_Hard(win, width, height, 5, 5);
        }

        public static void ResizeWindow_Hard(int win, int width, int height, int xOffset, int yOffset)
        {
            SetForegroundWindow(win);

            Rectangle rect = WindowAttrib.GetWindowRect(win);

            int dx = width - rect.Width;
            int dy = height - rect.Height;

            Point p = new Point(rect.Location.X + rect.Width - xOffset, rect.Location.Y + rect.Height - yOffset);

            MouseKeyboard.Drag(p, dx, dy);
        }


        public static void ResizeWindow_Hard(int win, Size size)
        {
            ResizeWindow_Hard(win, size.Width, size.Height);
        }
    }
}
