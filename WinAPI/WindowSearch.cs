using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Linq;

namespace WinAPI
{
    public static class WindowSearch
    {
        [DllImport("user32.dll", SetLastError = true)]
        static extern int FindWindowEx(int hwndParent, int hwndChildAfter, string lpszClass, string lpszWindow);
        
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(int window, EnumWindowsProc callback, int i);

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        public static List<int> GetChildWindows(int parent)
        {
            List<int> result = new List<int>();
            GCHandle listHandle = GCHandle.Alloc(result);
            try
            {
                EnumWindowsProc childProc = new EnumWindowsProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle).ToInt32());
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result;
        }

        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(int handle, int pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(new IntPtr(pointer));
            List<int> list = gch.Target as List<int>;
            if (list == null)
            {
                throw new InvalidCastException("GCHandle Target could not be cast as List<IntPtr>");
            }
            list.Add(handle);
            //  You can modify this to check to see if you want to cancel the operation, then return a null here
            return true;
        }

        /// <summary>
        /// Delegate for the EnumChildWindows method
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <param name="parameter">Caller-defined variable; we use it for a pointer to our list</param>
        /// <returns>True to continue enumerating, false to bail.</returns>
        public delegate bool EnumWindowsProc(int hWnd, int parameter);

        delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

        static IEnumerable<int> EnumerateProcessWindowHandles(int processId)
        {
            var handles = new List<int>();

            foreach (ProcessThread thread in Process.GetProcessById(processId).Threads)
                EnumThreadWindows(thread.Id, (hWnd, lParam) => { handles.Add(hWnd.ToInt32()); return true; }, IntPtr.Zero);

            return handles;
        }

        /// <summary>Пошук вікна за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static int GetWindowByText(string processName, string windowText, bool exact)
        {
            List<int> wins = GetAllWindowsByProcess(processName, true);

            foreach (int w in wins)
            {
                string text = WindowAttrib.GetWindowText(w);
                if ((exact && text == windowText) || (!exact && text.Contains(windowText))) return w;
            }

            return -1;
        }

        public static int GetWindowByCaption(string processName, string windowCaption, bool exact)
        {
            List<int> wins = GetAllWindowsByProcess(processName, true);

            foreach (int w in wins)
            {
                string text = WindowAttrib.GetWindowCaption(w);
                if ((exact && text == windowCaption) || (!exact && text.Contains(windowCaption))) return w;
            }

            return -1;
        }

        /// <summary>Пошук вікон за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static List<int> GetAllWindowsByText(string processName, string windowText, bool exact)
        {
            List<int> result = new List<int>();

            List<int> wins = GetAllWindowsByProcess(processName, true);

            foreach (int w in wins)
            {
                string text = WindowAttrib.GetWindowText(w);
                if ((exact && text == windowText) || (!exact && text.Contains(windowText))) result.Add(w);
            }

            return result;
        }

        /// <summary>Пошук вікон за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowCaption">Заголовок вікна</param>
        public static List<int> GetAllWindowsByCaption(string processName, string windowCaption, bool exact)
        {
            List<int> result = new List<int>();

            List<int> wins = GetAllWindowsByProcess(processName, true);

            foreach (int w in wins)
            {
                string text = WindowAttrib.GetWindowCaption(w);
                if ((exact && text == windowCaption) || (!exact && text.Contains(windowCaption))) result.Add(w);
            }

            return result;
        }


        /// <summary>Пошук вікна за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static int GetWindowByText(string processName, string windowText)
        {
            return GetWindowByText(processName, windowText, false);
        }

        /// <summary>Повертає усі вікна, що належать заданому процесу</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="onlyVisible">Лише видимі вікна</param>
        public static List<int> GetAllWindowsByProcess(string processName, bool onlyVisible)
        {
            List<int> result = new List<int>();

            Process process = null;

            Process[] p = Process.GetProcessesByName(processName);
            if (p.Count() > 0)
            {
                process = p[0];
            }
            else return result;

            int win = process.MainWindowHandle.ToInt32();

            result.Add(win);

            result.AddRange(GetChildWindows(win));

            List<int> additional = new List<int>();
            foreach (int w in result) if (w != win) additional.AddRange(GetChildWindows(w));

            result.AddRange(additional);

            result = result.Distinct().ToList();

            for (int i = 0; i < result.Count; i++)
            {
                if ((onlyVisible && !WindowAttrib.IsWindowVisible(result[i])) || (WindowAttrib.GetProcessId(result[i]) != process.Id))
                {
                    result.RemoveAt(i);
                    i--;
                }
            }

            return result;
        }

        public static List<int> GetAllThreadWindows(string processName, bool onlyVisible)
        {
            List<int> result = new List<int>();

            Process process = null;

            Process[] p = Process.GetProcessesByName(processName);
            if (p.Count() > 0)
            {
                process = p[0];
            }
            else return result;

            result.AddRange(EnumerateProcessWindowHandles(process.Id));

            for (int i = 0; i < result.Count; i++)
            {
                if (onlyVisible && !WindowAttrib.IsWindowVisible(result[i]))
                {
                    result.RemoveAt(i);
                    i--;
                }
            }

            return result;
        }

        /// <summary>Розбиває нод</summary>
        /// <param name="node">Нод</param>
        /// <param name="name">Назва вікна</param>
        /// <param name="caption">Текст вікна</param>
        /// <param name="index">Порядковий номер</param>
        static void ParseNode(string node, out string name, out string caption, out int index)
        {
            name = null;
            caption = null;
            index = 0;

            string[] split = node.Split('&');
            foreach (string s in split)
            {
                string[] argVal = s.Split('=');
                switch (argVal[0])
                {
                    case "n":
                        name = argVal[1];
                        break;
                    case "c":
                        caption = argVal[1];
                        break;
                }
            }

            split = node.Split('[');
            if (split.Length > 1) index = Convert.ToInt32(split[1].Replace("]", ""));
        }

        /// <summary>Знайти вікно за шляхом</summary>
        /// <param name="path">Шлях до вікна</param>
        public static int GetWindowByPath(string path)
        {
            int win = 0;

            List<string> nodes = path.Split('/').ToList();

            foreach (string node in nodes)
            {
                string name, caption;
                int index;

                ParseNode(node, out name, out caption, out index);

                int childAfter = 0;
                int curWin = 0;

                for (int i = 0; i <= index; i++)
                {
                    curWin = FindWindowEx(win, childAfter, name, caption);
                    if (curWin <= 0) return curWin;
                    childAfter = curWin;
                }

                win = curWin;
            }

            return win;
        }
    }
}
