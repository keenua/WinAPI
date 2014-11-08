using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Linq;

namespace WinAPI
{
    public class WindowSearch
    {
        const int GW_HWNDNEXT = 2;
        const int GW_CHILD = 5;

        const int TV_FIRST = 0x1100;
        const int TVGN_ROOT = 0x0;
        const int TVGN_NEXT = 0x1;
        const int TVGN_CHILD = 0x4;
        const int TVGN_FIRSTVISIBLE = 0x5;
        const int TVGN_NEXTVISIBLE = 0x6;
        const int TVGN_CARET = 0x9;
        const int TVM_SELECTITEM = (TV_FIRST + 11);
        const int TVM_GETNEXTITEM = (TV_FIRST + 10);
        const int TVM_GETITEM = (TV_FIRST + 12);
        const int TVM_EXPAND = (TV_FIRST + 2);
        const int TVE_COLLAPSE = 0x0001;
        const int TVE_EXPAND = 0x0002;
        const int TVE_COLLAPSERESET = 0x8000;
        const int LVM_GETNEXTITEM = 0x100c;
        const int LVNI_BELOW = 0x0200;

        // Enumeration is set to unicode, ANSI counterparts are commented out.
        // Contains a few undocumented messages of which the name was invented.
        public enum LVM
        {
            FIRST = 0x1000,
            SETUNICODEFORMAT = 0x2005,        // CCM_SETUNICODEFORMAT,
            GETUNICODEFORMAT = 0x2006,        // CCM_GETUNICODEFORMAT,
            GETBKCOLOR = (FIRST + 0),
            SETBKCOLOR = (FIRST + 1),
            GETIMAGELIST = (FIRST + 2),
            SETIMAGELIST = (FIRST + 3),
            GETITEMCOUNT = (FIRST + 4),
            GETITEMA = (FIRST + 5),
            GETITEMW = (FIRST + 75),
            GETITEM = GETITEMW,
            //GETITEM                = GETITEMA,
            SETITEMA = (FIRST + 6),
            SETITEMW = (FIRST + 76),
            SETITEM = SETITEMW,
            //SETITEM                = SETITEMA,
            INSERTITEMA = (FIRST + 7),
            INSERTITEMW = (FIRST + 77),
            INSERTITEM = INSERTITEMW,
            //INSERTITEM             = INSERTITEMA,
            DELETEITEM = (FIRST + 8),
            DELETEALLITEMS = (FIRST + 9),
            GETCALLBACKMASK = (FIRST + 10),
            SETCALLBACKMASK = (FIRST + 11),
            GETNEXTITEM = (FIRST + 12),
            FINDITEMA = (FIRST + 13),
            FINDITEMW = (FIRST + 83),
            GETITEMRECT = (FIRST + 14),
            SETITEMPOSITION = (FIRST + 15),
            GETITEMPOSITION = (FIRST + 16),
            GETSTRINGWIDTHA = (FIRST + 17),
            GETSTRINGWIDTHW = (FIRST + 87),
            HITTEST = (FIRST + 18),
            ENSUREVISIBLE = (FIRST + 19),
            SCROLL = (FIRST + 20),
            REDRAWITEMS = (FIRST + 21),
            ARRANGE = (FIRST + 22),
            EDITLABELA = (FIRST + 23),
            EDITLABELW = (FIRST + 118),
            EDITLABEL = EDITLABELW,
            //EDITLABEL              = EDITLABELA,
            GETEDITCONTROL = (FIRST + 24),
            GETCOLUMNA = (FIRST + 25),
            GETCOLUMNW = (FIRST + 95),
            SETCOLUMNA = (FIRST + 26),
            SETCOLUMNW = (FIRST + 96),
            INSERTCOLUMNA = (FIRST + 27),
            INSERTCOLUMNW = (FIRST + 97),
            DELETECOLUMN = (FIRST + 28),
            GETCOLUMNWIDTH = (FIRST + 29),
            SETCOLUMNWIDTH = (FIRST + 30),
            GETHEADER = (FIRST + 31),
            CREATEDRAGIMAGE = (FIRST + 33),
            GETVIEWRECT = (FIRST + 34),
            GETTEXTCOLOR = (FIRST + 35),
            SETTEXTCOLOR = (FIRST + 36),
            GETTEXTBKCOLOR = (FIRST + 37),
            SETTEXTBKCOLOR = (FIRST + 38),
            GETTOPINDEX = (FIRST + 39),
            GETCOUNTPERPAGE = (FIRST + 40),
            GETORIGIN = (FIRST + 41),
            UPDATE = (FIRST + 42),
            SETITEMSTATE = (FIRST + 43),
            GETITEMSTATE = (FIRST + 44),
            GETITEMTEXTA = (FIRST + 45),
            GETITEMTEXTW = (FIRST + 115),
            SETITEMTEXTA = (FIRST + 46),
            SETITEMTEXTW = (FIRST + 116),
            SETITEMCOUNT = (FIRST + 47),
            SORTITEMS = (FIRST + 48),
            SETITEMPOSITION32 = (FIRST + 49),
            GETSELECTEDCOUNT = (FIRST + 50),
            GETITEMSPACING = (FIRST + 51),
            GETISEARCHSTRINGA = (FIRST + 52),
            GETISEARCHSTRINGW = (FIRST + 117),
            GETISEARCHSTRING = GETISEARCHSTRINGW,
            //GETISEARCHSTRING       = GETISEARCHSTRINGA,
            SETICONSPACING = (FIRST + 53),
            SETEXTENDEDLISTVIEWSTYLE = (FIRST + 54),            // optional wParam == mask
            GETEXTENDEDLISTVIEWSTYLE = (FIRST + 55),
            GETSUBITEMRECT = (FIRST + 56),
            SUBITEMHITTEST = (FIRST + 57),
            SETCOLUMNORDERARRAY = (FIRST + 58),
            GETCOLUMNORDERARRAY = (FIRST + 59),
            SETHOTITEM = (FIRST + 60),
            GETHOTITEM = (FIRST + 61),
            SETHOTCURSOR = (FIRST + 62),
            GETHOTCURSOR = (FIRST + 63),
            APPROXIMATEVIEWRECT = (FIRST + 64),
            SETWORKAREAS = (FIRST + 65),
            GETWORKAREAS = (FIRST + 70),
            GETNUMBEROFWORKAREAS = (FIRST + 73),
            GETSELECTIONMARK = (FIRST + 66),
            SETSELECTIONMARK = (FIRST + 67),
            SETHOVERTIME = (FIRST + 71),
            GETHOVERTIME = (FIRST + 72),
            SETTOOLTIPS = (FIRST + 74),
            GETTOOLTIPS = (FIRST + 78),
            SORTITEMSEX = (FIRST + 81),
            SETBKIMAGEA = (FIRST + 68),
            SETBKIMAGEW = (FIRST + 138),
            GETBKIMAGEA = (FIRST + 69),
            GETBKIMAGEW = (FIRST + 139),
            SETSELECTEDCOLUMN = (FIRST + 140),
            SETVIEW = (FIRST + 142),
            GETVIEW = (FIRST + 143),
            INSERTGROUP = (FIRST + 145),
            SETGROUPINFO = (FIRST + 147),
            GETGROUPINFO = (FIRST + 149),
            REMOVEGROUP = (FIRST + 150),
            MOVEGROUP = (FIRST + 151),
            GETGROUPCOUNT = (FIRST + 152),
            GETGROUPINFOBYINDEX = (FIRST + 153),
            MOVEITEMTOGROUP = (FIRST + 154),
            GETGROUPRECT = (FIRST + 98),
            SETGROUPMETRICS = (FIRST + 155),
            GETGROUPMETRICS = (FIRST + 156),
            ENABLEGROUPVIEW = (FIRST + 157),
            SORTGROUPS = (FIRST + 158),
            INSERTGROUPSORTED = (FIRST + 159),
            REMOVEALLGROUPS = (FIRST + 160),
            HASGROUP = (FIRST + 161),
            GETGROUPSTATE = (FIRST + 92),
            GETFOCUSEDGROUP = (FIRST + 93),
            SETTILEVIEWINFO = (FIRST + 162),
            GETTILEVIEWINFO = (FIRST + 163),
            SETTILEINFO = (FIRST + 164),
            GETTILEINFO = (FIRST + 165),
            SETINSERTMARK = (FIRST + 166),
            GETINSERTMARK = (FIRST + 167),
            INSERTMARKHITTEST = (FIRST + 168),
            GETINSERTMARKRECT = (FIRST + 169),
            SETINSERTMARKCOLOR = (FIRST + 170),
            GETINSERTMARKCOLOR = (FIRST + 171),
            GETSELECTEDCOLUMN = (FIRST + 174),
            ISGROUPVIEWENABLED = (FIRST + 175),
            GETOUTLINECOLOR = (FIRST + 176),
            SETOUTLINECOLOR = (FIRST + 177),
            CANCELEDITLABEL = (FIRST + 179),
            MAPINDEXTOID = (FIRST + 180),
            MAPIDTOINDEX = (FIRST + 181),
            ISITEMVISIBLE = (FIRST + 182),
            GETACCVERSION = (FIRST + 193),
            GETEMPTYTEXT = (FIRST + 204),
            GETFOOTERRECT = (FIRST + 205),
            GETFOOTERINFO = (FIRST + 206),
            GETFOOTERITEMRECT = (FIRST + 207),
            GETFOOTERITEM = (FIRST + 208),
            GETITEMINDEXRECT = (FIRST + 209),
            SETITEMINDEXSTATE = (FIRST + 210),
            GETNEXTITEMINDEX = (FIRST + 211),
            SETPRESERVEALPHA = (FIRST + 212),
            SETBKIMAGE = SETBKIMAGEW,
            GETBKIMAGE = GETBKIMAGEW,
            //SETBKIMAGE             = SETBKIMAGEA,
            //GETBKIMAGE             = GETBKIMAGEA,
        }

        [DllImport("user32.dll")]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);
        
        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumChildWindows(IntPtr window, EnumWindowsProc callback, IntPtr i);

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindow(IntPtr hWnd, int uCmd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, int Msg, int wParam, IntPtr lParam);

        /// <summary>
        /// Returns a list of child windows
        /// </summary>
        /// <param name="parent">Parent of the windows to return</param>
        /// <returns>List of child windows</returns>
        public static List<IntPtr> GetChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();
            GCHandle listHandle = GCHandle.Alloc(result);

            try
            {
                EnumWindowsProc childProc = new EnumWindowsProc(EnumWindow);
                EnumChildWindows(parent, childProc, GCHandle.ToIntPtr(listHandle));
            }
            finally
            {
                if (listHandle.IsAllocated)
                    listHandle.Free();
            }
            return result.Distinct().ToList();
        }

        /// <summary>
        /// Callback method to be used when enumerating windows.
        /// </summary>
        /// <param name="handle">Handle of the next window</param>
        /// <param name="pointer">Pointer to a GCHandle that holds a reference to the list to fill</param>
        /// <returns>True to continue the enumeration, false to bail</returns>
        private static bool EnumWindow(IntPtr handle, int pointer)
        {
            GCHandle gch = GCHandle.FromIntPtr(new IntPtr(pointer));
            List<IntPtr> list = gch.Target as List<IntPtr>;
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
        public delegate bool EnumWindowsProc(IntPtr hWnd, int parameter);

        delegate bool EnumThreadDelegate(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll")]
        static extern bool EnumThreadWindows(int dwThreadId, EnumThreadDelegate lpfn, IntPtr lParam);

        static IEnumerable<IntPtr> EnumerateProcessWindowHandles(int processId)
        {
            var handles = new List<IntPtr>();

            foreach (ProcessThread thread in Process.GetProcessById(processId).Threads)
                EnumThreadWindows(thread.Id, (hWnd, lParam) => { handles.Add(hWnd); return true; }, IntPtr.Zero);

            return handles;
        }

        /// <summary>Пошук вікна за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static IntPtr GetWindowByText(string processName, string windowText, bool exact, bool onlyVisible)
        {
            List<IntPtr> wins = GetAllWindowsByProcess(processName, onlyVisible);

            foreach (IntPtr w in wins)
            {
                string text = WindowAttrib.GetWindowText(w);
                if ((exact && text == windowText) || (!exact && text.Contains(windowText))) return w;
            }

            return IntPtr.Zero;
        }

        /// <summary>Пошук вікна за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static IntPtr GetWindowByText(string processName, string windowText, bool onlyVisible)
        {
            return GetWindowByText(processName, windowText, false, onlyVisible);
        }


        /// <summary>Пошук вікна за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static IntPtr GetWindowByText(string processName, string windowText)
        {
            return GetWindowByText(processName, windowText, false);
        }

        /// <summary>Пошук вікна за назвою процесу та заголовком вікна (з використанням GetAllThreadWindows)</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static IntPtr GetThreadWindowByText(string processName, string windowText, bool exact, bool onlyVisible)
        {
            List<IntPtr> wins = GetAllThreadWindows(processName, onlyVisible);

            foreach (IntPtr w in wins)
            {
                string text = WindowAttrib.GetWindowText(w);
                if ((exact && text == windowText) || (!exact && text.Contains(windowText))) return w;
            }

            return IntPtr.Zero;
        }

        public static IntPtr GetThreadWindowByText(string processName, string windowText, bool exact)
        {
            return GetThreadWindowByText(processName, windowText, exact, true);
        }

        public static IntPtr GetChildWindowByText(IntPtr parent, string windowText, bool exact)
        {
            var wins = WindowSearch.GetChildWindows(parent);

            foreach (IntPtr w in wins)
            {
                string text = WindowAttrib.GetWindowText(w);
                if ((exact && text == windowText) || (!exact && text.Contains(windowText))) return w;
            }

            return IntPtr.Zero;
        }

        public static IntPtr GetWindowByCaption(string processName, string windowCaption, bool exact)
        {
            List<IntPtr> wins = GetAllWindowsByProcess(processName, true);

            foreach (IntPtr w in wins)
            {
                string text = WindowAttrib.GetWindowCaption(w);
                if ((exact && text == windowCaption) || (!exact && text.Contains(windowCaption))) return w;
            }

            return IntPtr.Zero;
        }

        /// <summary>Пошук вікон за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowText">Заголовок вікна</param>
        public static List<IntPtr> GetAllWindowsByText(string processName, string windowText, bool exact)
        {
            List<IntPtr> result = new List<IntPtr>();

            List<IntPtr> wins = GetAllWindowsByProcess(processName, true);

            foreach (IntPtr w in wins)
            {
                string text = WindowAttrib.GetWindowText(w);
                if ((exact && text == windowText) || (!exact && text.Contains(windowText))) result.Add(w);
            }

            return result;
        }

        /// <summary>Пошук вікон за назвою процесу та заголовком вікна</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="windowCaption">Заголовок вікна</param>
        public static List<IntPtr> GetAllWindowsByCaption(string processName, string windowCaption, bool exact)
        {
            List<IntPtr> result = new List<IntPtr>();

            List<IntPtr> wins = GetAllWindowsByProcess(processName, true);

            foreach (IntPtr w in wins)
            {
                string text = WindowAttrib.GetWindowCaption(w);
                if ((exact && text == windowCaption) || (!exact && text.Contains(windowCaption))) result.Add(w);
            }

            return result;
        }

        /// <summary>Повертає усі вікна, що належать заданому процесу</summary>
        /// <param name="processName">Назва процесу</param>
        /// <param name="onlyVisible">Лише видимі вікна</param>
        public static List<IntPtr> GetAllWindowsByProcess(string processName, bool onlyVisible)
        {
            List<IntPtr> result = new List<IntPtr>();

            Process process = null;

            Process[] p = Process.GetProcessesByName(processName);
            if (p.Count() > 0)
            {
                process = p[0];
            }
            else return result;

            IntPtr win = process.MainWindowHandle;

            result.Add(win);

            result.AddRange(GetChildWindows(win));

            List<IntPtr> additional = new List<IntPtr>();
            foreach (IntPtr w in result) if (w != win) additional.AddRange(GetChildWindows(w));

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

        public static List<IntPtr> GetAllThreadWindows(string processName, bool onlyVisible)
        {
            List<IntPtr> result = new List<IntPtr>();

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
        public static IntPtr GetWindowByPath(string path)
        {
            IntPtr win = IntPtr.Zero;

            List<string> nodes = path.Split('/').ToList();

            foreach (string node in nodes)
            {
                string name, caption;
                int index;

                ParseNode(node, out name, out caption, out index);

                IntPtr childAfter = IntPtr.Zero;
                IntPtr curWin = IntPtr.Zero;

                for (int i = 0; i <= index; i++)
                {
                    curWin = FindWindowEx(win, childAfter, name, caption);
                    if (curWin == IntPtr.Zero) return curWin;
                    childAfter = curWin;
                }

                win = curWin;
            }

            return win;
        }

        public static List<IntPtr> GetDirectChildWindows(IntPtr parent)
        {
            List<IntPtr> result = new List<IntPtr>();

            var child = GetWindow(parent, GW_CHILD);

            while (child != IntPtr.Zero)
            {
                result.Add(child);

                child = GetWindow(child, GW_HWNDNEXT);
            }

            return result;
        }

        public static List<IntPtr> GetTreeViewItemChildItems(IntPtr treeView, IntPtr item)
        {
            List<IntPtr> children = new List<IntPtr>();

            if (item == IntPtr.Zero) return children;

            children.Add(item);

            var current = SendMessage(treeView, TVM_GETNEXTITEM, TVGN_CHILD, item);

            while (current != IntPtr.Zero)
            {
                children.AddRange(GetTreeViewItemChildItems(treeView, current));

                current = SendMessage(treeView, TVM_GETNEXTITEM, TVGN_NEXT, current);

                if (current == IntPtr.Zero) break;
            }

            return children;
        }

        public static List<IntPtr> GetAllTreeViewItems(IntPtr treeView)
        {
            var item = SendMessage(treeView, TVM_GETNEXTITEM, TVGN_ROOT, IntPtr.Zero);

            List<IntPtr> result = new List<IntPtr>();

            while (item != IntPtr.Zero)
            {
                result.AddRange(GetTreeViewItemChildItems(treeView, item));
                item = SendMessage(treeView, TVM_GETNEXTITEM, TVGN_NEXT, item);
            }

            return result;
        }
        
        public static IntPtr GetChildByPath(IntPtr parent, params int[] path)
        {
            if (parent == IntPtr.Zero) return IntPtr.Zero;

            var current = parent;

            foreach (var i in path)
            {
                var children = GetDirectChildWindows(current);

                if (children.Count <= i) return IntPtr.Zero;

                current = children[i];
            }

            return current;
        }

        public static IntPtr GetWindowByClass(IntPtr parent, string className)
        {
            string c = WindowAttrib.GetWindowClass(parent);

            if (c == className) return parent;

            foreach (var w in GetChildWindows(parent))
            {
                var found = GetWindowByClass(w, className);

                if (found != IntPtr.Zero) return found;
            }

            return IntPtr.Zero;
        }

        public static List<IntPtr> GetListViewItems(IntPtr listView)
        {
            var header = SendMessage(listView, (int)LVM.GETHEADER, 0, IntPtr.Zero);

            Console.WriteLine(header);

            Console.WriteLine(WindowAttrib.GetWindowText(header));

            return new List<IntPtr>();
        }
    }
}
