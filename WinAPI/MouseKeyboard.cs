using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Drawing;
using System.Windows.Forms;
using System.Linq;

namespace WinAPI
{
    /// <summary>
    /// 
    /// </summary>
    public static class MouseKeyboard
    {
        #region Константи

        [Flags]
        public enum MouseFlags { Move = 0x0001, LeftDown = 0x0002, LeftUp = 0x0004, RightDown = 0x0008, RightUp = 0x0010, Absolute = 0x8000 };
        public const UInt32 KEYEVENTF_EXTENDEDKEY = 1;
        public const UInt32 KEYEVENTF_KEYUP = 2;
        public const int KEY_ALT = 0x12;
        public const int KEY_CONTROL = 0x11;

        public enum WMessages : int
        {
            WM_LBUTTONDOWN = 0x201, //Left mousebutton down
            WM_LBUTTONUP = 0x202,  //Left mousebutton up
            WM_LBUTTONDBLCLK = 0x203, //Left mousebutton doubleclick
            WM_RBUTTONDOWN = 0x204, //Right mousebutton down
            WM_RBUTTONUP = 0x205,   //Right mousebutton up
            WM_RBUTTONDBLCLK = 0x206, //Right mousebutton doubleclick
            WM_KEYDOWN = 0x100,  //Key down
            WM_KEYUP = 0x101,   //Key up
        }

        #endregion

        #region User32.dll

        /// <summary>Функція для керування клавіатурою</summary>
        /// <param name="bVk">Код клавіші</param>
        /// <param name="bScan">Скан</param>
        /// <param name="dwFlags">Тип дії</param>
        /// <param name="dwExtraInfo">Додаткові дані</param>
        [DllImport("user32.dll")]
        public static extern void keybd_event(int bVk, byte bScan, UInt32 dwFlags, int dwExtraInfo);

        /// <summary>Взять координаты курсора</summary>
        /// <param name="lpPoint">Координаты курсора</param>
        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(ref Point lpPoint);

        /// <summary>Указать координаты курсора</summary>
        /// <param name="x">Х-координата</param>
        /// <param name="y">Y-координата</param>
        [DllImport("user32.dll")]
        private static extern void SetCursorPos(int x, int y);

        /// <summary>Функция для управления мышью</summary>
        /// <param name="dwFlags">Тип действия</param>
        /// <param name="dx">Х-координата</param>
        /// <param name="dy">Y-координата</param>
        /// <param name="dwData">Данные</param>
        /// <param name="dwExtraInfo">Дополнительные данные</param>
        [DllImport("user32.dll")]
        private static extern void mouse_event(MouseFlags dwFlags, int dx, int dy, int dwData, UIntPtr dwExtraInfo);

        #endregion

        static Random random = new Random((int)(DateTime.Now.Ticks % int.MaxValue));

        static Point RandomPoint(Rectangle rect)
        {
            // Випадково визначаємо точку всередині цього центру
            int x = random.Next(rect.X, rect.X + rect.Width);
            int y = random.Next(rect.Y, rect.Y + rect.Height);

            return new Point(x, y);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="p"></param>
        public static void SetCursor(Point p)
        {
            SetCursorPos(p.X, p.Y);
        }

        /// <summary>
        /// Натиснути ліву клавішу миші
        /// </summary>
        /// <param name="win">Вікно</param>
        /// <param name="time">Час зажимання</param>
        /// <param name="x">Розташування курсора (відносна координата у вікні)</param>
        /// <param name="y">Розташування курсора (відносна координата у вікні)</param>
        public static void PressLeftMouseButton(int win, int time, int x, int y)
        {
            if (win > 0)
            {
                Rectangle rect = WindowAttrib.GetClientRect(win);

                x += rect.X;
                y += rect.Y;
            }

            SetCursorPos(x, y);

            PressLeftMouseButton(time);
        }

        /// <summary>
        /// Натиснути ліву клавішу миші
        /// </summary>
        /// <param name="win">Вікно</param>
        /// <param name="time">Час зажимання</param>
        /// <param name="p">Розташування курсора (відносні координати у вікні)</param>
        public static void PressLeftMouseButton(int win, int time, Point p)
        {
            PressLeftMouseButton(win, time, p.X, p.Y);
        }

        /// <summary>
        /// Натиснути ліву клавішу миші
        /// </summary>
        /// <param name="time">Час зажимання</param>
        /// <param name="x">Розташування курсора</param>
        /// <param name="y">Розташування курсора</param>
        public static void PressLeftMouseButton(int time, int x, int y)
        {
            PressLeftMouseButton(0, time, x, y);
        }

        /// <summary>
        /// Натиснути ліву клавішу миші
        /// </summary>
        /// <param name="time">Час зажимання</param>
        /// <param name="p">Розташування курсора</param>
        public static void PressLeftMouseButton(int time, Point p)
        {
            SetCursor(p);

            PressLeftMouseButton(time);
        }

        /// <summary>
        /// Натиснути ліву клавішу миші
        /// </summary>
        /// <param name="p">Розташування курсора</param>
        public static void PressLeftMouseButton(Point p)
        {
            SetCursor(p);

            PressLeftMouseButton(0);
        }

        public static void Drag(Point p1, Point p2)
        {
            SetCursor(p1);

            Thread.Sleep(200);

            // Нажимаем левую кнопку мыши
            mouse_event(MouseFlags.Absolute | MouseFlags.LeftDown, p1.X, p1.Y, 0, new UIntPtr());

            Point cursor = new Point();
            
            GetCursorPos(ref cursor);

            while (cursor != p2)
            {
                cursor = new Point(cursor.X + Math.Sign(p2.X - cursor.X), cursor.Y + Math.Sign(p2.Y - cursor.Y));

                SetCursor(cursor);
            }

            //Thread.Sleep(50);

            // Отпускаем кнопку
            mouse_event(MouseFlags.Absolute | MouseFlags.LeftUp, p2.X, p2.Y, 0, new UIntPtr());

            //Thread.Sleep(50);
        }

        public static void Drag(Rectangle r1, Rectangle r2)
        {
            Point p1 = RandomPoint(r1);
            Point p2 = RandomPoint(r2);

            Drag(p1, p2);
        }

        public static void Drag(Point p1, int dx, int dy)
        {
            Point p2 = new Point(p1.X + dx, p1.Y + dy);

            Drag(p1, p2);
        }

        /// <summary>"Доїжджає" курсором до заданого прямокутника</summary>
        /// <param name="box">Прямокутник</param>
        /// <param name="speed">Швидкість переміщення</param>
        /// <param name="clicks">Кількість натискань на клавішу миші після руху</param>
        /// <remarks>** НЕ ТЕСТОВАНО **</remarks>
        public static void ClickOnRect(Rectangle rect, int clicks, bool left)
        {
            Point p = RandomPoint(rect);

            SetCursorPos(p.X, p.Y);
            
            // Якщо потрібно це зробити більше ніх раз
            for (int i = 0; i < clicks; i++)
            {
                // Затримка 100 мсек
                Thread.Sleep(100);
                // Натискаємо
                if (left) PressLeftMouseButton(0);
                else PressRightButton(0);
            }
        }

        /// <summary>"Доїжджає" курсором до заданого прямокутника</summary>
        /// <param name="box">Прямокутник</param>
        /// <param name="speed">Швидкість переміщення</param>
        /// <param name="clicks">Кількість натискань на клавішу миші після руху</param>
        /// <remarks>** НЕ ТЕСТОВАНО **</remarks>
        public static void ClickOnRect(Rectangle rect, int clicks)
        {
            ClickOnRect(rect, clicks, true);
        }

        /// <summary>"Доїжджає" курсором до заданого прямокутника</summary>
        /// <param name="box">Прямокутник</param>
        /// <param name="speed">Швидкість переміщення</param>
        /// <param name="clicks">Кількість натискань на клавішу миші після руху</param>
        /// <remarks>** НЕ ТЕСТОВАНО **</remarks>
        public static void ClickOnRect(int win, Rectangle rect, int clicks)
        {
            Rectangle winRect = WindowAttrib.GetClientRect(win);

            Rectangle clickRect = new Rectangle(rect.X + winRect.X, rect.Y + winRect.Y, rect.Width, rect.Height);

            ClickOnRect(clickRect, clicks);
        }

        public static bool ClickOnButton(string processName, string buttonText, bool exact)
        {
            List<int> wins = WindowSearch.GetAllWindowsByText(processName, buttonText, exact);
            wins.AddRange(WindowSearch.GetAllWindowsByCaption(processName, buttonText, exact));
            wins = wins.Distinct().ToList();
            
            if (wins.Count == 0) return false;

            int win = wins[wins.Count - 1];

            //WindowManipulation.SetForegroundWindow(win);

            Thread.Sleep(500);

            //WindowManipulation.SetForegroundWindow(win);

            Rectangle rect = WindowAttrib.GetWindowRect(win);

            rect.Location = new Point(rect.X + 7, rect.Y + 7);
            rect.Size = new Size(rect.Width - 10, rect.Height - 10);

            ClickOnRect(rect, 1);

            Thread.Sleep(200);

            return true;
        }

        /// <summary>
        /// Натиснути ліву клавішу миші
        /// </summary>
        /// <param name="time">Час зажимання</param>
        public static void PressLeftMouseButton(int time)
        {
            // Берем координаты курсора
            Point point = new Point();
            GetCursorPos(ref point);

            // Нажимаем левую кнопку мыши
            mouse_event(MouseFlags.Absolute | MouseFlags.LeftDown, point.X, point.Y, 0, new UIntPtr());

            // Если нужно, делаем задержку
            if (time > 0) Thread.Sleep(time);

            // Отпускаем кнопку
            mouse_event(MouseFlags.Absolute | MouseFlags.LeftUp, point.X, point.Y, 0, new UIntPtr());
        }

        public static void PressRightButton(int time)
        {
            // Берем координаты курсора
            Point point = new Point();
            GetCursorPos(ref point);

            // Нажимаем левую кнопку мыши
            mouse_event(MouseFlags.Absolute | MouseFlags.RightDown, point.X, point.Y, 0, new UIntPtr());

            // Если нужно, делаем задержку
            if (time > 0) Thread.Sleep(time);

            // Отпускаем кнопку
            mouse_event(MouseFlags.Absolute | MouseFlags.RightUp, point.X, point.Y, 0, new UIntPtr());
        }

        /// <summary>Натиснути на клавішу клавіатури</summary>
        /// <param name="key">Код клавіші</param>
        /// <param name="shift">З натисненим shift</param>
        /// <param name="alt">З натисненим alt</param>
        /// <param name="ctrl">З натисненим ctrl</param>
        /// <remarks>** НЕ ТЕСТОВАНО **</remarks>
        public static void PressKey(int key, bool shift, bool alt, bool ctrl)
        {
            // Якщо натискаємо PrintScreen, то bScan = 0
            byte bScan = 0x45;
            if (key == KeyConstants.VK_SNAPSHOT) bScan = 0;
            if (key == KeyConstants.VK_SPACE) bScan = 39;

            // Якщо потрібно, затискаємо alt, ctrl чи shift
            if (alt) keybd_event(KeyConstants.VK_MENU, 0, 0, 0);
            if (ctrl) keybd_event(KeyConstants.VK_LCONTROL, 0, 0, 0);
            if (shift) keybd_event(KeyConstants.VK_LSHIFT, 0, 0, 0);

            // Натискаємо вказану клавішу
            keybd_event(key, bScan, KEYEVENTF_EXTENDEDKEY, 0);

            // Відпускаємо її
            keybd_event(key, bScan, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);

            // Якщо затиснули котрісь клавіші, відпускаємо їх
            if (shift) keybd_event(KeyConstants.VK_LSHIFT, 0, KEYEVENTF_KEYUP, 0);
            if (ctrl) keybd_event(KeyConstants.VK_LCONTROL, 0, KEYEVENTF_KEYUP, 0);
            if (alt) keybd_event(KeyConstants.VK_MENU, 0, KEYEVENTF_KEYUP, 0);
        }

        /// <summary>Натиснути на клавішу клавіатури</summary>
        /// <param name="key">Код клавіші</param>
        /// <remarks>** НЕ ТЕСТОВАНО **</remarks>
        public static void PressKey(int key)
        {
            PressKey(key, false, false, false);
        }

        public static void Paste(string text, int delay)
        {
            if (text == "" || text == null) return;

            string backup = Clipboard.GetText();

            Thread.Sleep(delay);

            Clipboard.SetText(text);

            Thread.Sleep(delay);

            PressKey((int)'V', false, false, true);

            Thread.Sleep(delay);

            Clipboard.SetText(backup);
        }

        public static void Type(string text, int delay)
        {
            Type(text, delay, delay);
        }

        public static void Type(string text, int delayFrom, int delayTo)
        {
            for (int i = 0; i < text.Length; i++)
            {
                char c = text[i];

                bool shift = char.IsUpper(c);

                PressKey((int)char.ToUpper(c), shift, false, false);

                int delay = random.Next(delayFrom, delayTo);

                Thread.Sleep(delay);
            }
        }
    }
}
