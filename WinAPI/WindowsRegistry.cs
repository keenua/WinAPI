using System;
using System.Collections.Generic;
using System.Text;
using Microsoft.Win32;
using System.Windows.Forms;
using System.Threading;

namespace WinAPI
{
    public class WindowsRegistry
    {
        const string proxySettingsPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings";

        public static string GetValue(RegistryKey root, string path, string name)
        {
            try
            {
                RegistryKey key = root.OpenSubKey(path);
                string result = (string)key.GetValue(name);
                key.Close();

                return result;
            }
            catch
            {
                return null;
            }
        }

        public static void SetValue(RegistryKey root, string path, string name, object value)
        {
            try
            {
                RegistryKey key = root.OpenSubKey(path, true);
                key.SetValue(name, value);
            }
            catch
            {

            }
        }

        public static string GetValue(string path, string name)
        {
            return GetValue(Registry.CurrentUser, path, name);
        }

        public static void SetValue(string path, string name, object value)
        {
            SetValue(Registry.CurrentUser, path, name, value);
        }

        public static string GetValueLocalMachine(string path, string name)
        {
            return GetValue(Registry.LocalMachine, path, name);
        }

        public static void Delete(string path, string name)
        {
            try
            {
                RegistryKey parentKey = Registry.CurrentUser.OpenSubKey(path, true);

                RegistryKey key = parentKey.OpenSubKey(name);

                if (key != null) parentKey.DeleteSubKeyTree(name);
            }
            catch
            {

            }
        }

        public static void SetProxy(string proxy)
        {
            SetValue(proxySettingsPath, "ProxyEnable", 1);
            SetValue(proxySettingsPath, "ProxyServer", proxy);
        }

        public static void RemoveProxy()
        {
            SetValue(proxySettingsPath, "ProxyEnable", 0);
            SetValue(proxySettingsPath, "ProxyServer", "");
        }

        protected static string clipboardText = string.Empty;
        protected static void GetCliboardMethod()
        {
            try
            {
                clipboardText = Clipboard.GetText();
            }
            catch { }
        }

        public static string GetClipboard()
        {
            clipboardText = string.Empty;
            Thread staThread = new Thread(GetCliboardMethod);

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join();

            return clipboardText;
        }

        public static bool SetClipboard(string text)
        {
            try
            {
                Clipboard.SetText(text);
                Thread.Sleep(100);
            }
            catch { }

            return GetClipboard() == text;
        }

        public static void RemoveKey(string key)
        {
            try
            {
                Registry.CurrentUser.DeleteSubKeyTree(key);
            }
            catch { }
        }

        public static void RemoveKeyLocalMachine(string key)
        {
            try
            {
                Registry.LocalMachine.DeleteSubKeyTree(key);
            }
            catch
            { }
        }
    }
}
