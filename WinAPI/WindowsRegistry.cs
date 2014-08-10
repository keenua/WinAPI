using System;
using System.Collections.Generic;

using System.Text;
using Microsoft.Win32;

namespace WinAPI
{
    public static class WindowsRegistry
    {
        const string proxySettingsPath = "Software\\Microsoft\\Windows\\CurrentVersion\\Internet Settings";

        public static string GetValue(string path, string name)
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(path);
                string result = (string)key.GetValue(name);
                key.Close();

                return result;
            }
            catch
            {
                return null;
            }
        }

        public static void SetValue(string path, string name, object value)
        {
            try
            {
                RegistryKey key = Registry.CurrentUser.OpenSubKey(path, true);
                key.SetValue(name, value);
            }
            catch
            {

            }
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
    }
}
