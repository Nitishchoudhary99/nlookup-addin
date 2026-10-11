using System;
using System.Globalization;
using Microsoft.Win32;

namespace SmartPro
{
    // small per-user settings kept in the registry (for example the chosen dashboard style)
    internal static class Settings
    {
        const string RegPath = @"Software\DocBrisk\SmartPro";

        public static string Get(string name)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.OpenSubKey(RegPath))
                {
                    return k == null ? null : Convert.ToString(k.GetValue(name), CultureInfo.InvariantCulture);
                }
            }
            catch { return null; }
        }

        public static void Set(string name, string value)
        {
            try
            {
                using (RegistryKey k = Registry.CurrentUser.CreateSubKey(RegPath))
                {
                    k.SetValue(name, value);
                }
            }
            catch { }
        }

        public static int GetInt(string name, int def)
        {
            int n;
            return int.TryParse(Get(name), NumberStyles.Integer, CultureInfo.InvariantCulture, out n) ? n : def;
        }
    }
}
