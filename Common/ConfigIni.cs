using System;
using System.Reflection;
using System.IO;

namespace Common
{
    public class ConfigIni
    {
        public static string pptPath;
        public static string pptSusPath;
        public static string originalPath;
        public static string pdfPath;
        public static string docPath;
        public static string originalPathTwo;

        public static string appId;
        public static string passWord;
        public static string baiduUrl;

        public static void GetIniVal()
        {
            var fields = typeof(ConfigIni).GetFields();
            foreach (var item in fields)
            {
                string val = IniFunc.getString(item.Name, null);
                item.SetValue(null, val);
            }
        }
        public static void SetIniVal()
        {
            var fields = typeof(ConfigIni).GetFields();
            foreach (var item in fields)
            {
                string val = item.GetValue(null).ToString();
                IniFunc.writeString(item.Name, val);
      
            }
        }

    }

}
