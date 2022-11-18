﻿using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace Common
{
    /// <summary>
    /// 公共方法类
    /// </summary>
    public class IniFunc
    {
        /// <summary>
        /// 获取值
        /// </summary>
        /// <param name="section">段落名</param>
        /// <param name="key">键名</param>
        /// <param name="defval">读取异常是的缺省值</param>
        /// <param name="retval">键名所对应的的值，没有找到返回空值</param>
        /// <param name="size">返回值允许的大小</param>
        /// <param name="filepath">ini文件的完整路径</param>
        /// <returns></returns>
        [DllImport("kernel32")]
        public static extern int GetPrivateProfileString(string section, string key, string def, StringBuilder retVal, int size, string filePath);
        /// <summary>
        /// 写入
        /// </summary>
        /// <param name="section">需要写入的段落名</param>
        /// <param name="key">需要写入的键名</param>
        /// <param name="val">写入值</param>
        /// <param name="filepath">ini文件的完整路径</param>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        private static extern int WritePrivateProfileString(string section, string key, string val, string filepath);
        /// <summary>
        /// 获取数据
        /// </summary>
        /// <param name="section">段落名</param>
        /// <param name="key">键名</param>
        /// <param name="def">没有找到时返回的默认值</param>
        /// <param name="filename">ini文件完整路径</param>
        /// <returns></returns>
        public static string getString(string key, string def)
        {
            string filename = AppDomain.CurrentDomain.BaseDirectory + "\\Config.ini";
            StringBuilder sb = new StringBuilder(1024);
            GetPrivateProfileString("Information", key, def, sb, 1024, filename);
            return sb.ToString();
        }

        /// <summary>
        /// 写入数据
        /// </summary>
        /// <param name="section">段落名</param>
        /// <param name="key">键名</param>
        /// <param name="val">写入值</param>
        /// <param name="filename">ini文件完整路径</param>
        public static void writeString(string key, string val)
        {
            string filename = AppDomain.CurrentDomain.BaseDirectory + "\\Config.ini";
            WritePrivateProfileString("Information", key, val, filename);
        }
    }
}
