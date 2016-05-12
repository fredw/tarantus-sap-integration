using System;
using System.IO;

namespace TarantusSAPService
{
    public static class LogWriter
    {
        /// <summary>
        /// </summary>
        /// <param name="Message"></param>
        public static void write(string Message)
        {
            StreamWriter sw = null;
            DateTime now = DateTime.Now;
            String path = AppDomain.CurrentDomain.BaseDirectory + "\\logs\\" + now.ToString("yyyy-MM") + "\\";
            try {
                // Create log directory if don`t exist
                if (!Directory.Exists(path)) {
                    Directory.CreateDirectory(path);
                }
                // Create log file
                sw = new StreamWriter(path + "\\" + now.ToString("yyyy-MM-dd") + ".log", true);
                sw.WriteLine(now.ToString("HH:mm:ss") + ": " + Message);
                sw.Flush();
                sw.Close();
            } catch { }            
        }
    }
}
