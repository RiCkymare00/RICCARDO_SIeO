using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppTemp
{
    public class LogWriter
    {
        private string m_exePath = string.Empty;
        public LogWriter(string logMessage)
        {
            LogWrite(logMessage);
        }
        public void LogWrite(string logMessage)
        {
            m_exePath = AppDomain.CurrentDomain.BaseDirectory + "log";
            try
            {
                if (!System.IO.Directory.Exists(m_exePath))
                {
                    _ = System.IO.Directory.CreateDirectory(m_exePath);
                }

                string logFilePath = String.Format(m_exePath + "\\log_{0}.txt", DateTime.Today.ToString("yyyyMMdd"));
                bool fileExists = File.Exists(logFilePath);

                using StreamWriter w = File.AppendText(String.Format(m_exePath + "\\log_{0}.txt", DateTime.Today.ToString("yyyyMMdd")));
                if (!fileExists)
                {
                    w.WriteLine("Inizio dell'attività di logging...");
                }
                Log(logMessage, w);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public void Log(string logMessage, TextWriter txtWriter)
        {
            try
            {
                txtWriter.Write("\r\nLog Entry : ");
                txtWriter.WriteLine("{0} {1}", DateTime.Now.ToLongTimeString(), DateTime.Now.ToLongDateString());
                txtWriter.WriteLine("  :");
                txtWriter.WriteLine("  :{0}", logMessage);
                txtWriter.WriteLine("-------------------------------");
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }
}
