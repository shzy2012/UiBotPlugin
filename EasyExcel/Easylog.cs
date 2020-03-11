namespace Easylog
{
    using System;
    using System.IO;
    using System.Linq;
    using System.Text;

    public class Log
    {
        private writer w = null;
        public Log()
        {
            w = new writer();
        }

        public void Info(object message)
        {
            w.Writer("[Info]: " + message);
        }
        public void Info(string message)
        {
            try
            {
                w.Writer("[Info]:" + message);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void Error(Exception exception)
        {
            try
            {
                w.Writer("[Error]:", exception);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        public void Error(object message, Exception exception)
        {
            try
            {
                w.Writer("[Error]:" + message, exception);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
    }

    public class writer
    {
        private static string logPath = "log";

        private static string logFileName = "easy_excel.txt";

        private static long maximumFileSize = 200; //100kb

        private static int maxSizeRollBackups = 10000000;

        private static string currentApplicationPath = AppDomain.CurrentDomain.BaseDirectory;

        public writer() { }

        public void Writer(object message)
        {
            StringBuilder logmessage = new StringBuilder();
            logmessage.AppendLine(string.Format("[{0}]{1}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), message ?? string.Empty));
            var fullpath = checkRule();
            File.AppendAllText(fullpath, logmessage.ToString(), Encoding.UTF8);
        }

        public void Writer(object message, Exception ex)
        {
            StringBuilder logmessage = new StringBuilder();
            logmessage.AppendLine(string.Format("[{0}]{1}{2}", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"), message ?? string.Empty, " [Exception: " + ex.Message + "]"));
            logmessage.AppendLine(ex.StackTrace);
            logmessage.AppendLine(ex.Message);
            logmessage.AppendLine(ex.Source);
            if (ex.InnerException != null)
            {
                logmessage.AppendLine(ex.InnerException.StackTrace);
                logmessage.AppendLine(ex.InnerException.Message);
                logmessage.AppendLine(ex.InnerException.Source);
            }

            var fullpath = checkRule();
            File.AppendAllText(fullpath, logmessage.ToString(), Encoding.UTF8);

        }

        private string checkRule()
        {
            var outputpath = Path.Combine(currentApplicationPath, logPath);
            if (!Directory.Exists(outputpath))
            {
                Directory.CreateDirectory(outputpath);
            }

            var fullpath = Path.Combine(outputpath, logFileName);
            if (!File.Exists(fullpath))
            {
                return fullpath;
            }

            FileInfo fileinfo = new FileInfo(fullpath);
            if (fileinfo.Length / 1024 < maximumFileSize)
            {
                return fullpath;
            }

            var directory = new DirectoryInfo(outputpath);
            var filecount = directory.GetFiles().Length;
            if (filecount > maxSizeRollBackups)
            {
                var newest = directory.GetFiles().OrderBy(x => x.LastWriteTime).FirstOrDefault();
                if (newest != null)
                {
                    File.Copy(fullpath, newest.FullName, true); return fullpath;
                }
            }

            File.Move(fullpath, (fullpath + (filecount++).ToString())); return fullpath;
        }
    }
}
