using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Logging_Framework
{
    public class Logger
    {
        private static Logger _logger = null;
        private string logPath;
        private string fileName;
        private LogLevels logLvl;
        private ConcurrentQueue<LoggerMessages> cq = new ConcurrentQueue<LoggerMessages>();

        private Logger(string logPath, string FileName, LogLevels logLevel)
        {
            this.logPath = logPath;
            this.fileName = FileName;
            this.logLvl = logLevel;
            CreateDirectory(logPath);
            Thread loggingThread = new Thread(() =>
            {
                RunLoggerThread();
            });
            loggingThread.Start();
            cq.Enqueue(new LoggerMessages() { LogLevel = LogLevels.Debug, data = "Logger Instantiated." });
        }
        private static void CreateDirectory(string path)
        {
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
        }
        public static Logger GetLogger(string LogPath, string FileName, LogLevels LogLevel)
        {
            if (_logger == null)
            {
                _logger = new Logger(LogPath, FileName, LogLevel);
            }
            else
                _logger.logLvl = LogLevel;
            return _logger;
        }

        private void RunLoggerThread()
        {
            while (true)
            {
                try
                {
                    
                    LoggerMessages loggerMessages = null;

                    if (cq.TryDequeue(out loggerMessages))
                    {
                        if (loggerMessages.LogLevel <= logLvl)
                        {
                            using (System.IO.StreamWriter file =
                                           new System.IO.StreamWriter(logPath + "\\" + fileName + DateTime.Now.ToString("yyyyMMdd") + ".txt", true))
                            {
                                file.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + " " + loggerMessages.LogLevel.ToString() + ":" + loggerMessages.data);
                            }
                        }
                    }
                    else
                    {
                        Thread.Sleep(1000);
                    }
                }
                catch (Exception ex)
                {
                    Thread.Sleep(1000);
                    using (System.IO.StreamWriter file =
                                     new System.IO.StreamWriter(logPath + "\\" + fileName + DateTime.Now.ToString("yyyyMMdd") + ".txt", true))
                    {
                        file.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  ERROR:" + ex.Message);
                        file.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "  ERROR:" + ex.StackTrace);
                    }

                }
            }
        }

        public void Log(LogLevels LogLevel, string data)
        {
            cq.Enqueue(new LoggerMessages() { LogLevel = LogLevel, data = data });
        }
    }

  
    public enum LogLevels{
        Error,
        Info,
        Debug
    }

}
