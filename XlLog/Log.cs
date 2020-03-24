using System;
using System.IO;
using System.Reflection;


namespace Kreutztraeger
{

    public class Log
    {
        public static int DebugWord { get; set; } = 0;
        /* Debug Bits in DebugWord
         *      MethodCall
         * 0    Info
         * 1    OnStart
         * 2    InTouchDB
         * 3    InTouchVar
         * 4    FileSystem
         * 5    ExcelRead
         * 6    ExcelWrite
         * 7    ExcelShock
         * 8    PdfWrite
         * 9    Print
         * 10   Scheduler
         * 11   SqlQuery
         * 
         */

        public enum Category : int { Info, OnStart, InTouchDB, InTouchVar, FileSystem, ExcelRead, ExcelWrite, ExcelShock, PdfWrite, Print, Scheduler, SqlQuery, MethodCall };

        /// <summary>
        /// Schreibt in eine Log-Datei wenn DebugBitPos in DegbuByte gesetzt ist oder errorNo < 0
        /// </summary>
        /// <param name="DebugBitPos">Bit-Position in DebugByte</param>
        /// <param name="errorNo">Eindeutige Fehlernummer</param>
        /// <param name="logMessage">Text</param>
        public static void Write(Category DebugBitPos, int errorNo, string logMessage)
        {
            if (Tools.IsBitSet(DebugWord, (int)DebugBitPos) || errorNo < 0)
            {
                LogWrite(string.Format("{0:00}-{1:D9}", (int)DebugBitPos, errorNo), logMessage);
            }

            if (DebugBitPos == (int)Log.Category.Info && Tools.IsBitSet(DebugWord, (int)Log.Category.Info))
            {
                Console.WriteLine(logMessage);
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="errorNo">Eindeutige Fehlernummer zur Wiedererkennung bei Fehlersuche im Quelltext. Format YYmmDDHHMM</param>
        /// <param name="logMessage">Zu loggender Text</param>
        private static void LogWrite(string errorNo, string logMessage)
        {
            string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Log");
            logPath = Path.Combine(logPath, DateTime.Now.Year.ToString());
            Directory.CreateDirectory(logPath);
            //logPath = Path.Combine(logPath, DateTime.Now.Month.ToString("00"));
            //Directory.CreateDirectory(logPath);
            logPath = Path.Combine(logPath, string.Format("Log_{0}.txt", DateTime.Now.ToString("yyyyMM")));

            using (StreamWriter w = File.AppendText(logPath))
            {
                try
                {
                    w.WriteLine("{0:00} {1} {2}\t{3}", DateTime.Now.Day, DateTime.Now.ToLongTimeString(), errorNo, logMessage);

                }
                catch (Exception ex)
                {
                    Console.WriteLine("FEHLER Logging:");
                    Console.WriteLine("Typ:\t\t" + ex.GetType().ToString());
                    Console.WriteLine("Message:\t" + ex.Message);
                    Console.WriteLine("InnerException:\t" + ex.InnerException);
                    Console.WriteLine("StackTrace:\t" + ex.StackTrace);
                }
            }

        }

    }

}
