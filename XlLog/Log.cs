using System;
using System.IO;
using System.Reflection;

namespace Kreutztraeger
{

    public static class Log //Fehlernummern siehe Log.cs 07YYZZ
    {
        /* Eindeutige Fehlernummerierung
         * 
         * -- Format XXYYZZ --
         * XX:  Lfd. Nr. der Klasse 
         *      01: Program
         *      02: Config
         *      03: EmbededDLL
         *      04: Excel
         *      05: Intouch
         *      06: NativeMethods
         *      07: Log
         *      08: Pdf
         *      09: Print    
         *      10: Scheduler
         *      11: Sql
         *      12: SetNewSystemTime
         *      13: Tools
         * 
         * YY:  Lfd. Nr. der Methode in der Klasse
         * 
         * ZZ:  Lfd. Nr. des Fehlers in der Methode
         */

        public static int DebugWord { get; set; } = 0;
        /* Debug Bits in DebugWord
         *      
         * 0    Info            Allgemeine Information zum Programmablauf
         * 1    OnStart         Routineprüfungen beim Start
         * 2    InTouchDB       Verbindungsprobleme zu InTouch
         * 3    InTouchVar      Variablen nicht erreichbar
         * 4    FileSystem      Dateisystem
         * 5    ExcelRead       Lesen aus Excel-Datei
         * 6    ExcelWrite      Schreiben in Excel-Datei
         * 7    ExcelShock      Schockkühler
         * 8    PdfWrite        PDF-erzeugen
         * 9    Print           Drucken
         * 10   Scheduler       TaskScheduler
         * 11   SqlQuery        SQL-Datenbankabfrage
         * 12   MethodCall      Bei Aufruf einer Methode in diesem Programm
         */

        public enum Prio : int { Error, Warning, LogAlways, Info};
        /* Typ       InTouchAlarm    Logging
         * Error:       x               x   
         * Warning      0               x
         * LogAlways    0               x
         * Info         0              (x) wenn Bit gesetzt         
         */

        public enum Cat : int { Info, OnStart, InTouchDB, InTouchVar, FileSystem, ExcelRead, ExcelWrite, ExcelShock, PdfWrite, Print, Scheduler, SqlQuery, MethodCall };

        /// <summary>
        /// Schreibt in eine Log-Datei wenn DebugBitPos in DegbuByte gesetzt ist oder errorNo < 0
        /// </summary>
        /// <param name="DebugBitPos">Bit-Position in DebugByte</param>
        /// <param name="errorNo">Eindeutige Fehlernummer</param>
        /// <param name="logMessage">Text</param>
        public static void Write(Cat DebugBitPos, Prio prio, int errorNo, string logMessage) //Fehlernummern siehe Log.cs 0701ZZ
        {
            if (Tools.IsBitSet(DebugWord, (int)DebugBitPos) || prio != Prio.Info)
            {
                LogWrite(string.Format("{0:00} {1} {2:D6}", (int)DebugBitPos, prio.ToString().Substring(0,1), errorNo), logMessage);
            }

            if (DebugBitPos == (int)Log.Cat.Info && Tools.IsBitSet(DebugWord, (int)Log.Cat.Info))
            {
                Console.WriteLine(logMessage);
            }

            if (prio == Prio.Error)
            {
                Program.AppErrorOccured = true;
                
                if (Program.AppErrorNumber < 0) //nur, wenn AppErrorCategory noch nicht gesetzt ist
                {
                    Program.AppErrorNumber = errorNo;//(int)DebugBitPos;
                }

            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="errorNo">Eindeutige Fehlernummer zur Wiedererkennung bei Fehlersuche im Quelltext. Format YYmmDDHHMM</param>
        /// <param name="logMessage">Zu loggender Text</param>
        private static void LogWrite(string errorNo, string logMessage) //Fehlernummern siehe Log.cs 0702ZZ
        {
            string logPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Log");
            //logPath = Path.Combine(logPath, DateTime.Now.Year.ToString()); //11.09.2020 Jahresordner weglassen
            Directory.CreateDirectory(logPath);
            //logPath = Path.Combine(logPath, DateTime.Now.Month.ToString("00"));
            //Directory.CreateDirectory(logPath);
            logPath = Path.Combine(logPath, string.Format("Log_{0}.txt", DateTime.Now.ToString("yyyy_MM")));

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
