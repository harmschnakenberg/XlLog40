using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Kreutztraeger
{
    class Print
    {
        private const string lastPrintLogFileName = "LetzterAusdruckT.txt";
        private const int hoursBetweenPrints = 24;

        private static int printStartHour = 4;
        internal static int PrintStartHour { get => printStartHour; set => printStartHour = value; }

        private static string printAppPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "XlsxPrint.exe");
        internal static string PrintAppPath { get => printAppPath; set => printAppPath = value; }

        internal static string PrinterAppArgs { get; set; } = "*Quelle*";

        /// <summary>
        /// Schreibt das Aktuelle Datum mit der Uhrzeit 'PrintStartHour' in eine Textdatei im Stammordner als Referenz für den letzten Druckzeitpunkt.
        /// </summary>
        /// <param name="LastPrintLogFileName">Name der Textdatei, in die geschrieebn werdne soll.</param>
        private static void WriteLastPrintLog(string LastPrintLogFileName)
        {
            Log.Write(Log.Category.MethodCall, 1907261355, string.Format("WriteLastPrintLog({0})", LastPrintLogFileName));

            string printerLogPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), LastPrintLogFileName);

            try
            {
                DateTime printTime = DateTime.Now.AddHours(PrintStartHour - DateTime.Now.Hour).AddSeconds(-DateTime.Now.Second); // sollte 0 Min und 0 Sekunden schreiben
                File.WriteAllText(printerLogPath, printTime.ToString("G"));
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.FileSystem, -906271521, string.Format("Die Datei zur Dokumentation des letzten Ausdruckks konnte nicht geschrieben werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", printerLogPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                Console.WriteLine("FEHLER beim Erstellen von {0}. Siehe Log.", printerLogPath);
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Liest den letzten Druckzeitpunkt aus der Datei 'LastPrintLogFileName' im Stammverzeichnis.
        /// Kann keine DateTime ermittelt werden, wird DatTime.Now - 24 Std. (gestern um die Zeit) ausgegeben.
        /// </summary>
        /// <param name="LastPrintLogFileName"></param>
        /// <returns></returns>
        private static DateTime ReadLastPrintLog(string LastPrintLogFileName)
        {
            Log.Write(Log.Category.MethodCall, 1907261356, string.Format("ReadLastPrintLog({0})", LastPrintLogFileName));

            string printerLogPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), LastPrintLogFileName);

            try
            {
                string SLastPrintTime = File.ReadAllText(printerLogPath);
                if (DateTime.TryParse(SLastPrintTime, out DateTime lastPrintTime))
                {
                    return lastPrintTime;
                }
                else
                {
                    //Textdatei nicht lesbar. Schreibe neue Datei:
                    WriteLastPrintLog(LastPrintLogFileName);
                    Log.Write(Log.Category.Print, 1906271539, string.Format("Die Datei zur Dokumentation des letzten Ausdruckks konnte nicht gelesen werden und wird neu erstellt."));

                    // Tut so, als ob vor 24 Std. das letzte Mal gedruckt wurde
                    return DateTime.Now.AddDays(-1);
                }
            }
            catch
            {
                //Textdatei nicht lesbar. Schreibe neue Datei:
                WriteLastPrintLog(LastPrintLogFileName);
                Log.Write(Log.Category.Print, 1907011036, string.Format("Die Datei zur Dokumentation des letzten Ausdruckks konnte nicht gelesen werden und wird neu erstellt."));
                return DateTime.Now;
            }
        }

        /// <summary>
        /// Startet das Hilfsprogramm 'PrintAppPath' xlSourcePath BitMaskSheets um Blätter aus der Datei zu drucken.
        /// Die Spezifikation und Ansprache des Druckers erfolgt im Hilfsprogramm.
        /// </summary>
        /// <param name="xlSourcePath"></param>
        /// <param name="BitMaskSheets"></param>
        public static void PrintReport(string xlSourcePath, int BitMaskSheets)
        {
            Log.Write(Log.Category.MethodCall, 1907261357, string.Format("PrintReport({0},{1})", xlSourcePath, BitMaskSheets));

            if (BitMaskSheets == 0) return; // Keine Blätter zum Druck ausgewählt

            //Alle PrintApps nehmen *.pdf-Dateien an - PdfToPrinter.exe nimmt nur *.pdf an.
            xlSourcePath = Path.ChangeExtension(xlSourcePath, ".pdf");
            if (!File.Exists(xlSourcePath))
            {
                Log.Write(Log.Category.Print, 1907241217, string.Format("Die Datei {0} konnte nicht gefunden und deshalb nicht gedruckt werden.", xlSourcePath));
                return;
            }
            else
            {
                Log.Write(Log.Category.Print, 1907241334, string.Format("Info: Die Datei {0} wird gedruckt.", xlSourcePath));
            }

            //Für Monatsausdrucke eigene Datei erstellen.
            if (Path.GetFileNameWithoutExtension(xlSourcePath).Contains("M")) lastPrintLogFileName.Replace('T', 'M');

            DateTime lastPrint = ReadLastPrintLog(lastPrintLogFileName);
            if (DateTime.Now.AddHours(-hoursBetweenPrints).CompareTo(lastPrint) < 0)
            {
                Log.Write(Log.Category.Print, 1907011102, string.Format("Die Zeit hoursBetweenPrints ist noch nicht abgelaufen. Angefragter Druck: {0} <-> Letzer Druck: {1}", DateTime.Now.AddHours(-hoursBetweenPrints), lastPrint));
                // Die Zeit hoursBetweenPrints ist noch nicht abgelaufen.
                return;
            }

            string pagesToPrint = PrintPageSelection(BitMaskSheets);
            Log.Write(Log.Category.Print, 1906271600, string.Format("Starte Druck mit {0} Seiten {1} für {2}\tLetzter Ausdruck war am {3} Uhr.", Path.GetFileName(PrintAppPath), pagesToPrint, Path.GetFileName(xlSourcePath), lastPrint));
            WriteLastPrintLog(lastPrintLogFileName);



            string printJobArgs = PrinterAppArgs.Replace("*Quelle*", xlSourcePath);
            printJobArgs = printJobArgs.Replace("*Seiten*", pagesToPrint);

            ProcessStartInfo startInfo = new ProcessStartInfo
            {
                CreateNoWindow = true,
                UseShellExecute = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                FileName = PrintAppPath,
                Arguments = printJobArgs
            };

            try
            {
                using (Process exeProcess = Process.Start(startInfo))
                {
                    int secondsToPrint = 30;
                    if (!exeProcess.WaitForExit(secondsToPrint * 1000))
                    {
                        Log.Write(Log.Category.Print, -910301253, string.Format("Der Druckauftrag konnte nicht in der vorgegebenen Zeit von {0} sec. abgeschlossen werden.", secondsToPrint) );
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.Print, -906271610, string.Format("Das Hilfsprogramm zum Drucken von Excel-Dateien konnte nicht ordnungsgemäß ausgeführt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", PrintAppPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Erzeugt aus einer Bitmaske eine für Drucker lesbare Seitenauswahl.
        /// </summary>
        /// <param name="BitMaskSheets">Bitmaske aus InTouch-Tag 'ExT_Druck' oder 'ExM_Druck'</param>
        /// <returns>Seitenauswahl für Drucker z.B '1,3,4,5,8'</returns>
        internal static string PrintPageSelection(int BitMaskSheets)
        {
            Log.Write(Log.Category.MethodCall, 1907261358, string.Format("PrintReport({0})", BitMaskSheets));

            string pagesSelection = "";

            for (int i = 0; i < 32; i++)
            {
                if (((BitMaskSheets >> i) & 1) == 1)
                {
                    //Bitmaske startet bei 0, Excel-Seiten startet bei 1
                    pagesSelection += string.Format("{0},", i + 1);
                }
            }

            return pagesSelection.Substring(0, pagesSelection.Length - 1);
        }

        /// <summary>
        /// Druckaufruf aus Hauptprogramm.
        /// </summary>
        internal static void PrintRoutine()
        {
            Log.Write(Log.Category.MethodCall, 1907261359, string.Format("PrintRoutine()"));

            string file = Excel.CeateXlFilePath(-1);
            int BitMaskSheets = (int)InTouch.ReadTag("ExT_Druck");
            //Console.WriteLine("Prüfe Druck " + file + " " + BitMaskSheets);
            Print.PrintReport(file, BitMaskSheets);

            if (DateTime.Now.Day == 1)
            {
                file = Excel.CeateXlFilePath(-1, true);
                BitMaskSheets = (int)InTouch.ReadTag("ExM_Druck");
                //Console.WriteLine("Prüfe Druck " + file + " " + BitMaskSheets);
                Print.PrintReport(file, BitMaskSheets);
            }
        }
    }
}
