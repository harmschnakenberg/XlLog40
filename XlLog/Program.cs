using System;
using System.Diagnostics;
using System.IO;

namespace Kreutztraeger
{
    // Strg+M, Strg+O => Alles zuklappen
    // Strg+M, Strg+L => Alles aufklappen

    /*
     *  SETUP: Detail siehe Word-Datei!
     *  \\192.168.160.8\Schnakenberg\lesen\NoMakro\XlLog-Dokumentation.docx
     */

    /// <summary>
    /// Programm zum schreiben von Excel-Tabellen aus InTouch-Variablen.
    /// -Schreibt in Excel-Vorlage Variablenwerte aus InTouch.
    /// -Funktioniert ohne Office-Installation mit EPPlus.dll
    /// -Funktioniert nur, wenn InTouch-Viewer direkt gestartet wird (nicht über Maker).
    /// -Benutzt *.dll aus InTouch zur Verbindung mit InTouch-Datenbank.
    /// -Setzt einen Scheduler-Task, der diese exe immer zur vollen Viertelstunde ausführt.
    /// -Scheduler-Task, startet diese *.exe über ein *.vbs-Skript, um versteckte Ausführung mit dem angemeldeten Benutzer zu ermöglcihen
    /// -PDF-Erzeugung funktioniert nur, wenn ein Programm instlalliert ist, mit dem sich Excel-Dateien öffnen und drucken lassen (Excel, LibreOffice). 
    /// </summary>
    class Program //Fehlernummern siehe Log.cs 01YYZZ
    {

        #region globale Felder
        internal static bool AppErrorOccured = false; // setzt Bits in Intouch bei XlLog-Alarmen.
               
        internal static int AppErrorNumber = -1; // Fehlerkategorie, die nach InTouch gemeldet werden soll.
        internal static string[] CmdArgs;
        internal static string AppStartedBy = "unbekannt"; // Mögliche Werte null, -Task, -Schock, -AlmDruck, -PdfDruck Pfad\Datei.xslx,  

        #region InTouch-TagNames
        internal static string InTouchDiscXlLogFlag = "XlLogUse"; // InTouch TagName, muss != 0 sein, damit diese *.exe weiter ausgeführt wird.
        internal static string InTouchDiscAlarm = "XlLogAlarm"; // Alarmvariable in InTouch zum Trennen von Makro-Alarmen und XlLog-Alarmen.
        internal static string InTouchDiscTimeOut = "XlLogTimeoutBit"; // Bit in Intouch zurücksetzen für Timeout-Umsetzung.
        internal static string InTouchDiscSetCalculations = "ExBestStdWerte"; //Triggert in InTouch die Bildung von Mittelwerten zur vollen Stunde.
        internal static string InTouchDiscResetHourCounter = "ExLöscheStdMin"; //Setzt in Intouch zur voleln Stunde Minutenzähler zurück.
        internal static string InTouchDiscResetQuarterHourCounter = "ExLösche15StdMin"; //Setzt in Intouch zur Viertelstunde Minutenzähler zurück.
        internal static string InTouchDIntErrorNumber = "XlLogErrorNumber"; //An InTouch weiterzugebene Fehlernummer
        #endregion

        #endregion

        static void Main(string[] args) //Fehlernummern siehe Log.cs 0101ZZ
        {
            #region Vorbereitende Abfragen
            try
            {
                CmdArgs = args;

                if (CmdArgs.Length < 1) AppStartedBy = Environment.UserName;
                else AppStartedBy = CmdArgs[0].Remove(0, 1);

                Config.LoadConfig();
                
                Log.Write(Log.Cat.OnStart, Log.Prio.LogAlways, 010101, string.Format("Gestartet durch {0}, Debug {1}, V{2}", AppStartedBy, Log.DebugWord, System.Reflection.Assembly.GetExecutingAssembly().GetName().Version));

                EmbededDLL.LoadDlls();

                bool makerIsRunning = Process.GetProcessesByName("wm").Length != 0;
                if (makerIsRunning)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010102, "Das Programm kann nicht zusammen mit dem InTouch WindowMaker / Manager ausgeführt werden und wird deshalb beendet.");
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht zusammen mit dem InTouch WindowMaker / Manager ausgeführt werden und wird deshalb beendet.");
                    Tools.Wait(10);
                    return;
                }

                bool viewerIsRunning = Process.GetProcessesByName("view").Length != 0;
                if (!viewerIsRunning)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010103, "Das Programm kann nicht ohne den InTouch Viewer ausgeführt werden und wird deshalb beendet.");
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne den InTouch Viewer ausgeführt werden und wird deshalb beendet.");
                    Tools.Wait(10);
                    return;
                }

                if (!File.Exists(NativeMethods.PtaccPath))
                {
                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010104, string.Format("Datei nicht gefunden: " + NativeMethods.PtaccPath));
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + NativeMethods.PtaccPath + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");

                    if (Path.GetDirectoryName(NativeMethods.PtaccPath).Contains(" (x86)"))
                    {
                        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010105, string.Format("Dieses Programm ist für ein 64-Bit Betriebssystem ausgelegt."));
                    }
                    else if (Path.GetDirectoryName(NativeMethods.PtaccPath).StartsWith(@"C:\Program Files\"))
                    {
                        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010106, string.Format("Dieses Programm ist für ein 32-Bit Betriebssystem ausgelegt."));
                    }                    

                    Tools.Wait(10);
                    return;
                }

                if (!File.Exists(NativeMethods.WwheapPath))
                {
                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010107, string.Format("Datei nicht gefunden: " + NativeMethods.WwheapPath));
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + NativeMethods.WwheapPath + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");

                    if (Path.GetDirectoryName(NativeMethods.WwheapPath).Contains(" (x86)"))
                    {
                        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010108, string.Format("Dieses Programm ist für ein 64-Bit Betriebssystem ausgelegt."));
                    }
                    else if (Path.GetDirectoryName(NativeMethods.PtaccPath).StartsWith(@"C:\Program Files\"))
                    {
                        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010109, string.Format("Dieses Programm ist für ein 32-Bit Betriebssystem ausgelegt."));
                    }

                    Tools.Wait(10);
                    return;
                }

                if (!File.Exists(Excel.XlTemplateDayFilePath))
                {
                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010110, string.Format("Vorlage für Tagesdatei nicht gefunden unter: " + Excel.XlTemplateDayFilePath));
                    //AppErrorOccured = true;
                }

                if (!File.Exists(Excel.XlTemplateMonthFilePath))
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 010111, string.Format("Keine Vorlage für Monatsdatei gefunden."));
                    //kein Fehler
                }

                if ((bool)InTouch.ReadTag(InTouchDiscXlLogFlag) != true)
                {
                    Log.Write(Log.Cat.InTouchVar, Log.Prio.Error, 010112, "Freigabe-Flag >" + InTouchDiscXlLogFlag + "< wurde nicht in InTouch gesetzt. Das Programm wird beendet.");
                    //AppErrorOccured = true;
                    return;
                }

                string Operator = (string)InTouch.ReadTag("$Operator");
                Log.Write(Log.Cat.Info, Log.Prio.Info, 010113, "Angemeldet in InTouch: >" + Operator + "<");

                Scheduler.CeckOrCreateTaskScheduler();

                if (!Directory.Exists(Excel.XlArchiveDir))
                {
                    try
                    {
                        Directory.CreateDirectory(Excel.XlArchiveDir);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 010114, string.Format("Archivordner konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", Excel.XlArchiveDir, ex.GetType().ToString(), ex.Message, ex.InnerException));                        
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010115, string.Format("Fehler beim initialisieren der Anwendung: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                return;
            }
            #endregion

            Excel.XlFillWorkbook();

            Print.PrintRoutine();

            #region Diese *.exe beenden   
            InTouch.SetExcelAliveBit(Program.AppErrorOccured);

            if (AppErrorOccured)
            {
                Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010116, "XlLog.exe beendet. Es ist ein Fehler aufgetreten.\r\n\r\n");
            }
            else
            {
                Log.Write(Log.Cat.OnStart, Log.Prio.Info, 010117, "XlLog.exe ohne Fehler beendet.\r\n");
            }

            // Bei manuellem Start Fenster kurz offen halten.
            if (AppStartedBy == Environment.UserName)
            {
                Tools.Wait(Tools.WaitToClose);
            }
            #endregion
        }

    }
}
