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
        internal static bool AlwaysResetTimeoutBit = false; // Setzt das Timeoutbit bei jedem Programmstart.        
        internal static int AppErrorNumber = -1; // Fehlerkategorie, die nach InTouch gemeldet werden soll.        
        internal static string[] CmdArgs; // Übergabeparameter von Programmstart
        internal static string AppStartedBy = "unbekannt"; // Mögliche Werte null, -Task, -Schock, -AlmDruck, -PdfDruck Pfad\Datei.xslx,  

        #region InTouch-TagNames
        internal static string InTouchDiscXlLogFlag = "XlLogUse"; // InTouch TagName, muss != 0 sein, damit diese *.exe weiter ausgeführt wird.
        internal static string InTouchDiscAlarm = "XlLogAlarm"; // Alarmvariable in InTouch zum Trennen von Makro-Alarmen und XlLog-Alarmen.
        internal static string InTouchDiscTimeOut = "XlLogTimeoutBit"; // Bit in Intouch zurücksetzen für Timeout-Umsetzung.
        internal static string InTouchDiscSetCalculations = "XlLogBestStdWerte"; //Triggert in InTouch die Bildung von Mittelwerten zur vollen Stunde.
        internal static string InTouchDiscResetHourCounter = "XlLogLöscheStdMin"; //Setzt in Intouch zur voleln Stunde Minutenzähler zurück.
        internal static string InTouchDiscResetQuarterHourCounter = "XlLogLösche15StdMin"; //Setzt in Intouch zur Viertelstunde Minutenzähler zurück.
        internal static string InTouchDIntErrorNumber = "XlLogErrorNumber"; //An InTouch weiterzugebene Fehlernummer
        #endregion

        #endregion

        static void Main(string[] args) //Fehlernummern siehe Log.cs 0101ZZ
        {
            //Es darf nur eine Instanz des Programms laufen. Freigabe für den erneuten Start des Programms sperren. Eindeutiger Mutex gilt Computer-weit. 
            //Quelle: https://saebamini.com/Allowing-only-one-instance-of-a-C-app-to-run/
            using (var mutex = new System.Threading.Mutex(false, "XlLog"))
            {
                // TimeSpan.Zero to test the mutex's signal state and return immediately without blocking
                bool isAnotherInstanceOpen = !mutex.WaitOne(TimeSpan.Zero);
                if (isAnotherInstanceOpen)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010100, string.Format($"Es darf nur eine Instanz des Programms laufen. " +
                        $"Zweite Instanz, aufgerufen durch {System.Security.Principal.WindowsIdentity.GetCurrent().Name} ({Environment.UserName}) wird beendet."));
                    return;
                }

                CmdArgs = args;
                if (!PrepareProgram(args)) 
                    return; //Initialisierung fehlgeschlagen

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

                mutex.ReleaseMutex(); // Freigabe für den erneuten Start des Programms geben. 
            }
        }


        /// <summary>
        /// Fragt Voraussetzungen zum Ablauf des Programms ab. 
        /// Verzeigt ggf. ab zu PDF-Erstellung
        /// </summary>
        /// <param name="CmdArgs">Bei Programmaufruf übergebene Argumente bzw. Drag&Drop</param>
        /// <returns>true = Programm fortfahren, false = programm beenden</returns>
        private static bool PrepareProgram(string[] CmdArgs) //Fehlernummern siehe Log.cs 0102ZZ
        {
            #region Vorbereitende Abfragen
            try
            {
                if (CmdArgs.Length < 1) AppStartedBy = Environment.UserName;
                else
                {
                    AppStartedBy = CmdArgs[0].Remove(0, 1);
                }
                Config.LoadConfig();
                
                Log.Write(Log.Cat.OnStart, Log.Prio.LogAlways, 010201, string.Format("Gestartet durch {0}, Debug {1}, V{2}", AppStartedBy, Log.DebugWord, System.Reflection.Assembly.GetExecutingAssembly().GetName().Version));

                #region PDF erstellen per Drag&Drop
                try
                {
                    if (CmdArgs.Length > 0)
                    {
                        if (File.Exists(CmdArgs[0]) && Path.GetExtension(CmdArgs[0]) == ".xlsx")
                        {
                            //Wenn der Pfad zu einer Excel-Dateie übergebenen wurde, diese in PDF umwandeln, danach beenden
                            Console.WriteLine("Wandle Excel-Dateie in PDF " + CmdArgs[0]);
                            Log.Write(Log.Cat.PdfWrite, Log.Prio.LogAlways, 010202, "Wandle Excel-Datei in PDF " + CmdArgs[0]);
                            Pdf.CreatePdf(CmdArgs[0]);
                            Console.WriteLine("Exel-Datei " + CmdArgs[0] + " umgewandelt in PDF.\r\nBeliebige Taste drücken zum Beenden...");
                            Console.ReadKey();
                            return false;
                        }
                        else if (!File.Exists(CmdArgs[0]) && Directory.Exists(CmdArgs[0]))
                        {
                            //Alle Excel-Dateien im übergebenen Ordner in PDF umwandeln, danach beenden
                            Console.WriteLine("Wandle alle Excel-Dateien in PDF im Ordner " + CmdArgs[0]);
                            Log.Write(Log.Cat.PdfWrite, Log.Prio.LogAlways, 010203, "Wandle alle Excel-Dateien in PDF im Ordner " + CmdArgs[0]);
                            Pdf.CreatePdf4AllXlsxInDir(CmdArgs[0], false);
                            Console.WriteLine("Exel-Dateien umgewandelt in " + CmdArgs[0] + "\r\nBeliebige Taste drücken zum Beenden...");
                            Console.ReadKey();
                            return false;
                        }
                    }
                }
                catch
                {
                    Log.Write(Log.Cat.PdfWrite, Log.Prio.Error, 010204, string.Format("Fehler beim Erstellen von PDF durch Drag'n'Drop. Aufrufargumente prüfen."));
                }
                #endregion

                EmbededDLL.LoadDlls();

                bool makerIsRunning = Process.GetProcessesByName("wm").Length != 0;
                if (makerIsRunning)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010205, "Das Programm kann nicht zusammen mit dem InTouch WindowMaker / Manager ausgeführt werden und wird deshalb beendet.");
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht zusammen mit dem InTouch WindowMaker / Manager ausgeführt werden und wird deshalb beendet.");
                    Tools.Wait(10);
                    return false;
                }

                bool viewerIsRunning = Process.GetProcessesByName("view").Length != 0;

                if (!viewerIsRunning)
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010206, "Das Programm kann nicht ohne den InTouch Viewer ausgeführt werden und wird deshalb beendet.");
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne den InTouch Viewer ausgeführt werden und wird deshalb beendet.");
                    Tools.Wait(10);
                    return false;
                }

                //_ = Config.SetEnvironmentVariables("PtaccPath", @"C:\Program Files (x86)\Wonderware\InTouch\ptacc.dll");
                //_ = Config.SetEnvironmentVariables("WwheapPath", @"C:\Program Files (x86)\Common Files\ArchestrA\wwheap.dll");

                if (!File.Exists(NativeMethods.PtaccPath))
                {

                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Info, 010207, string.Format("Datei für 64bit-OS nicht gefunden: " + NativeMethods.PtaccPath));

                    if (!File.Exists(NativeMethods32.PtaccPath))
                    {
                        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010208, string.Format("Datei für 64bit oder 32bit-OS nicht gefunden: \r\n" +
                            NativeMethods.PtaccPath + "\r\n" +
                            NativeMethods32.PtaccPath + "\r\n"));
                        Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + Path.GetFileName(NativeMethods32.PtaccPath) + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");
                        Tools.Wait(10);
                        return false;
                    }
                    else
                    {
                        InTouch.Is32BitSystem = true;
                    }
                }
                else
                {
                    InTouch.Is32BitSystem = false;
                }

                if (!File.Exists(NativeMethods.WwheapPath))
                {

                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Info, 010209, string.Format("Datei für 64bit-OS nicht gefunden: " + NativeMethods.WwheapPath));

                    if (!File.Exists(NativeMethods32.WwheapPath))
                    {
                        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010210, string.Format("Datei für 64bit oder 32bit-OS nicht gefunden: \r\n" +
                            NativeMethods.WwheapPath + "\r\n" +
                            NativeMethods32.WwheapPath + "\r\n"));
                        Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + Path.GetFileName(NativeMethods32.WwheapPath) + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");
                        Tools.Wait(10);
                        return false;
                    }
                    else
                    {
                        InTouch.Is32BitSystem = true;
                    }
                }
                else
                {
                    InTouch.Is32BitSystem = false;
                }
   //*/

                if (!(bool)InTouch.ReadTag(Program.InTouchDiscXlLogFlag))
                {
                    Log.Write(Log.Cat.OnStart, Log.Prio.Warning, 010218, string.Format($"Keine Programmfreigabe ({Program.InTouchDiscXlLogFlag}=0)"));
                    return false;
                }

                //if ((bool)InTouch.ReadTag(InTouchDiscXlLogFlag) != true)
                //{
                //    Log.Write(Log.Cat.InTouchVar, Log.Prio.Error, 010214, "Freigabe-Flag >" + InTouchDiscXlLogFlag + "< wurde nicht in InTouch gesetzt. Das Programm wird beendet.");
                //    //AppErrorOccured = true;
                //    return false;
                //}

                //if (!File.Exists(NativeMethods.WwheapPath))
                //{
                //    Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010107, string.Format("Datei nicht gefunden: " + NativeMethods.WwheapPath));
                //    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + NativeMethods.WwheapPath + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");

                //    if (Path.GetDirectoryName(NativeMethods.WwheapPath).Contains(" (x86)"))
                //    {
                //        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010108, string.Format("Dieses Programm ist für ein 64-Bit Betriebssystem ausgelegt."));
                //    }
                //    else if (Path.GetDirectoryName(NativeMethods.PtaccPath).StartsWith(@"C:\Program Files\"))
                //    {
                //        Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010109, string.Format("Dieses Programm ist für ein 32-Bit Betriebssystem ausgelegt."));
                //    }

                //    Tools.Wait(10);
                //    return;
                //}

                if (!File.Exists(Excel.XlTemplateDayFilePath))
                {
                    Log.Write(Log.Cat.InTouchDB, Log.Prio.Error, 010212, string.Format("Vorlage für Tagesdatei nicht gefunden unter: " + Excel.XlTemplateDayFilePath));
                    //AppErrorOccured = true;
                }

                if (!File.Exists(Excel.XlTemplateMonthFilePath))
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 010213, string.Format("Keine Vorlage für Monatsdatei gefunden."));
                    //kein Fehler
                }

                string Operator = (string)InTouch.ReadTag("$Operator");
                Log.Write(Log.Cat.Info, Log.Prio.Info, 010215, "Angemeldet in InTouch: >" + Operator + "<");

                Scheduler.CeckOrCreateTaskScheduler();

                if (!Directory.Exists(Excel.XlArchiveDir))
                {
                    try
                    {
                        Directory.CreateDirectory(Excel.XlArchiveDir);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 010216, string.Format("Archivordner konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", Excel.XlArchiveDir, ex.GetType().ToString(), ex.Message, ex.InnerException));
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.OnStart, Log.Prio.Error, 010217, string.Format("Fehler beim initialisieren der Anwendung: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                return false;
            }
            #endregion
            return true; // Programms oll fortfahren mit Excel-Tabellen füllen
        }


    }
}
