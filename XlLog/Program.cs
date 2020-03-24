using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace Kreutztraeger
{
    // Strg+M, Strg+O => Alles zuklappen
    // Strg+M, Strg+L => Alles aufklappen

    /* ToDo:
     * ok 1) Zeilen Abfragen auf Viertelstundenwerte --> TK-Räume
     * ok 2) Inhalte (Variablennamen) aus erster Zeile löschen.
     * ok 3) Log-File-System einrichten.
     * ok 4) Konfig-Datei implementieren.
     * ok 5) Schockkühler-Dokumentation.
     * ok 6) TaskScheduler automatisch einrichten beim ersten Programmstart.
     * ok 7) Wenn Excel-Mappe bereits geöffnet ist, schließen um schreiben zu ermöglichen. 
     * ~~ 8) Wenn geöffnete Excel-Mappe zwangsweise schließen: Abfrage beim nächsten Öffnen unterdrücken.
     * ok 9) Task kann nur zur vollen Minute gestartet werden (schreiben dauert 1-2 Sekunden!) --> XML-Konfiguration Scheduler? Lösung: siehe 12)
     * ok 10) PDF-Druck; bisher nur möglich mit Excel oder LibreOffice. Lösung: Eigene EXE geschrieben Xl2Pdf.exe (basiert auf Spire.XLS)
     * ok 11) ExcelAktiv Status-Bits in InTouch setzen, um Excel-Timeout-Alarme zu unterdrücken --> Ist nach Anpassung unnötig. Siehe 33)
     * ok 12) Excel: (int)XlOffsetMin Minuten nach Voll (0 / 15 / 30 / 45 Min.) noch zu dem  vorherigen Zeitintervall zählen. (Sonst kommt z.B. 8:01 Uhr in die Zeile 9:00 Uhr)
     * ok 13) EPPlus.dll als Embedded Resource speichern und beim Start entpacken bzw. Setup.exe erstellen, um nur eine (exe)Datei weitergeben zu müssen. 
     * ok 14) XLConfig.ini automatisch erstellen, wenn nicht vorhanden.
     * ok 15) Monatstabelle schreiben
     * ok 16) Vorjahreswert in Monatstabelle schreiben -> Abhänig von Dateinamen Beginnend mit "M".
     * ok 17) NiceToHave: PDF-Druck Alternative über Ghostscript --> PDF-Druck benötigt XLSX-Interpreter. Daher Ghostscript ungeeignet.
     * ok 18) Datum in Exceltabellen beim erstellen schreiben (Benannte Zellen in Workbook und Worksheet erkennen)
     * ok 19) Excel-Alarme testen
     * ok 20) AlarmDBAbfrage erstellen. Aus InTouch PDF-Druck starten mit xlLog.exe -PDFDruck statt  WWExecute("excel", "system", BatCommand); BatMacro = "AlarmDBMakro.xlsm!xlPDFAbfrage"; 
     * ok 21) Druckbereich Schockkühler-Blatt auf tatsächlich gefüllte Zeilenanzahl reduzieren.
     * ok 22) Message-Variablen aus InTouch lesen funktioniert nicht: PtAccReadM(int accid, int hPt, string nm, int nMax); gibt keinen String aus. Lösung StringBuilder statt string einsetzen.
     * ok 23) PDF-Druck: Leerseitendruck vermeiden; Formatierung ausserhalb der Blätter entfernen. --> Formatierung der Vorlage!
     * ok 24) Solange Excel läuft, ist für Aufzeichnung von Tages-/Monatstabellen keine Änderung in InTouch notwendig --> Wenn Excel installiert ist und nicht läuft, Excel starten ohne Mappe. NACHTRAG: Unnötig; besser Sperrvariable für Excel-Alarme in Intouch setzen.
     * ok 25) PDF-Tagesdateien werden nicht ertsellt, wenn ab Folgetag keine Tagesdatei geschrieben wird. -> Es werden PDFs zu allen *.xlsx im Zielordner erstellt, die noch keins haben.
     * ok 26) Sollten LogFiles Älter als 2 Jahre gelöscht werden? Bei DebugWord = 0 sind das ca. 10 kB/d gibt 3,5MB/a gibt ca. 35MB in 10 Jahren --> antwort nein.
     * ok 27) In Monatstabelle, wird nur der gestrige Wert geschrieben. Bei Ausfällen fehlt der letzte aufgezeichnete Tag. => Alle *.xlsx in Monatsordner durchsuchen und in Monatstabelle schreiben. 
     * ok 28) Ausdruck nach Tages- / Monatswechsel auf einstellbaren Drucker.
     * ok 29) C.S. ExBestStdWerte erzeugt bei Verdichter 101,7 % Leistung im Mittelwert, wenn ganze Zeit 100% gefahren wird => D.S. $Minute benötigt bis zu 10 Sek. um Mittelwerte zu bilden; Warte 10 Sek. vor Mittelwertbildung ExBestStdWerte. (siehe XlWriteToDayFile()) 
     * ok 30) Bei Monatswechsel wird der letze Wert nicht mehr in die Monatsdatei geschrieben. Letzte Tagesdatei wird am letzten des Monats um 01:00 Uhr geschrieben, Schreiben der Monatstabelle wird aber schon 00:00 Uhr getriggert.
     * ok 31) In Vorlagedatei Zeile mit TagNames beim neu erstellen Schriftfarbe auf schwarz setzen. Dadurch können in den Vorlagedateien weiße Schrift auf weißem Grund genutzt werden.
     * ok 32) Setzte ExLöscheStdMin bzw. ExLösche15StdMin von hier aus. --> siehe XlWriteToDayFile()
     * ok 33) von Henry: Alte Excel-Überwachung (ExcelActiv, ExcelMacro) und Überwachung für XlLog in Intouch trennen. --> Alarm- und Sperrvariable in InTouch dazu.
     * ok 34) von Henry: PrintStartHour Standardwert von 3 Uhr auf 4 Uhr setzen, um Fehler durch Zeitumstellung zu vermeiden.
     * ok 35) beim Lesen aus der XlConfig.ini wurden Umlaute nicht richtig erkannt. Lösung: Beim Lesen und Schreiben der Datei explizit Encoding.UTF8 erzwingen.
     * ok 36) Tagesweise LogFiles sind viel Klickerei. Deshalb auf Monats-Logfiles umgestellt.
     * ~~ 37) Vorlagedateien aus älteren Excel-Versionen als 2016 können korrupte *.xlsx-Dateien erzeugen, die sich zwar öffnen lassen, aber in Excel einen Fehlerdialog triggern. Das fürht zu Abstürzen, wenn Zugriff aus dem Programm erfolgt. Lösung bisher: Vorlagedateien neu erstellen mit Excel 2016 oder höher.
     * ok 38) In der XlConfig.ini unter [Druck] wird PrintFileExtention= rausgeschmissen: Alle Druckprogramme nehmen mittlerweile *.pdf an.
     * ok 39) Wenn in XlConfig.ini eingestellt ist XlPosOffsetMin=30 und XlNegOffsetMin=30, kann ein Eintrag bei Stunde 25 (Zeile unter 00:00 Uhr) erfolgen. --> Abgefangen in XlSetRowAndCol(): Stunde > 24 wird zu Stunde 0.
     * 49) Monatsdatei: zweiter Wert in einem Blatt wird eine Spalte zu früh geschrieben.
     * ok 50) Wunsch: In der XlConfig.ini unter [PDF] kann mit der Zeit ConvertToPdfHour die Stunde festgelegt werden, in der PDFs automatisch erstellt werden.
     * ok 51) Da der Pfad zu InTouch-DLLs als const angegeben werden muss, sind unterschiedliche Kompilate für 32-Bit OS und 64-Bit OS notwendig.
     * ok 52) Wird XlLog nicht am ersten Tag eines Monats ausgeführt, werden die letzten PDF-Dateien im Ordner des Vormonats nicht geschrieben. --> Zusätzlich wird auch der Vormonat auf zu schreibende PDF untersucht.
     * ok 53) Die Zeit, die XlLog nach manuellem Start wartet soll einstellbar sein -> XlConfig.ini EIntrag WaitToClose
     */

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
    class Program
    {

        #region globale Felder
        internal static bool AppErrorOccured = false; // setzt Bits in Intouch bei XlLog-Alarmen.
        internal static string[] CmdArgs;

            #region InTouch-TagNames
            internal static string XlLogFlag = "XlLogUse"; // InTouch TagName, muss != 0 sein, damit diese *.exe weiter ausgeführt wird.
            internal static string InTouchDiscAlarm = "XlLogAlarm"; // Alarmvariable in InTouch zum Trennen von Makro-Alarmen und XlLog-Alarmen.
            internal static string InTouchDiscTimeOut = "XlLogTimeoutBit"; // Bit in Intouch zurücksetzen für Timeout-Umsetzung.
            internal static string AppStartedBy = "unbekannt"; // Mögliche Werte null, -Task, -Schock, -AlmDruck, -PdfDruck Pfad\Datei.xslx,    
            #endregion
       
        #endregion

        static void Main(string[] args)
        {
            #region Vorbereitende Abfragen
            try
            {
                CmdArgs = args;

                if (CmdArgs.Length < 1) AppStartedBy = Environment.UserName;
                else AppStartedBy = CmdArgs[0].Remove(0, 1);

                Config.LoadConfig();

                Log.Write(Log.Category.OnStart, -000000001, string.Format("Gestartet durch {0}, Debug {1}", AppStartedBy, Log.DebugWord));

                EmbededDLL.LoadDlls();

                bool makerIsRunning = Process.GetProcessesByName("wm").Length != 0;
                if (makerIsRunning)
                {
                    Log.Write(Log.Category.InTouchDB, -902011249, "Das Programm kann nicht zusammen mit dem InTouch WindowMaker / Manager ausgeführt werden und wird deshalb beendet.");
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht zusammen mit dem InTouch WindowMaker / Manager ausgeführt werden und wird deshalb beendet.");
                    Tools.Wait(10);
                    return;
                }

                bool viewerIsRunning = Process.GetProcessesByName("view").Length != 0;
                if (!viewerIsRunning)
                {
                    Log.Write(Log.Category.InTouchDB, -902040738, "Das Programm kann nicht ohne den InTouch Viewer ausgeführt werden und wird deshalb beendet.");
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne den InTouch Viewer ausgeführt werden und wird deshalb beendet.");
                    Tools.Wait(10);
                    return;
                }

                if (!File.Exists(NativeMethods.PtaccPath))
                {
                    Log.Write(Log.Category.InTouchDB, -902011253, string.Format("Datei nicht gefunden: " + NativeMethods.PtaccPath));
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + NativeMethods.PtaccPath + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");

                    if (Path.GetDirectoryName(NativeMethods.PtaccPath).Contains(" (x86)"))
                    {
                        Log.Write(Log.Category.InTouchDB, -002141408, string.Format("Dieses Programm ist für ein 64-Bit Betriebssystem ausgelegt."));
                    }
                    else
                    {
                        Log.Write(Log.Category.InTouchDB, -002141409, string.Format("Dieses Programm ist für ein 32-Bit Betriebssystem ausgelegt."));
                    }

                    Tools.Wait(10);
                    return;
                }

                if (!File.Exists(NativeMethods.WwheapPath))
                {
                    Log.Write(Log.Category.InTouchDB, -902011254, string.Format("Datei nicht gefunden: " + NativeMethods.WwheapPath));
                    Console.WriteLine("ACHTUNG: Das Programm kann nicht ohne die Datei " + NativeMethods.WwheapPath + " ausgeführt werden und wird deshalb beendet. Beachte Log-Datei.");

                    if (Path.GetDirectoryName(NativeMethods.WwheapPath).Contains(" (x86)"))
                    {
                        Log.Write(Log.Category.InTouchDB, -002141410, string.Format("Dieses Programm ist für ein 64-Bit Betriebssystem ausgelegt."));
                    }
                    else
                    {
                        Log.Write(Log.Category.InTouchDB, -002141411, string.Format("Dieses Programm ist für ein 32-Bit Betriebssystem ausgelegt."));
                    }

                    Tools.Wait(10);
                    return;
                }

                if (!File.Exists(Excel.XlTemplateDayFilePath))
                {
                    Log.Write(Log.Category.InTouchDB, -001301523, string.Format("Vorlage für Tagesdatei nicht gefunden unter: " + Excel.XlTemplateDayFilePath));
                    AppErrorOccured = true;
                }

                if (!File.Exists(Excel.XlTemplateMonthFilePath))
                {
                    Log.Write(Log.Category.ExcelRead, 2001301524, string.Format("Keine Vorlage für Monatsdatei gefunden."));
                    //kein Fehler
                }

                if ((bool)InTouch.ReadTag(XlLogFlag) != true)
                {
                    Log.Write(Log.Category.InTouchVar, -001301527, "Freigabe-Flag >" + XlLogFlag + "< wurde nicht in InTouch gesetzt. Das Programm wird beendet.");
                    AppErrorOccured = true;
                    return;
                }

                string Operator = (string)InTouch.ReadTag("$Operator");
                Log.Write(Log.Category.Info, 1911221212, "Angemeldet in InTouch: >" + Operator + "<");

                Scheduler.CeckOrCreateTaskScheduler();

                if (!Directory.Exists(Excel.XlArchiveDir))
                {
                    try
                    {
                        Directory.CreateDirectory(Excel.XlArchiveDir);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(Log.Category.FileSystem, -902060750, string.Format("Archivordner konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", Excel.XlArchiveDir, ex.GetType().ToString(), ex.Message, ex.InnerException));
                        InTouch.SetExcelAliveBit(true);
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.OnStart, -902060751, string.Format("Fehler beim initialisieren der Anwendung: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                return;
            }
            #endregion

            Excel.XlFillWorkbook();

            Print.PrintRoutine();

            #region Diese *.exe beenden
            InTouch.SetExcelAliveBit(AppErrorOccured);

            if (AppErrorOccured)
            {
                Log.Write(Log.Category.OnStart, -907261431, "Excel-Aufzeichnung beendet. Es ist ein Fehler aufgetreten.\r\n\r\n");
            }
            else
            {
                Log.Write(Log.Category.OnStart, 1902011311, "Excel-Aufzeichnung ohne Fehler beendet.\r\n");
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
