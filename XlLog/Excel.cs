using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;
using System.Data;
using Microsoft.Win32;
using System.Windows.Media.Effects;
using System.Drawing;
using OfficeOpenXml.Style;
//Fehlernummer 300 bis 399
namespace Kreutztraeger
{
    enum XlShockCol // Spaltennummer in Excel (erste Spalte = 1)
    {
        ChargeNo = 1,
        ShockNo = 2,    // Schockkühlernummer
        StartTime = 3,
        EndTime = 4,
        ShockDuration = 5,
        ShockProgramNo = 6,
        MaxDuration = 7,
        ProductStartTemp = 8,
        ProductEndTemp = 9,
        ProductMinTemp = 10,
        ProductMaxTemp = 11,
        RoomStartTemp = 12,
        RoomEndTemp = 13,
        RoomMinTemp = 14,
        RoomMaxTemp = 15
    };


    internal static class Excel //Fehlernummern siehe Log.cs 04YYZZ
    {
        static readonly string AppDir = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);

        #region Fields for WriteToExcelFiles
        // Minuten, die zwischen XlPosOffsetMin und 60 - XlNegOffsetMin liegen, werden nicht für Stundenwerte aufgezeichnet.
        private static int xlPosOffsetMin = 2; //min. nach Voller (Viertel-)Stunde, die noch zu der vorherigen (Viertel-)Stunde zählen.
        public static int XlPosOffsetMin { get => xlPosOffsetMin; set => xlPosOffsetMin = value; }

        private static int xlNegOffsetMin = 2; //min. vor Voller Stunde, die zu der kommenden vollen Stunde zählen.
        public static int XlNegOffsetMin { get => xlNegOffsetMin; set => xlNegOffsetMin = value; }

        public static string XlTemplateDayFilePath { get; set; } = Path.Combine(AppDir, "T_vorl.xlsx");
        public static string XlTemplateMonthFilePath { get; set; } = Path.Combine(AppDir, "M_vorl.xlsx");
        public static int XlDayFileFirstRowToWrite { get; set; } = 10;
        public static int XlMonthFileFirstRowToWrite { get; set; } = 8;
        public static string XlArchiveDir { get; set; } = @"D:\Archiv";
        internal static string XlPassword { get; set; } = string.Empty;
        internal static string XlPasswordEncrypted { get; set;}

        internal static bool XlImmediatelyCreatePdf { get; set; } = false;

        private static readonly string yellow = Color.Yellow.ToArgb().ToString("X2");
        private static readonly string white = Color.White.ToArgb().ToString("X2");

        #endregion

        #region WriteToExcelFiles

        /// <summary>
        /// Schreibt die nächste Zeile in Excel-Datei
        /// </summary>
        public static void XlFillWorkbook() //Fehlernummern siehe Log.cs 0401ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040101, string.Format("XlFillWorkbook()"));

            string xlDayFilePath = CeateXlFilePath();
            string xlYesterdayFilePath = CeateXlFilePath(-1);
            string xlMonthFilePath = CeateXlFilePath(-1, true);
            string xlMonthLastYearFilePath = CeateXlFilePath(-1, true, -1);

            try
            {
                #region Schließe Excel-Mappen, wenn sie geöffnet sind.
                Process[] processes = Process.GetProcessesByName("EXCEL");

                foreach (var process in processes)
                {
                    KillExcel(process, xlDayFilePath);
                    KillExcel(process, xlYesterdayFilePath);
                    KillExcel(process, xlMonthFilePath);
                    KillExcel(process, xlMonthLastYearFilePath);
                }
                #endregion

                #region In Tagestabelle schreiben, PDF erzeugen
                
                //Lese InTouch-TagNames aus Vorlage-Datei
                List<List<string>> items = XlReadRowValues(Excel.XlTemplateDayFilePath, Excel.XlDayFileFirstRowToWrite);

                switch (Program.AppStartedBy)
                {
                    case "Schock":
                        Log.Write(Log.Cat.ExcelShock, Log.Prio.Error, 040102, "Start mit Parameter 'Schock' wird nicht mehr unterstützt. XlLog.exe aktualisieren!");
                        //Program.AppErrorOccured = true;
                        break;
                    case "SchockStart":
                        XlWriteShockFreezer(xlDayFilePath, items, true);
                        break;
                    case "SchockEnde":
                        XlWriteShockFreezer(xlDayFilePath, items, false);
                        break;
                    case "AlmDruck":
                        Sql.AlmListToExcel(true);
                        break;
                    case "PdfDruck":
                        Pdf.CreatePdfFromCmd();
                        break;
                    case "Uhrstellen":
                        if (Program.CmdArgs.Length > 2)
                        {
                            SetNewSystemTime.SetNewSystemtimeAndScheduler(Program.CmdArgs[1] + " " + Program.CmdArgs[2]);
                        }
                        break;
                    case "Monatsdatei":
                        if (Program.CmdArgs.Length > 2)
                        {
                            string pathMonthFile = Program.CmdArgs[1];
                            string xlDayFilesDir = Program.CmdArgs[2];
                            Excel.XlNewMonthFileWriteRow(pathMonthFile, xlDayFilesDir);
                        }
                        break;
                    default:

                        //in Excel-Mappe schreiben; ggf. neue Excel-Mappe erstellen
                        XlWriteToDayFile(xlDayFilePath, items);

                        #region PDF erzeugen
                        //Bei jedem Schreiben in Excel-Mappe auch PDF erzeugen.
                        if (XlImmediatelyCreatePdf)
                        {
                            //heutige PDF neu erstellen
                            string pdfTodayFilepath = Path.ChangeExtension(xlDayFilePath, ".pdf");
                            File.Delete(pdfTodayFilepath);
                            Log.Write(Log.Cat.PdfWrite, Log.Prio.Info, 040109, "PDF von heute gelöscht.");

                            Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(0, false, 0, true), false);
                        }
                        //Wenn die Zeit Pdf.PdfConvertStartHour erreicht ist 
                        else if (DateTime.Now.Hour >= Pdf.PdfConvertStartHour)
                        {
                            string pdfYesterdayFilepath = Path.ChangeExtension(xlYesterdayFilePath, ".pdf");

                            //Wenn der PDF - File von gestriger Tagesdatei nicht existiert
                            if (!File.Exists(pdfYesterdayFilepath))
                            {
                                Log.Write(Log.Cat.PdfWrite, Log.Prio.Info, 040103, "Erzeuge PDF " + pdfYesterdayFilepath);
                                Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-1, false, 0, true));
                            }
                        }
                        else
                        {
                            Log.Write(Log.Cat.PdfWrite, Log.Prio.Info, 040104, "PDF-Erzeugung erst ab " + Pdf.PdfConvertStartHour + " Uhr.");
                        }

                        //Am letzten Tag des Monats wird der Vormonat nochmal geprüft.
                        if (DateTime.Now.Day == DateTime.DaysInMonth(DateTime.Now.Year, DateTime.Now.Month))
                        {
                            Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-DateTime.Now.Day - 2, false, 0, true));
                        }

                        #endregion

                        break;
                }
                
                #endregion

                #region Monatstabelle, Monats-PDF schreiben 
                try
                {
                    string pathMonthFile = Excel.CeateXlFilePath(-1, true);                   
                    DateTime monthFileLastModified = File.GetLastAccessTime(pathMonthFile);

                    //Nur einmal täglich in Monatstabelle schreiben! || Edit: am ersten Tag des Monats muss der letzte Tag desvorherigen Monats beachtet werden.
                    bool bed1 = (monthFileLastModified.Day < DateTime.Now.Day); // heute wurde noch nicht geschrieben
                    bool bed2 = (DateTime.Now.Day == 1 && monthFileLastModified.Day > 1); //es ist der erste des Monats und es wurde noch nicht geschrieben
                    bool bed3 = !File.Exists(pathMonthFile); // die Monatsdatei existiert nicht
                    bool bed4 = DateTime.Now.Hour >= Pdf.PdfConvertStartHour; // es ist spät genug
                    bool writeToMonthFilePermission = ( bed1 || bed2 || bed3 ) && bed4 ;

                    Log.Write(Log.Cat.Info, Log.Prio.Info, 040105, string.Format("{0} in Monatsdatei schreiben.", writeToMonthFilePermission ? "Jetzt": "Nicht") );
                    
                    if (writeToMonthFilePermission)
                    {
                        string xlDayFilesDir = CeateXlFilePath(-1, false, 0, true);
                        Excel.XlNewMonthFileWriteRow(pathMonthFile, xlDayFilesDir);

                        //PDF erstellen.
                        Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-1, true, 0, true));

                        //Wenn nicht der erste des Monats ist und keine Monatsdatei vorhanden war, könnten auch Daten im Vormonat fehlen
                        //auskommentiert 21.07.2020 nicht notwendig, verzögert die Programmausführung ...
                        //if (bed3 && DateTime.Now.Day > 1)
                        //{
                        //    //Prüfe auch den Vormonat
                        //    xlDayFilesDir = CeateXlFilePath(-28, false, 0, true);
                        //    Excel.XlNewMonthFileWriteRow(pathMonthFile, xlDayFilesDir);
                        //    Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-28, true, 0, true));
                        //}
                    }                    
                }
                catch (Exception ex)
                {
                    Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040106, string.Format("Fehler beim erstellen der Excel-Monatsdatei : Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                }
                #endregion

            }
            catch (IOException)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040107, string.Format("Die Datei {0} ist bereits geöffnet. Es wird nicht versucht erneut zu schreiben.", xlDayFilePath));
                //Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040108, string.Format("Fehler beim erstellen der Excel-Tagesdatei : Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                InTouch.SetExcelAliveBit(true);
            }
        }

        /// <summary>
        /// Neue Excel-Datei von Vorlage erstellen.
        /// </summary>
        /// <param name="xlTemplateFilePath">Pfad zur Vorlagedatei.</param>
        /// <param name="xlWorkingWorkbookFilePath">Pfad der zu ertsellenden Datei.</param>
        /// <returns>true = Datei erfolgreich erstellt.</returns>
        private static bool XlTryCreateWorkbookFromTemplate(string xlTemplateFilePath, string xlWorkingWorkbookFilePath, int firstRowToWrite) //Fehlernummern siehe Log.cs 0402ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040201, string.Format("XlTryCreateWorkbookFromTemplate({0},{1},{2})", xlTemplateFilePath, xlWorkingWorkbookFilePath, firstRowToWrite));

            try
            {
                if (!File.Exists(xlWorkingWorkbookFilePath))
                {
                    if (!File.Exists(xlTemplateFilePath))
                    {
                        if (Path.GetFileNameWithoutExtension(xlWorkingWorkbookFilePath).Substring(0, 1) == "M")
                        {
                            Log.Write(Log.Cat.FileSystem, Log.Prio.Warning, 040202, string.Format("Keine Monatsdatei-Vorlage gefunden für {0}", xlTemplateFilePath));
                            //kein Fehler
                            return false;
                        }
                        else
                        {
                            Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 040203, string.Format("Es kann keine Excel-Mappe erstellt werden. Die Vorlagedatei für die Tagesdatei wurde nicht gefunden: {0}", xlTemplateFilePath));
                            //Program.AppErrorOccured = true;
                            return false;
                        }
                    }

                    //Kopieren von Vorlage
                    Log.Write(Log.Cat.ExcelWrite, Log.Prio.Info, 040204, string.Format("Erstelle Datei: {0} aus {1}", xlWorkingWorkbookFilePath, xlTemplateFilePath));
                    FileInfo file1 = new FileInfo(xlTemplateFilePath);
                    file1.CopyTo(xlWorkingWorkbookFilePath, false);
                    
                    // Entferne TagNames aus erster zu schreibender Zeile
                    XlDeleteRowValuesExceptBackgroundcolor(xlWorkingWorkbookFilePath, firstRowToWrite);
                    Tools.Wait(1);

                    // Schreibe aktuelles Datum in Datei
                    XlWriteDateToNamedRange(xlWorkingWorkbookFilePath, "Datum");
                   
                    // Bei Monatstabelle: Vorjahreswerte eintragen
                    if (Path.GetFileNameWithoutExtension(xlWorkingWorkbookFilePath).Substring(0, 1) == "M")
                    {
                        XlCopyMonthValuesFromLastYear();
                    }

                    //Setzt den Blattschutz für jedes einzelne Tabellenblatt
                    if (XlPassword.Length > 0)
                    {
                        ProtectSheets(xlWorkingWorkbookFilePath, XlPassword);
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040205, string.Format("Fehler Erstellen der Excel-Datei von Vorlage: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                //Program.AppErrorOccured = true;
                return false;
            }
        }

        /// <summary>   
        /// Sucht die Monatstabelle 1 Jahr vor xlThisMonthFilePath und kopiert von dort die Werte in die Vorjahresspalte von xlThisMonthFilePath.
        /// Setzt die Jahreszahlen in die Spaltenüberschriften von xlThisMonthFilePath.
        /// </summary>
        internal static void XlCopyMonthValuesFromLastYear() //Fehlernummern siehe Log.cs 0403ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040301, string.Format("XlCopyMonthValuesFromLastYear()"));

            int row = Excel.XlMonthFileFirstRowToWrite;
            string xlThisMonthFilePath = CeateXlFilePath(-1, true);
            string xlLastYearMonthFilePath = CeateXlFilePath(-1, true, -1);
            List<Tuple<int, int>> tupleList = new List<Tuple<int, int>>();

            try
            {                
                if (!File.Exists(xlThisMonthFilePath) || !File.Exists(xlLastYearMonthFilePath) || !File.Exists(Excel.XlTemplateMonthFilePath))
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 040302, 
                        string.Format($"Letztjahreswerte konnten nicht in die Datei {xlThisMonthFilePath} eingetragen werden. Es fehlen Quelldateien.\r\n\t\t\t" +
                        $"Letztes Jahr vorhanden: \t{(File.Exists(xlLastYearMonthFilePath) ? "ja" : "nein")}\r\n\t\t\t" +
                        $"Monatsvorlage vorhanden: \t{(File.Exists(Excel.XlTemplateMonthFilePath) ? "ja" : "nein")}\r\n\t\t\t" +
                        $"aktuelle Monatsdatei vorhanden: {(File.Exists(xlThisMonthFilePath) ? "ja" : "nein")}")
                        );
                    //kein Fehler
                    return;
                }

                //Finde Spalten in Template zum kopieren                
                FileInfo file1 = new FileInfo(Excel.XlTemplateMonthFilePath);
                using (ExcelPackage excelPackage1 = new ExcelPackage(file1))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage1.Workbook.Worksheets)
                    {
                        Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040303, "XlCopyMonthValuesFromLastYear() Blatt " + worksheet.Name);

                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                        {
                            //Änderung 02.12.2021: Prüfung auf Hintergrundfarbe herausgenommen, da Fehler bei Badenhop, Ver und scheinbar redundant durch Formelprüfung 

                            //Log.Write(Log.Cat.ExcelRead,Log.Prio.Info, 040307, "Spalte " + col + " = '" + worksheet.Cells[row, col].Value + "'");
                            //Wenn Zelle keinen Farbhintergrund hat (oder weiß oder gelb) und nicht leer ist -> TagName gefunden - (neu 21.01.2020) Wenn Zelle darüber eine Formel hat, die mit '=Datum..' beginnt, in die Tuple aufnehmen.
                            //string backgroundColor = worksheet.Cells[row, col].Style.Fill.BackgroundColor.LookupColor();
                            // bool colorOk = (backgroundColor == null || backgroundColor == yellow || backgroundColor == white || backgroundColor == "#FF000000");
                            
                            bool hasCellValue = worksheet.Cells[row, col].Value != null;
                            bool isDateFormula = worksheet.Cells[row - 1, col].Formula.Contains("Datum"); //Änderung 08.02.2024 .StartsWith() ersetzt durch .Contains(), da der Mappenname in der Formel vor 'Datum' stehen kann.

                            Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040312,
                                string.Format($"Monatsvorlage - Bedingung Spalte {col}, Zeile {row}: " +
                                // $"Farbe {(colorOk ? "ok": "Spalte überspringen; Farbe:" + backgroundColor.ToString())}, " +
                                $"Wert: {(hasCellValue ? "ok" : "ohne Wert")}, " +
                                $"Datum-Formel: {(isDateFormula ? "ok" : "nicht erfüllt")}"));

                            if (hasCellValue && isDateFormula) //colorOk &&
                            {
                                tupleList.Add(new Tuple<int, int>(worksheet.Index, col));
                                Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040308, string.Format("Kopiere Letztjahreswerte in Monatsdatei:\tBlatt: {0}\t Spalte: {1}\tWert: {2}[{3}]", worksheet.Index, col, worksheet.Cells[row, col].Value, worksheet.Cells[row - 1, col].Formula));
                            }
                        }
                    }

                    if (tupleList.Count == 0)
                        Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 040309, string.Format($"In die Monatsdatei {file1} werden keine Letzjahreswerte geschrieben."));
                    else
                        Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040310, string.Format($"In die Monatsdatei {file1} werden {tupleList.Count} Letzjahreswerte geschrieben."));
                }

            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040304, string.Format("Fehler beim Lesen der TagNames aus Excel-Monats-Datei-Vorlage für Übernahme der Vorjahreswerte: {0} \r\n\t\t  Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3} \r\n\t\t StackTrace: {4}", 0, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                //Program.AppErrorOccured = true;
            }

            try
            { 
            //Lese Spalten aus xlLastYearMonthFilePath und schreibe in in xlThisMonthFilePath
            FileInfo file2 = new FileInfo(xlLastYearMonthFilePath);
                FileInfo file3 = new FileInfo(xlThisMonthFilePath);
                using (ExcelPackage excelPackage2 = new ExcelPackage(file2))
                using (ExcelPackage excelPackage3 = new ExcelPackage(file3))
                {
                    foreach (Tuple<int, int> tuple in tupleList)
                    {
                        int worksheetno = tuple.Item1; //.First;
                        int col = tuple.Item2; //.Second;

                        //Werte eintragen
                        ExcelRange source = excelPackage2.Workbook.Worksheets[worksheetno].Cells[row, col, row + 30, col];
                        ExcelRange target = excelPackage3.Workbook.Worksheets[worksheetno].Cells[row, col + 1, row + 30, col + 1];
                        target.Value = source.Value;                        
                    }

                    excelPackage3.Save();
                }
            }
            catch (System.IndexOutOfRangeException ex_index)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.LogAlways, 040311,
                    string.Format("Fehler beim Kopieren der Vorjahreswerte: Die Spaltenbelegung stimmt nicht mit dem Vorjahr überein.\r\n" + ex_index));
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040305, string.Format("Fehler beim Kopieren der Vorjahreswerte in die Excel-Monats-Datei : {0} \r\n\t\t  Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3} \r\n\t\t StackTrace: {4}", 0, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                //Program.AppErrorOccured = true;
            }
        }
    
        /// <summary>
        /// Wartet auf InTouch (D.S. $Minute), bildet ggf. Mittelwerte, schreibt in Tagesdatei, setzt Mittelwerte/Minutenzähler zurück.
        /// </summary>
        /// <param name="xlDayFilePath">Pfad zur Tagesdatei.</param>
        /// <param name="items">Liste der zu lesenden TagNames.</param>
        public static void XlWriteToDayFile(string xlDayFilePath, List<List<string>> items) //Fehlernummern siehe Log.cs 0404ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040401, string.Format("XlWriteToDayFile({0},{1})", xlDayFilePath, "*content*"));

            //Warte auf Berechnung in InTouch (D.S. $Minute) und Bestimme Mittelwerte.
            if (Program.AppStartedBy == "Task" && DateTime.Now.Second < Tools.WaitForScripts)
            {
                Log.Write(Log.Cat.InTouchVar, Log.Prio.Info, 040402, "Warte auf InTouch-Skript D.S. $Minute.");
                Tools.Wait(Tools.WaitForScripts);
            }

            //     Beispiel: true  =                          10:56    ,     4               ,(60 - 5 ) = 55   
            //     Beispiel: false =                          10:53    ,     4               ,(60 - 5 ) = 55    
            bool fullHourTimeframe = !Excel.Between(DateTime.Now.Minute, Excel.XlPosOffsetMin, 60 - Excel.XlNegOffsetMin);

            if (fullHourTimeframe)
            {
                //neu 03.11.2022, Frage den Minutenzähler ab, um die Mittelwertberechnung nachprüfen zu können.
                var minCounter = (int)InTouch.ReadTag("MinutenZähler");

                // Wenn Zeit um Stundensprung, ermittle Mittelwerte in InTouch; Setzte ExBestStdWerte
                // geändert 12.06.2023 Minutenzähler protokollieren
                Log.Write(Log.Cat.InTouchVar, Log.Prio.Info, 040403, $"Setze InTouch-Variable >{Program.InTouchDiscSetCalculations}< = {true} bei Minutenzähler = {minCounter}");
                InTouch.WriteDiscTag(Program.InTouchDiscSetCalculations, true);

                //neu 4.5.2021: Warte nochmal, weil Skript für Mittelwerte bis zu 5 sek benötigt
                Log.Write(Log.Cat.InTouchVar, Log.Prio.Info, 040406, "Warte auf InTouch-Skripte.");
                Tools.Wait(Tools.WaitForScripts);
            }

            Tools.Wait((int)Tools.WaitForScripts/4);
            
            XlDayFileWriteRange(xlDayFilePath, items);

            Tools.Wait((int)Tools.WaitForScripts/4);

            if (fullHourTimeframe)
            {
                //Edit: Im D. S. $Hour wird gesetzt   ExLöscheStdMin = 1;
                Log.Write(Log.Cat.InTouchVar, Log.Prio.Info, 040404, string.Format("Setze InTouch-Variable >{0}< = {1}.", Program.InTouchDiscResetHourCounter, true));
                InTouch.WriteDiscTag(Program.InTouchDiscResetHourCounter, true);
            }

            //Viertelstundenwerte auch bei Stundensprung zurücksetzen!
            Log.Write(Log.Cat.InTouchVar, Log.Prio.Info, 040405, string.Format("Setze InTouch-Variable >{0}< = {1}.", Program.InTouchDiscResetQuarterHourCounter, true));
            InTouch.WriteDiscTag(Program.InTouchDiscResetQuarterHourCounter, true);
            
        }

        /// <summary>
        /// Schreibt eine Zeile in die Excel-Datei xlFilePath, wenn das Tabellenblatt mindestens eine Spalte "Uhrzeit" enthält".
        /// </summary>
        /// <param name="xlFilePath"></param>
        /// <param name="content">2-Dimensionale Liste der Tabellenblätter und Variablen</param>
        public static void XlDayFileWriteRange(string xlFilePath, List<List<string>> content) //Fehlernummern siehe Log.cs 0405ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040501, string.Format("XlDayFileWriteRange({0},{1})", xlFilePath, "*content*"));

            if (!XlTryCreateWorkbookFromTemplate(Excel.XlTemplateDayFilePath, xlFilePath, Excel.XlDayFileFirstRowToWrite)) return;

            if (content == null) return;

            try
            {
                //create a fileinfo object of an excel file on the disk
                FileInfo file = new FileInfo(xlFilePath);

                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    int worksheetNo = 0;

                    foreach (List<string> wsContent in content)
                    {
                        ++worksheetNo;
                        if (excelPackage.Workbook?.Worksheets?.Count < worksheetNo) break;
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheetNo];

                        //Liste Uhrzeiten-Spalten von Worksheet
                        int[] timeCols = XlTimeColCount(worksheet);

                        if (timeCols.Length == 0) // Wenn keine Zeiten-Spalten auf dem Blatt vorhanden sind vermutlich ein Schockkühler-Blatt
                        {
                            Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040502, string.Format("Blatt {0} '{1}' aus {2} hat keine Zeitspalte.", worksheetNo, excelPackage.Workbook.Worksheets[worksheetNo].Name, Path.GetFileNameWithoutExtension(xlFilePath) ) );
                        }
                        else
                        {
                            #region Zeile in Tabellenblatt Kühlraum oder TK-Raum schreiben
                            int row = XlSetRowAndCol(timeCols, out int col);
                            if (col < 1 || row < 1)
                            {
                                // Es muss in dieses Blatt nichts geschrieben werden.
                                Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040503, string.Format("Schreibe keine Einträge in Blatt {0} '{1}' ({2} Zeitspalten)", worksheetNo, excelPackage.Workbook.Worksheets[worksheetNo].Name, timeCols.Length));
                                continue;
                            }

                            foreach (string TagName in wsContent) //Zeile füllen
                            {
                                if (TagName.Length < 1 || TagName == null)
                                {
                                    //Wenn der Spalte keine Variable zugeordnet wurde
                                    worksheet.Cells[row, col].Value = null;
                                }
                                else
                                {
                                    if (TagName != "DoNotChangeCell") //"DoNotChangeCell" -> Inhalt belassen z.B. statische Texte, Formeln usw.
                                    {
                                        var result = InTouch.ReadTag(TagName);
                                        // Wenn Variable nicht vorhanden ist, wird von Intouch float.MaxValue ausgegeben.
                                        if (Convert.ToSingle(result) != float.MaxValue)
                                        {
                                            worksheet.Cells[row, col].Value = result;

                                            int.TryParse(result.ToString(), out int resultInt);

                                            if (TagName.StartsWith("Min") && resultInt > 60)
                                            {
                                                Log.Write(Log.Cat.InTouchVar, Log.Prio.Warning, 040504, string.Format(
                                                    "Der Wert {0} = {1} scheint eine Betriebszeit darzustellen, ist aber größer als 60.", TagName, result));
                                            }
                                        }
                                        else
                                        {
                                            worksheet.Cells[row, col].Formula = "NA()";
                                            worksheet.Cells[row, col].Calculate();
                                        }
                                    }
                                }
                                ++col;
                            }
                            #endregion
                        }
                       
                        // Druckeinstellung "Blatt auf einer Seite darstellen" 
                       worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
                    }

                    //calculate all the values of the formulas in the Excel file                    
                    excelPackage.Workbook.Calculate();
                    excelPackage.Workbook.CalcMode = ExcelCalcMode.Automatic;
                    excelPackage.Workbook.Properties.LastModifiedBy = Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetAssembly(typeof(Program)).Location);
                    //Set Focus to First Worksheet.
                    excelPackage.Workbook.Worksheets[1].Cells[1, 1].Value = null;
                    //save the changes
                    excelPackage.Save();

                    Log.Write(Log.Cat.Info, Log.Prio.Info, 040505, string.Format("In {0} wurden in {1} Tabellenblättern {2} Werte bearbeitet.", Path.GetFileName(xlFilePath), content.Count, content.SelectMany(list => list).Distinct().Count()));
                }
            }
            catch(InvalidOperationException)
            {
                Log.Write(Log.Cat.ExcelWrite,  Log.Prio.Error,040506, string.Format("Die Excel-Datei {0} konnte nicht beschrieben werden. Sie ist vermutlich durch ein anderes Programm geöffnet.", Path.GetFileNameWithoutExtension(xlFilePath) ) );
                //Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040507, string.Format("Fehler beim schreiben in die Excel-Datei: {0} \r\n\t\t\t\t Typ: {1} \r\n\t\t\t\t Fehlertext: {2}  \r\n\t\t\t\t InnerException: {3} \r\n\t\t\t\t StackTrace: {4}", xlFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Schreibe in Monatstabelle die Tagesummen von allen Tagen des Monats.
        /// </summary>
        public static void XlNewMonthFileWriteRow(string xlMonthFilePath, string xlDayFilesDir) //Fehlernummern siehe Log.cs 0406ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040601, string.Format("XlNewMonthFileWriteRow({0}, {1})", xlMonthFilePath, xlDayFilesDir));

            //Datei für Monatstabelle erzeugen            
            if (!XlTryCreateWorkbookFromTemplate(Excel.XlTemplateMonthFilePath, xlMonthFilePath, Excel.XlMonthFileFirstRowToWrite)) return;
            Tools.Wait(1);
            
            try
            {
                #region //Lese InTouch-TagNames aus Vorlage-Datei

                List<List<string>> itemsDay = XlReadRowValues(Excel.XlTemplateDayFilePath, Excel.XlDayFileFirstRowToWrite);
                string[] itemsDayArray = itemsDay.SelectMany(x => x).ToArray();
                List<List<string>> itemsMonth = XlReadRowValues(Excel.XlTemplateMonthFilePath, Excel.XlMonthFileFirstRowToWrite);

                //Zeilennummer für "Summe" finden
                int sumRowNo = XlGetRowInColByValue(Excel.XlTemplateDayFilePath, 1, 1, "Summe");
                //Log.Write(Log.Category.ExcelRead, 2001201442, Path.GetFileNameWithoutExtension(Excel.XlTemplateDayFilePath) + ": Summenzeile " + sumRowNo);
        
                //Liste alle *.xlsx-Dateien im Monatsordner, deren Tages-Datum vor heute liegt, aufsteigend nach Erstelldatum sortiert.
                DirectoryInfo directory = new DirectoryInfo(xlDayFilesDir);
                IOrderedEnumerable<FileInfo> xlDayFileList = directory.GetFiles().Where(x => x.Extension.Equals(".xlsx")).OrderBy(f => f.CreationTime);
                Dictionary<int, List<List<string>>> allDayValues = new Dictionary<int, List<List<string>>>();

                try
                {
                    foreach (FileInfo xlDayFileInfo in xlDayFileList)
                    {
                        //Überspringe temporäre Dateien 
                        // if (xlDayFileInfo.Name.StartsWith("~")) continue;

                        // neu 06.10.2020: Überspringe wenn Name nicht mit "T_" anfängt oder länger als ist als das Muster
                        if (!xlDayFileInfo.Name.StartsWith("T_") || xlDayFileInfo.Name.Length > 15) continue;

                        //Nehme den Tag aus dem Dateinamen, da Erstelldatum nicht eindeutig ist (T_01... wird erst am 2. Tag um 0:15 Uhr erzeugt.)!
                        if (int.TryParse(xlDayFileInfo.Name.Substring(2, 2), out int dayNo)) // Dateiname-Muster: M_ddmmyyyy.xlsx
                        {
                            //Log.Write(Log.Category.FileSystem, 2001201441, "Lese Tabelle von Tag Nr. " + dayNo);
                            if (dayNo < DateTime.Now.Day || DateTime.Now.Day == 1) // am 1. des Monats alle Dateien des Vormonats zusammenfassen.
                            {
                                List<List<string>> valuesDay = XlReadRowValues(xlDayFileInfo.FullName, sumRowNo, true);
                                allDayValues.Add(dayNo, valuesDay);
                            }
                        }
                    }
                }
                catch
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 040609, "foreach (FileInfo xlDayFileInfo in xlDayFileList)");
                }
                #endregion

                FileInfo file = new FileInfo(xlMonthFilePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    foreach (var dayValues in allDayValues) // Tagesdateien in Monatsordner
                    {
                        int worksheetNo = 0;
                        int row = dayValues.Key + Excel.XlMonthFileFirstRowToWrite - 1;
                        Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 040602, string.Format("Schreibe Zeile {0:00} in Monatsdatei {1}", row,  Path.GetFileNameWithoutExtension(xlMonthFilePath) ));

                        string[] valuesDayArray = dayValues.Value.SelectMany(x => x).ToArray();
                        
                        foreach (List<string> WorksheetsItemsM in itemsMonth) // Arbeitsblätter in Monatstabelle
                        {
                            worksheetNo++;
                            int col = 0;

                            try
                            {
                                foreach (string item in WorksheetsItemsM) // Tagnames in Arbeitsblatt
                                {
                                    col++;
                                    if (item.Length < 3) continue; //leerer TagName                                    
                                    int pos = Array.IndexOf(itemsDayArray, item); //Position Tagname in Vorlagedatei; -1 = nicht gefunden.

                                    if (pos < 0) //TagName nicht gefunden 
                                    {
                                        Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 040603, string.Format("Der TagName '{0}' konnte nicht im Blatt {1}, Zeile {2}, Spalte {3} in der Tagesvorlage gefunden werden und wird übersprungen.", item, worksheetNo, row, col));
                                        continue;
                                    }

                                    string writeValue = valuesDayArray[pos]; // Wert in Tagesdatei
                                    string itemValue = itemsDayArray[pos];
                                    if (itemValue.Length < 1 || itemValue == "DoNotChangeCell") continue; // Wenn Zelle leer, oder nicht zu verändern. 
                   
                                    float.TryParse(writeValue, out float result);
                                    excelPackage.Workbook.Worksheets[worksheetNo].Cells[row, col].Value = result; // "R" + row + " C" + col; 
                                }
                            }
                            catch (Exception exTest)
                            {
                                Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 040604, "Fehler in foreach (string item in WorksheetsItemsM)\r\nSpalte " + col + "\r\n\t" + exTest.Message + " \r\n\t" + exTest.StackTrace );
                                //Program.AppErrorOccured = true;
                            }
                            excelPackage.Workbook.Worksheets[worksheetNo].PrinterSettings.PaperSize = ePaperSize.A4;
                        }                  
                    }

                    //calculate all the values of the formulas in the Excel file                    
                    excelPackage.Workbook.Calculate();
                    excelPackage.Workbook.CalcMode = ExcelCalcMode.Automatic;
                    excelPackage.Workbook.Properties.LastModifiedBy = Path.GetFileNameWithoutExtension(System.Reflection.Assembly.GetAssembly(typeof(Program)).Location);
                    //Set Focus to First Worksheet.
                    excelPackage.Workbook.Worksheets[1].Cells[1, 1].Value = null;
                    //save the changes
                    excelPackage.Save();
                }

                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Info, 040605, "Bearbeitung Monatsdatei abgeschlossen.");

                //PDF erstellen.
                Pdf.CreatePdf(xlMonthFilePath);
            }
            catch (System.ArgumentException exArg)
            {
                Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 040606, "Die Namen der Tagesdateien im Monatsordner sind nicht eindeutig für\r\n\t\t\t\t" +xlMonthFilePath + "\r\n\t\t\t\tDie Monatsdatei kann nicht erstellt werden.\r\n\t\t\t\t" + exArg.Message);
                //Program.AppErrorOccured = true;
            }
            catch (IndexOutOfRangeException index_ex)
            {
                Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 040607, "Arrayfehler in Vorlagedatei. Mögliche Ursache: Farbhintergründe im Wertefeld in T_vorl.xlsx oder M_vorl.xlsx\r\nFehlermeldung: " + index_ex.Message + "\r\n\t" + index_ex.InnerException + "\r\n\t" + index_ex.StackTrace );
                //Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 040608, string.Format("Fehler beim erstellen der Excel-Monats-Datei : xlMonthFilePath {0} \r\n\t\t\t\tTyp: {1} \r\n\t\t\t\tFehlertext: {2}  \r\n\t\t\t\tInnerException: {3} \r\n\t\t\t\tStackTrace: {4}", xlMonthFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                //Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Schreibt eine Zeile in die Excel-Datei xlFilePath, wenn das Tabellenblatt keine Spalte "Uhrzeit" enthält.
        /// </summary>
        /// <param name="xlFilePath">Pfad zur Excel-Mappe</param>
        /// <param name="content">Liste der Tabellenblätter mit Liste der zu schreibenden Variablen</param>
        public static void XlWriteShockFreezer(string xlFilePath, List<List<string>> content, bool IsSchockStart) //Fehlernummern siehe Log.cs 0407ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040701, string.Format("XlWriteShockFreezer({0}, *content*, {1})", xlFilePath, IsSchockStart? "Start":"Ende"));

            if (!XlTryCreateWorkbookFromTemplate(Excel.XlTemplateDayFilePath, xlFilePath, Excel.XlDayFileFirstRowToWrite)) return;

            FileInfo file = new FileInfo(xlFilePath);

            using (ExcelPackage excelPackage = new ExcelPackage(file))
            {
                int worksheetNo = 0;
                try
                {
                    // gehe durch alle Tabellenblätter
                    foreach (List<string> wsContent in content)
                    {
                        #region Finde Schockkühler-Blatt in Workbook
                        ++worksheetNo;
                        if (excelPackage.Workbook?.Worksheets?.Count < worksheetNo) break;
                        ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheetNo];

                        //Liste Uhrzeiten-Spalten von Worksheet
                        int[] timeCols = XlTimeColCount(worksheet);

                        if (timeCols.Length > 0) continue; // Wenn Zeiten-Spalten auf dem Blatt vorhanden sind überspringen (-> kein Schockkühler)
                        #endregion

                        #region Lese Spalte Schockkühler-Nr. und finde die erste leere Zelle darin

                        int rangeMaxRows = worksheet.Dimension.End.Row;
                        int firstEmptyRow;

                        //Finde erste Zeile, in der keine Schockkühler-Nr. steht
                        for (firstEmptyRow = Excel.XlDayFileFirstRowToWrite; firstEmptyRow <= rangeMaxRows; firstEmptyRow++)
                        {
                            object cellValue = worksheet.Cells[firstEmptyRow, (int)XlShockCol.ShockNo].Value;
                            if (cellValue == null) break;
                        }
                        Log.Write(Log.Cat.ExcelShock, Log.Prio.Info, 040702, string.Format("In Blatt {0} gibt es {1} Zeilen. Erste Leerzeile: {2}", worksheetNo, rangeMaxRows, firstEmptyRow));

                        //Liste alle Werte der Spalte Schockkühler-Nr. auf.
                        ExcelRange shockNoCol = worksheet.Cells[Excel.XlDayFileFirstRowToWrite, (int)XlShockCol.ShockNo, firstEmptyRow, (int)XlShockCol.ShockNo];
                        List<int> shockNoList = new List<int>();
                        
                        foreach (var cell in shockNoCol)
                        {
                            shockNoList.Add(cell.GetValue<int>());
                        }
                        #endregion

                        #region Lese InTouch-Werte

                      //  NativeMethods intouch = new NativeMethods(0, 0); //Ungetestet: Wenn keine Fehler beim lesen der TagNames auftreten: löschen

                        int SchockNr = (int)InTouch.ReadTag("SchockNr"); // intouch.ReadInteger("SchockNr");
                        if (SchockNr < 1)
                        {
                            Log.Write(Log.Cat.InTouchVar, Log.Prio.Warning, 040703, "InTouch-Tag \"SchockNr\": Es wurde kein Schockkkühler definiert.");
                            break; // Es wurde kein Schockkühler gesetzt.
                        }
                        #endregion

                        #region Starte/beende Schockkühlvorgang; schreibe in Excel-Tabelle 

                        #region Finde passende Zeile                        
                        int currentRowNo = firstEmptyRow;

                        // Wenn kein neuer Schockkühlvorgang und Werte in Spalte SchockNr vorhanden: Finde Zeile mit letztem Auftreten von SchockNr
                        if (!IsSchockStart && shockNoList.Count > 0)
                        {
                            currentRowNo = shockNoList.LastIndexOf(SchockNr) + XlDayFileFirstRowToWrite;
                            if (currentRowNo < XlDayFileFirstRowToWrite) currentRowNo = firstEmptyRow; // Dieser Schockkühler wurde noch nicht aufgezeichnet: gehe in erste leere Zeile.
                        }
                 
                        bool startTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Value == null;
                        bool endTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.EndTime].Value == null;

                        // Wenn Start- und End-Zeit schon gefüllt sind, dann gehe in erste leere Zeile
                        if (!startTimeIsEmpty && !endTimeIsEmpty)
                        {
                            currentRowNo = firstEmptyRow;
                            //startTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Value == null;
                            //endTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.EndTime].Value == null;    
                        }

                        #endregion

                        //LeseTagNames für Schockkühler-Temperaturen
                        string[] ProductTempTagNames = wsContent[(int)XlShockCol.ProductStartTemp - 1].Split(';');                       
                        if (ProductTempTagNames.Length < SchockNr)
                        {
                            Log.Write(Log.Cat.ExcelShock, Log.Prio.Error, 040704, string.Format("Excelvorlage Blatt Schockkühler prüfen: Für Schockkühler {0} ist kein TagName Produkttemperatur eingetragen; Eintrag: {0}", SchockNr, ProductTempTagNames.ToArray()));            
                        }

                        string[] RoomTempTagNames = wsContent[(int)XlShockCol.RoomStartTemp - 1].Split(';');
                        if (RoomTempTagNames.Length < SchockNr)
                        {
                            Log.Write(Log.Cat.ExcelShock, Log.Prio.Error, 040705, string.Format("Excelvorlage Blatt Schockkühler prüfen: Für Schockkühler {0} ist kein TagName Raumtemperatur eingetragen; Eintrag: {0}", SchockNr, RoomTempTagNames.ToArray()));
                        }

                        // neuen Schockkühlvorgang starten
                        if (IsSchockStart)
                        {
                            string ChargenNrS = (string)InTouch.ReadTag("ChargenNrS"); //intouch.ReadString("ChargenNrS");
                            int SchockRProgX = (int)InTouch.ReadTag("SchockRProg" + SchockNr); //intouch.ReadInteger("SchockRProg" + SchockNr);
                            int SchockSollDauer = (int)InTouch.ReadTag("SchockSollDauer");// intouch.ReadInteger("SchockSollDauer");
                            float ProductStartTemp = (float)InTouch.ReadTag(ProductTempTagNames[SchockNr - 1]); // intouch.ReadFloat(ProductTempTagNames[SchockNr - 1]);
                            float RoomStartTemp = (float)InTouch.ReadTag(RoomTempTagNames[SchockNr - 1]); //intouch.ReadFloat(RoomTempTagNames[SchockNr - 1]);

                            Log.Write(Log.Cat.ExcelShock, Log.Prio.LogAlways, 040706, string.Format(
                                "Neuer Vorgang wird in Zeile {0} geschrieben: Schockkühler {1}, Charge {2}, Programm {3}, max. Soll-Dauer {4} min, Temp. R{5:F1}°C/P{6:F1}°C", 
                                currentRowNo, SchockNr, ChargenNrS, SchockRProgX, SchockSollDauer, RoomStartTemp, ProductStartTemp)
                                );

                            worksheet.Cells[currentRowNo, (int)XlShockCol.ChargeNo].Value = ChargenNrS;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ShockNo].Value = SchockNr;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Value = DateTime.Now.ToLongTimeString();
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ShockProgramNo].Value = ( SchockRProgX == int.MaxValue ) ? (object)"TagName?" : SchockRProgX ;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.MaxDuration].Value = (SchockSollDauer == int.MaxValue) ? (object)"TagName?" : SchockSollDauer;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ProductStartTemp].Value = (ProductStartTemp == float.MaxValue) ? (object)"TagName?" : ProductStartTemp;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.RoomStartTemp].Value = (RoomStartTemp == float.MaxValue) ? (object)"TagName?" : RoomStartTemp;
                        }
                        else 
                        {
                            float ProductEndTemp = (float)InTouch.ReadTag(ProductTempTagNames[SchockNr - 1]); //intouch.ReadFloat(ProductTempTagNames[SchockNr - 1]);
                            float RoomEndTemp = (float)InTouch.ReadTag(RoomTempTagNames[SchockNr - 1]); //intouch.ReadFloat(RoomTempTagNames[SchockNr - 1]);
                            float SchockMinX = (float)InTouch.ReadTag("SchockMin" + SchockNr); //intouch.ReadFloat("SchockMin" + SchockNr);
                            float SchockMaxX = (float)InTouch.ReadTag("SchockMax" + SchockNr); //intouch.ReadFloat("SchockMax" + SchockNr);
                            float SchockMinXK = (float)InTouch.ReadTag("SchockMin" + SchockNr + "K"); //intouch.ReadFloat("SchockMin" + SchockNr + "K");
                            float SchockMaxXK = (float)InTouch.ReadTag("SchockMax" + SchockNr + "K"); //intouch.ReadFloat("SchockMax" + SchockNr + "K");

                            Log.Write(Log.Cat.ExcelShock, Log.Prio.LogAlways, 040707, string.Format(
                                "Vorgang Schockkühler {0} in Zeile {1} beendet. Ende R{2:F1}°C/P{3:F1}°C, Min R{4:F1}°C/P{5:F1}°C, Max R{6:F1}°C/P{7:F1}°C", 
                                SchockNr, currentRowNo, RoomEndTemp, ProductEndTemp, SchockMinX, SchockMinXK, SchockMaxX, SchockMaxXK)
                                );

                            worksheet.Cells[currentRowNo, (int)XlShockCol.EndTime].Value = DateTime.Now.ToLongTimeString();
                            string strStartTime = worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Text;
                            if (!TimeSpan.TryParse(strStartTime, out TimeSpan startTime)) //geändert 07.03.2023
                            {
                                Log.Write(Log.Cat.ExcelShock, Log.Prio.Warning, 040709, $"Der Beginn des Schockkühlvorgangs '{strStartTime}' in Zeile {currentRowNo} konnte nicht als Zeit gelesen werden.");
                                startTime = DateTime.Now.TimeOfDay;
                            }
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ShockDuration].Value = DateTime.Now.Add(-startTime).ToShortTimeString();
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ProductEndTemp].Value = (ProductEndTemp == float.MaxValue) ? (object)"TagName?" : ProductEndTemp;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ProductMinTemp].Value = (SchockMinXK == float.MaxValue) ? (object)"TagName?" : SchockMinXK;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ProductMaxTemp].Value = (SchockMaxXK == float.MaxValue) ? (object)"TagName?" : SchockMaxXK;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.RoomEndTemp].Value = (RoomEndTemp == float.MaxValue) ? (object)"TagName?" : RoomEndTemp;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.RoomMinTemp].Value = (SchockMinX == float.MaxValue) ? (object)"TagName?" : SchockMinX;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.RoomMaxTemp].Value = (SchockMaxX == float.MaxValue) ? (object)"TagName?" : SchockMaxX;
                        }

                        //calculate all the values of the formulas in the Excel file
                        excelPackage.Workbook.Calculate();
                        excelPackage.Workbook.CalcMode = ExcelCalcMode.Automatic;
                        excelPackage.Workbook.Properties.LastModifiedBy = "Schockkühler " + SchockNr;
                        excelPackage.Save();

                        #endregion
                    }
                }
                catch (Exception ex)
                {
                    Log.Write(Log.Cat.ExcelShock, Log.Prio.Error, 040708, string.Format("Fehler beim Schreiben in die Excel-Datei (Schockkühler): \r\n\t\t Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3} \r\n\t\t GetBaseException().Message: {4}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace, ex.GetBaseException().Message));
                    Program.AppErrorOccured = true;
                }
            }

        }

        #endregion

        #region Tools for WriteToExcelFiles

        /// <summary>
        /// Gibt an, ob ein Wert num zwischen zwei Werten lower und upper liegt. 
        /// </summary>
        /// <param name="num">Zu ntersuchende Zahl</param>
        /// <param name="lower">UNtere Zahl</param>
        /// <param name="upper">Obere Zahl</param>
        /// <param name="inclusive">Sind lower und upper innerhalb des Suchbereichs?</param>
        /// <returns>true = der Wert liegt zwischen lower und upper.</returns>
        public static bool Between(this int num, int lower, int upper, bool inclusive = true) //Fehlernummern siehe Log.cs 0408ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040801, string.Format("Between({0},{1},{2},{3})",num, lower, upper,inclusive));

            return inclusive
                ? lower <= num && num <= upper
                : lower < num && num < upper;
        }
        
        /// <summary>
        /// Findet in der angegebenen Spalte col das erste Auftreten des Wertes searchValue und gibt die Zeilennummer aus.
        /// </summary>
        /// <param name="xlFilePath">Pfad zur Excel-Datei</param>
        /// <param name="worksheetno">Arbeitsblatt-Nr.</param>
        /// <param name="col">Spalten-Nr.</param>
        /// <param name="searchValue">Inhalt der Zelle, deren Zeilennumer ausgegeben werdne soll</param>
        /// <returns>Zeilennummer des gesuchten Werts; bei Fehler -1</returns>
        public static int XlGetRowInColByValue(string xlFilePath, int worksheetno, int col, string searchValue) //Fehlernummern siehe Log.cs 0409ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 040901, string.Format("XlGetRowInColByValue({0},{1},{2},{3},)", xlFilePath, worksheetno, col, searchValue));

            try
            {
                if (!File.Exists(xlFilePath))
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 040902, string.Format("Datei nicht gefunden: {0}", xlFilePath));
                    //Program.AppErrorOccured = true;
                    return -1;
                }

                //read the Excel file as byte array
                byte[] bin = File.ReadAllBytes(xlFilePath);

                //create a new Excel package in a memorystream
                //using (MemoryStream stream = new MemoryStream(bin))
                MemoryStream stream = new MemoryStream(bin);

                using (ExcelPackage excelPackage = new ExcelPackage(stream))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheetno];

                    if (col > worksheet.Dimension.Columns) return -1;
                   
                    for (int row = worksheet.Dimension.End.Row ; row >= Excel.XlMonthFileFirstRowToWrite; row--)
                    {
                        if (worksheet.Cells[row, col].Value.ToString() == searchValue)
                        {
                            return row;
                        }
                    }
                }

                Log.Write(Log.Cat.ExcelRead, Log.Prio.Warning, 040903, string.Format("Wert >{0}< wurde nicht gefunden in Datei {1}, Blatt {2}, Spalte {3}.", searchValue, xlFilePath, worksheetno, col));
                return -1;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 040904, string.Format("Fehler beim finden der Zeile >{0}< in Excel-Datei: {1}\r\n\t\t\t Typ: {2}\r\n\t\t\t Fehlertext: {3}\r\n\t\t\t InnerException: {4}\r\n\t\t\t StackTrace: {5}\r\n\t\t\t Source: {6}\r\n\t\t\t GetBaseException: {7}", searchValue, xlFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace, ex.Source, ex.GetBaseException().Message ));
                return -1;
            }
        }

        /// <summary>
        /// Liest eine komplette Reihe in einem Excel-Sheet
        /// </summary>
        /// <param name="xlFilePath">Pfad zum Excel-Sheet</param>
        /// <param name="worksheetNo">Lfd. Nr. des Tabellenblatts</param>
        /// <param name="row">Nummer der zu lesenden Zeile </param>
        /// <returns>Liste der Zellenwerte in dieser Reihe. Listen-Eintrag "DoNotChangeCell" markiert, dass Zellwerte beim Schreiben der Liste nicht geändert werden sollen.
        /// neu 10.03.2021: Wenn Listen-Eintrag eine Formel enthält wie "DoNotChangeCell".
        /// </returns>
        public static List<List<string>> XlReadRowValues(string xlFilePath, int row, bool recordBGColoredCells = false) //Fehlernummern siehe Log.cs 0410ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041001, string.Format("XlReadRowValues({0},{1},{2})", xlFilePath, row, recordBGColoredCells));

            try
            {
                if (!File.Exists(xlFilePath))
                {
                    Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 041002, string.Format("Datei nicht gefunden: {0}", xlFilePath));
                    //Program.AppErrorOccured = true;
                    return null;
                }

                //create a list to hold all the values
                List<List<string>> excelData = new List<List<string>>(); //Creates new nested List

                //read the Excel file as byte array
                byte[] bin = File.ReadAllBytes(xlFilePath);

                using (ExcelPackage excelPackage = new ExcelPackage(new MemoryStream(bin)))
                {
                    int wsCount = -1;
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {                        
                        if (row > worksheet.Dimension.Rows) continue; //Wenn die übergebene Zeile nicht auf dem Blatt vorhanden ist

                        excelData.Add(new List<String>()); //Adds new sub List
                        ++wsCount;

                        //loop all columns in a row
                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++) //vorher nur kleiner als -> letzte Spalte wurde nicht gefüllt
                        {
                            var val = worksheet.Cells[row, col].Value;
                            
                            //Console.Write("R{0},C{1} | ", row, col);

                            //add the cell data to the List 
                            if (val == null)
                            {
                                excelData[wsCount].Add("");
                            } 
                            else if (val.ToString().StartsWith("=") ) //neu 10.03.2021; #UNGETESTET# Wenn die Zelle eine Formel enthält
                            {
                                excelData[wsCount].Add("DoNotChangeCell");
                            }
                            else
                            {
                                //Wenn keine Hintergrundfarbe oder gelb gesetzt ist
                                string backgroundColor = worksheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb;
                                if (backgroundColor == null || backgroundColor == yellow || backgroundColor == white || recordBGColoredCells)
                                {
                                    excelData[wsCount].Add(val.ToString().Trim());
                                }
                                else
                                {
                                    excelData[wsCount].Add("DoNotChangeCell");
                                }
                            }

                        }
                    }
                }

                return excelData;
            }
            catch(IOException)
            {
                Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 041003, string.Format("Die Excel-Datei {0} konnte nicht beschrieben werden. Sie ist vermutlich durch ein anderes Programm geöffnet.", Path.GetFileNameWithoutExtension( xlFilePath ) ) );
                Program.AppErrorOccured = true;
                return null;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelRead, Log.Prio.Error, 041004, string.Format("Fehler beim Lesen aus Excel-Datei: {3} \r\n\t\t\t\t Typ: {0} \r\n\t\t\t\t Fehlertext: {1}  \r\n\t\t\t\t InnerException: {2}  \r\n\t\t\t\t StackTrace: {4}", ex.GetType().ToString(), ex.Message, ex.InnerException, xlFilePath, ex.StackTrace));
                Program.AppErrorOccured = true;
                return null;
            }

        }
        
        /// <summary>
        /// XLSX-Filename + Ordner für Jahr/Monat + Pfad erzeugen.
        /// </summary>
        /// <param name="OffsetDays">Tage von heute (negative Werte liegen in der Vergangenheit).</param>
        /// <returns>Pfad zur Excel-Tages- oder Monats-Datei oder Ordner.</returns>
        public static string CeateXlFilePath(int OffsetDays = 0, bool MonthFile = false, int OffsetYears = 0, bool directoryOnly = false) //Fehlernummern siehe Log.cs 0411ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041101, string.Format("CeateXlFilePath({0},{1},{2},{3})", OffsetDays, MonthFile, OffsetYears , directoryOnly ));

            DateTime date = DateTime.Now.AddYears(OffsetYears).AddDays(OffsetDays).AddMinutes(-XlNegOffsetMin); // z.B,um 00:05 Uhr soll der Vortag beschrieben werden. 

            // Ordner Jahr
            string xlFilePath = Path.Combine(XlArchiveDir, "Tabellen");
            xlFilePath = Path.Combine(xlFilePath, date.Year.ToString());
            //Directory.CreateDirectory(xlFilePath);

            string[] MonthNames = { "unbekannt_", "Jan", "Feb", "Mrz", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez" };
            string filename;

            if (MonthFile)
            {
                filename = string.Format("M_{0}{1}.xlsx", MonthNames[date.Month], date.Year.ToString());
            }
            else
            {
                //Ordner Monat
                // Format MMM z.B. /Okt/
                xlFilePath = Path.Combine(xlFilePath, MonthNames[date.Month] );
                // Format MMM_YYYY z.B. /Okt_2019/
                //xlFilePath = Path.Combine(xlFilePath, MonthNames[date.Month] + "_" + date.Year.ToString());
                Directory.CreateDirectory(xlFilePath);
                filename = string.Format("T_{0}.xlsx", date.ToString("ddMMyyyy"));
            }
            //Datei Pfad
            string returnPath;
            if (directoryOnly) returnPath = xlFilePath;
            else returnPath = Path.Combine(xlFilePath, filename);
            return returnPath;
        }
        
        /// <summary>
        /// Löscht die Inhalte der Zeile mit der Nummer XlFirstRowToWrite in allen Tabellenblättern von xlWorkingWorkbookFilePath, ausgenommen Zellen mit Farbhintergrund (Uhrzeit).
        /// </summary>
        /// <param name="xlWorkingWorkbookFilePath">Excel-Datei</param>
        private static void XlDeleteRowValuesExceptBackgroundcolor(string xlWorkingWorkbookFilePath, int row) //Fehlernummern siehe Log.cs 0412ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041201, string.Format("XlDeleteRowValuesExceptBackgroundcolor({0},{1})", xlWorkingWorkbookFilePath, row));

            // aus Vorlage übernommene TagNames löschen
            if (File.Exists(xlWorkingWorkbookFilePath))
            {
                //create a fileinfo object of an excel file on the disk
                FileInfo file = new FileInfo(xlWorkingWorkbookFilePath);

                //create a new Excel package from the file
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                    {
                        //loop all columns in a row
                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                        {
                            //Wenn keine Hintergrundfarbe oder weiß oder gelb gesetzt ist
                            string backgroundColor = worksheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb; // .LookupColor(); funktioniert nicht!
                            bool isFormula = worksheet.Cells[row, col].Formula.Length > 0;

                            if (backgroundColor == null || backgroundColor == yellow || backgroundColor == white)
                            {
                                //Setze Schriftfarbe auf schwarz und Hintergrund transparent.
                                worksheet.Cells[row, col].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                worksheet.Cells[row, col].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.None;

                                if (isFormula)
                                    continue;
                                if (col == 1)
                                {
                                    //Erste Spalte in erster (Schreib-)Zeile hat keinen Farbhintergrund => Schockkühler-Blatt
                                    worksheet.Cells[row, col].Value = "keine Schockkühlvorgänge registriert";
                                }
                                else
                                {
                                    // .Clear() löscht die Formatierung, daher null setzen
                                    worksheet.Cells[row, col].Value = null;
                                }
                            }
                            else if (isFormula)
                            {
                                Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 041204, string.Format("Blatt '{0}' Zeile {1} Spalte {2} enthält die Formel '{3}'", worksheet.Name, row, col, worksheet.Cells[row, col].Formula));
                            }
                            else
                            {
                                if (col > 1 && (backgroundColor != null))
                                Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 041203, string.Format("Blatt '{0}' Zeile {1} Spalte {2} hat die Farbe '{3}'", worksheet.Name, row, col, backgroundColor));
                            }
                        }
                    }

                    excelPackage.Save();
                }
            }
        }
        
        /// <summary>
        /// Schreibe aktuelles Datum in Excel NamedRange mit dem Namen xlNamedRangeName.
        /// </summary>
        /// <param name="xlWorkbookFilePath">Pfad zur zu beschreibenden Excel-Mappe.</param>
        /// <param name="xlNamedRangeName">Name des Zellbereich, in den das aktuelle Datum geschrieben werdne soll.</param>
        internal static void XlWriteDateToNamedRange(string xlWorkbookFilePath, string xlNamedRangeName) //Fehlernummern siehe Log.cs 0413ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041301, string.Format("XlWriteDateToNamedRange({0},{1})", xlWorkbookFilePath, xlNamedRangeName));

            if (File.Exists(xlWorkbookFilePath))
            {
                FileInfo file = new FileInfo(xlWorkbookFilePath);

                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    
                    try
                    {
                        // Suche Named Ranges in Worksheet 1
                        ExcelNamedRange namedRange = excelPackage.Workbook.Worksheets[1].Names[xlNamedRangeName];
                        namedRange.Value = DateTime.Now.ToLocalTime();
                        Log.Write(Log.Cat.Info, Log.Prio.Info, 041302, string.Format("Bezeichnung \"{1}\" in Blatt 1 in {0} gefunden bei {2}.", xlWorkbookFilePath, xlNamedRangeName, namedRange.Address));
                    }
                    catch
                    {
                        //Suche  Named Ranges in Workbook                        
                        try
                        {
                            ExcelNamedRange namedRange = excelPackage.Workbook.Names[xlNamedRangeName];
                            namedRange.Value = DateTime.Now.ToLocalTime();

                            Log.Write(Log.Cat.Info, Log.Prio.Info, 041303, string.Format("Bezeichnung \"{1}\" in Workbook in {0} gefunden bei {2}.", xlWorkbookFilePath, xlNamedRangeName, namedRange.Address));
                        }
                        catch
                        {
                            //Named Range nicht gefunden.
                            ExcelRange range = excelPackage.Workbook.Worksheets[1].Cells["B4"];
                            excelPackage.Workbook.Names.Add(xlNamedRangeName, range);
                            range.Value = DateTime.Now.ToLocalTime();
                            Log.Write(Log.Cat.ExcelRead, Log.Prio.Info, 041304, string.Format("Bezeichnung \"{1}\" nicht in Workbook {0} gefunden und deshalb neu erstellt.", xlWorkbookFilePath, xlNamedRangeName));
                        }
                    }
                                        
                    //Test: Im unteren Monatstabellen-Diagramm werden die Jaherszahlen im PDF manchmal nicht richtig angezeigt.
                    excelPackage.Workbook.Calculate();
                 
                    //save the changes
                    excelPackage.Save();
                }
            }
        }
        
        /// <summary>
        /// Zählt die Spalten mit der Überschrift "Uhrzeit".
        /// </summary>
        /// <param name="worksheet"></param>
        /// <returns>Anzahl der Spalten mit der Überschrift "Uhrzeit"</returns>
        public static int[] XlTimeColCount(ExcelWorksheet worksheet) //Fehlernummern siehe Log.cs 0414ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041401, string.Format("XlTimeColCount({0})",worksheet.Name));

            List<int> timeCols = new List<int>();
            for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
            {
                string valCol = worksheet.Cells[Excel.XlDayFileFirstRowToWrite - 1, col].Value?.ToString() ?? "";
                if (valCol == "Uhrzeit")
                {
                    timeCols.Add(col);
                }
            }
            int[] result = timeCols.ToArray();
            return result;
        }
        
        /// <summary>
        /// Gibt die zu beschreibende Zellposition mit row, column aus.
        /// </summary>
        /// <param name="timeCols">Nummern der Spalten mit Überschrift "Uhrzeit"</param>
        /// <param name="column">zu schreibende Spalte in Excel-Tabelle.(-1 = keine Spalte)</param>
        /// <returns>row: zu schreibende Zeile in Excel-Tabelle. (-1 = keine Zeile)</returns>
        private static int XlSetRowAndCol(int[] timeCols, out int column) //Fehlernummern siehe Log.cs 0415ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041501, string.Format("XlSetRowAndCol({0},{1}) - timeCols.Length ={2}", timeCols, "out int column", timeCols.Length));

            if (timeCols.Length < 1) { column = -1; return -1; }

            int row;
            int col = -1;

            if (timeCols.Length > 1) //TK-Räume mit viertelstündlicher Aufzeichnung
            {
                int min = DateTime.Now.Minute;
                int hour = DateTime.Now.Hour;
                row = 4 * hour + min / 15;
                if (row == 0) row = 96; // 00:00 Uhr ist in der Tabelle ganz am Ende.

                // Gehe in die vorherige Zeile, wenn die Minute zwischen lower und upper liegt ( => z.B. Schreibe 8:01 Uhr in Zeile für 8:00 und nicht in Zeile für 8:15 Uhr ).
                if (
                    Between(min, 0, XlPosOffsetMin) ||
                    Between(min, 15, 15 + XlPosOffsetMin) ||
                    Between(min, 30, 30 + XlPosOffsetMin) ||
                    Between(min, 45, 45 + XlPosOffsetMin)
                    ) --row;

                //Spalten werden nach rechts unter die Überschrift "Uhrzeit" versetzt. Zeilen fangen bei Versetzung wieder von oben an.
                if (row < 24)
                {
                    col = timeCols[0];
                    row += Excel.XlDayFileFirstRowToWrite;
                }
                else if (row < 48)
                {
                    col = timeCols[1];
                    row += Excel.XlDayFileFirstRowToWrite - 24;
                }
                else if (row < 72)
                {
                    col = timeCols[2];
                    row += Excel.XlDayFileFirstRowToWrite - 48;
                }
                else if (row < 96)
                {
                    col = timeCols[3];
                    row += Excel.XlDayFileFirstRowToWrite - 72;
                }
            }
            else //Kühlraum mit stündlicher Aufzeichnung
            {
                row = DateTime.Now.Hour;
                if (row == 0) row = 24; // 00:00 Uhr ist in der Tabelle ganz am Ende.
                row += Excel.XlDayFileFirstRowToWrite;
                col = 1;
                if (DateTime.Now.Minute < XlPosOffsetMin) { --row; } //5 min nach voll wird der Wert noch der vorherigen Stunde hinzugefügt.
                else if (DateTime.Now.Minute < 60 - XlNegOffsetMin) { col = -1; row = -1; } // Wenn es vor 5 min vor Voll ist, wird kein Wert geschrieben. 

                // Wenn XlPosOffsetMin==30 && XlNegOffsetMin==30 wird um 0:30 Uhr und 0:45 Uhr in Zeile 25 (MIN-Wert) geschrieben! (neu 20.01.2020)
                if (row > 24 + Excel.XlDayFileFirstRowToWrite) row = Excel.XlDayFileFirstRowToWrite;
            }

            column = col;
            return row;
        }
        
        /// <summary>
        /// Schließt die geöffnete Excle-Mappe und wartet, bis dies beendet ist. Räumt COM-Objekte nicht auf! 
        /// Beim erneuten Start der Excel-Mappe wird ein Wiederherstellen-Dialog angezeigt, der aber ignoriert werden kann.
        /// </summary>
        /// <param name="process"></param>
        /// <param name="filePath"></param>
        internal static void KillExcel(Process process, string filePath) //Fehlernummern siehe Log.cs 0416ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041601, string.Format("KillExcel({0},{1})",process,filePath));

            if (process.MainWindowTitle.Contains(Path.GetFileName(filePath)))
            {
                Log.Write(Log.Cat.FileSystem, Log.Prio.Warning, 041602, string.Format("Die Datei {0} ist bereits geöffnet und wird jetzt geschlossen.", Path.GetFileName(filePath)));
                process.Kill();
                process.WaitForExit();
            }
        }

        /// <summary>
        /// Setzt den Blattschutz für jedes einzelne Tabellenblatt.
        /// </summary>
        /// <param name="xlFilePath">Pfad zur Excel-Datei</param>
        /// <param name="Password">Passwort für den Blattschutz</param>
        private static void ProtectSheets(string xlFilePath, string password) //Fehlernummern siehe Log.cs 0417ZZ
        {
            if (password.Length < 3) return;

            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 041701, string.Format("Setze Blattschutz für {0}", Path.GetFileNameWithoutExtension(xlFilePath)));

            FileInfo file1 = new FileInfo(xlFilePath);
            using (ExcelPackage excelPackage = new ExcelPackage(file1))
            {
                // excelPackage.Encryption.IsEncrypted = true;
                excelPackage.Workbook.Properties.Keywords += XlPasswordEncrypted;

                foreach (ExcelWorksheet sheet in excelPackage.Workbook.Worksheets)
                {
                    sheet.Protection.SetPassword(password);
                    sheet.Protection.IsProtected = true;
                }

                excelPackage.Save();
            }
        }

        #endregion
    }
}

//
//    /*							                                                                                    Produktfühler				        Raumfühler			
//    Chargen-nummer	Schock-kühler	Startzeit	Endzeit	    Dauer	Pro-gramm	"Sollwert Max.-dauer"	Start	Ende	Min	        Max	        Start	Ende	Min	    Max
//    Ist	            Ist	            Ist	        Ist	        Ist	    Soll	    Soll	                Ist	    Ist	    Ist	        Ist	        Ist	    Ist	    Ist	    Ist
//    Nr	            Nr	            hh:mm:ss	hh:mm:ss	hh:mm	Nr	        min	                    °C	    °C	    °C	        °C	        °C	    °C	    °C	    °C
//    */
//      ChargenNrS      SchockNr        >Zeit1<      >Zeit2<    >Z2-Z1<  SchockRProg1    SchockSollDauer    A01..   A01..  SchockMin1K, SchockMax1K,A01..,  A01.., SchockMin1, SchockMax1
//                                                                       SchockRProg2
//                                                                           ...   

// SchockRProg1 = AOLDB301_DBW404
// ChargenNrS = Text( $Day, "00") + Text( $Month, "00") + Text(A01_DB301_DBW306,
// SchockRChargel S = ChargenNrS; 
// SchockAStarti = $Hour • 60 + $Minute; Anzeige Start and Hist Linge 
// SchockRStartEll = EndlosMin; Fiir Druck Ja/Nein 
// SchockSollDauer = A01_DB301_DBW420r 
// SchockRSollDauerl = SchockSollDauer; kir RealTrend }
// SchockNr =1; { kir Obergabe an Excel } IF NOT A00_Statistik.Alarm THEN Macro = "hm15_StartSchocV; Command = "[Run(" + StringChar(34) + Macro + StringChar(34) + ", 0)]"; WWExecute["excel", "system", Command]; ENDIF; 
// SchockMin1 = A01_DB301_DBD20; 
// SchockMax1 = A01_DB301_DBD20; 
// SchockMin1K = A01_DB301_DBD60; 
// SchockMax1K = A01_DB301_DBD60; 