using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using OfficeOpenXml;
using System.Diagnostics;
using System.Data;

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


    internal static class Excel
    {
        #region Fields for WriteToExcelFiles
        // Minuten, die zwischen XlPosOffsetMin und 60 - XlNegOffsetMin liegen, werden nicht für Stundenwerte aufgezeichnet.
        private static int xlPosOffsetMin = 5; //min. nach Voller (Viertel-)Stunde, die noch zu der vorherigen (Viertel-)Stunde zählen.
        public static int XlPosOffsetMin { get => xlPosOffsetMin; set => xlPosOffsetMin = value; }

        private static int xlNegOffsetMin = 5; //min. vor Voller Stunde, die zu der kommenden vollen Stunde zählen.
        public static int XlNegOffsetMin { get => xlNegOffsetMin; set => xlNegOffsetMin = value; }

        public static string XlTemplateDayFilePath { get; set; } = @"D:\XlLog\T_vorl.xlsx";
        public static string XlTemplateMonthFilePath { get; set; } = @"D:\XlLog\M_vorl.xlsx";
        public static int XlDayFileFirstRowToWrite { get; set; } = 10;
        public static int XlMonthFileFirstRowToWrite { get; set; } = 8;
        public static string XlArchiveDir { get; set; } = @"D:\Archiv";
        internal static string XlPassword { get; set; } = "henryk";

        #endregion

        #region WriteToExcelFiles

        /// <summary>
        /// Schreibt die nächste Zeile in Excel-Datei
        /// </summary>
        public static void XlFillWorkbook()
        {
            Log.Write(Log.Category.MethodCall, 1907261320, string.Format("XlFillWorkbook()"));

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

                #region In Excel-Tabellen schreiben / PDF erzeugen
                
                //Lese InTouch-TagNames aus Vorlage-Datei
                List<List<string>> items = XlReadRowValues(Excel.XlTemplateDayFilePath, Excel.XlDayFileFirstRowToWrite);

                switch (Program.AppStartedBy)
                {
                    case "Schock":
                        XlWriteShockFreezer(xlDayFilePath, items);
                        break;
                    case "AlmDruck":
                        Sql.AlmListToExcel(true);
                        break;
                    case "PdfDruck":
                        Pdf.CreatePdfFromCmd();
                        break;
                    case "Uhrstellen":
                        SetNewSystemTime.SetNewSystemtimeAndScheduler(Program.CmdArgs[1]);
                        break;
                    case "Monatsdatei":
                        string pathMonthFile = Program.CmdArgs[1];
                        string xlDayFilesDir = Program.CmdArgs[2];
                        Excel.XlNewMonthFileWriteRow(pathMonthFile, xlDayFilesDir);
                        break;
                    default:
                        //Wenn der PDF-File von gestriger Tagesdatei nicht existiert und die Zeit Pdf.PdfConvertStartHour erreicht ist fehlende PDFs erstellen
                        if (DateTime.Now.Hour >= Pdf.PdfConvertStartHour)
                        {
                            string pdfYesterdayFilepath = Path.ChangeExtension(xlYesterdayFilePath, ".pdf");

                            if (!File.Exists(pdfYesterdayFilepath))
                            {
                                Log.Write(Log.Category.PdfWrite, 2002190825, "Erzeuge PDF " + pdfYesterdayFilepath);
                                Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-1, false, 0, true));

                                //Wenn XlLog nicht am ersten des Monats ausgeführt wird, muss der Vormonat nochmal erstellt werden.
                                if (DateTime.Now.Day > 1)
                                {
                                    Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-28, false, 0, true));
                                }
                            }
                        }
                        else
                        {
                            Log.Write(Log.Category.PdfWrite, 2002190826, "PDF-Erzeugung erst ab " + Pdf.PdfConvertStartHour + " Uhr.");
                        }

                        //in Excel-Mappe schreiben; ggf. neue Excel-Mappe erstellen
                        XlWriteToDayFile(xlDayFilePath, items);
                        break;
                }         
                #endregion

                #region Monatstabelle + Monats-PDF schreiben 
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

                    Log.Write(Log.Category.Info, 1908010951, string.Format("Freigabe Monatsdatei schreiben: {0}", writeToMonthFilePermission) );
                    //Log.Write(Log.Category.Info, 1908010951, string.Format("Monatsdatei beschreiben: {4}, denn; Tag der letzen Änderung < Heute = {0}; Es ist der 1. des Monats und die letzte Änderung war vorher am {1}.  = {2}, Monatsdatei existiert: {3}", bed1, monthFileLastModified.Day, bed2, !bed3, writeToMonthFilePermission));

                    if (writeToMonthFilePermission)
                    {
                        string xlDayFilesDir = CeateXlFilePath(-1, false, 0, true);
                        Excel.XlNewMonthFileWriteRow(pathMonthFile, xlDayFilesDir);

                        //PDF erstellen.
                        Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-1, true, 0, true));

                        //Wenn am ersten des Monats keine Monatsdatei erzeugt wurde
                        if (bed3 && DateTime.Now.Day > 1)
                        {
                            //Prüfe auch den Vormonat
                            xlDayFilesDir = CeateXlFilePath(-28, false, 0, true);
                            Excel.XlNewMonthFileWriteRow(pathMonthFile, xlDayFilesDir);
                            Pdf.CreatePdf4AllXlsxInDir(CeateXlFilePath(-28, true, 0, true));
                        }
                    }                    
                }
                catch (Exception ex)
                {
                    Log.Write(Log.Category.ExcelWrite, -903081303, string.Format("Fehler beim erstellen der Excel-Monatsdatei : Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                }
                #endregion

            }
            catch (IOException)
            {
                Log.Write(Log.Category.ExcelWrite, -905021025, string.Format("Die Datei {0} ist bereits geöffnet. Es wird nicht versucht erneut zu schreiben.", xlDayFilePath));
                Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelWrite, -902011251, string.Format("Fehler beim erstellen der Excel-Tagesdatei : Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                InTouch.SetExcelAliveBit(true);
            }
        }

        /// <summary>
        /// Neue Excel-Datei von Vorlage erstellen.
        /// </summary>
        /// <param name="xlTemplateFilePath">Pfad zur Vorlagedatei.</param>
        /// <param name="xlWorkingWorkbookFilePath">Pfad der zu ertsellenden Datei.</param>
        /// <returns>true = Datei erfolgreich erstellt.</returns>
        private static bool XlTryCreateWorkbookFromTemplate(string xlTemplateFilePath, string xlWorkingWorkbookFilePath, int firstRowToWrite)
        {
            Log.Write(Log.Category.MethodCall, 1907261321, string.Format("XlTryCreateWorkbookFromTemplate({0},{1},{2})", xlTemplateFilePath, xlWorkingWorkbookFilePath, firstRowToWrite));

            try
            {
                if (!File.Exists(xlWorkingWorkbookFilePath))
                {
                    if (!File.Exists(xlTemplateFilePath))
                    {
                        if (Path.GetFileNameWithoutExtension(xlWorkingWorkbookFilePath).Substring(0, 1) == "M")
                        {
                            Log.Write(Log.Category.FileSystem, 2002171623, string.Format("Keine Monatsdatei-Vorlage gefunden für {0}", xlTemplateFilePath));
                            //kein Fehler
                            return false;
                        }
                        else
                        {
                            Log.Write(Log.Category.FileSystem, -902041123, string.Format("Es kann keine Excel-Mappe erstellt werden. Die Vorlagedatei für die Tagesdatei wurde nicht gefunden: {0}", xlTemplateFilePath));
                            Program.AppErrorOccured = true;
                            return false;
                        }
                    }

                    //Kopieren von Vorlage
                    Log.Write(Log.Category.ExcelWrite, 1902011259, string.Format("Erstelle Datei: {0} aus {1}", xlWorkingWorkbookFilePath, xlTemplateFilePath));
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
                }
                return true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelWrite, -902251533, string.Format("Fehler Erstellen der Excel-Datei von Vorlage: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
                return false;
            }
        }

        /// <summary>   
        /// Sucht die Monatstabelle 1 Jahr vor xlThisMonthFilePath und kopiert von dort die Werte in die Vorjahresspalte von xlThisMonthFilePath.
        /// Setzt die Jahreszahlen in die Spaltenüberschriften von xlThisMonthFilePath.
        /// </summary>
        /// <param name="xlThisMonthFilePath">Pfad zur aktuellen Monats-Datei.</param>
        internal static void XlCopyMonthValuesFromLastYear()
        {
            Log.Write(Log.Category.MethodCall, 1907261322, string.Format("XlCopyMonthValuesFromLastYear()"));

            int row = Excel.XlMonthFileFirstRowToWrite;
            string xlThisMonthFilePath = CeateXlFilePath(-1, true);
            string xlLastYearMonthFilePath = CeateXlFilePath(-1, true, -1);
            List<Tuple<int, int>> tupleList = new List<Tuple<int, int>>();

            try
            {                
                if (!File.Exists(xlThisMonthFilePath) || !File.Exists(xlLastYearMonthFilePath) || !File.Exists(Excel.XlTemplateMonthFilePath))
                {
                    Log.Write(Log.Category.ExcelRead, 1902271506, string.Format("Letztjahreswerte konnten nicht in die Datei {0} eingetragen werden. Es fehlen Quelldateien.\r\n\t\t\t\tLetztes Jahr vorhanden: {1}\r\n\t\t\t\tMonatsvorlage vorhanden: {2}\r\n\t\t\t\taktuelle Monatsdatei vorhanden: {3}", xlThisMonthFilePath, File.Exists(xlLastYearMonthFilePath), File.Exists(Excel.XlTemplateMonthFilePath), File.Exists(xlThisMonthFilePath)));
                    //kein Fehler
                    return;
                }

                //Finde Spalten in Template zum kopieren                
                FileInfo file1 = new FileInfo(Excel.XlTemplateMonthFilePath);
                using (ExcelPackage excelPackage1 = new ExcelPackage(file1))
                {
                    foreach (ExcelWorksheet worksheet in excelPackage1.Workbook.Worksheets)
                    {
                        Log.Write(Log.Category.ExcelRead, 2001211231, "XlCopyMonthValuesFromLastYear() Blatt " + worksheet.Name);

                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                        {
                            //Log.Write(Log.Category.ExcelRead, 2001211232, "Spalte " + col + " = '" + worksheet.Cells[row, col].Value + "'");
                            //Wenn Zelle keinen Farbhintergrund hat und nicht leer ist -> TagName gefunden - (neu 21.01.2020) Wenn Zelle darüber eine Formel hat, die mit '=Datum..' beginnt, in die Tuple aufnehmen.
                            if (worksheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb == null && worksheet.Cells[row, col].Value != null && worksheet.Cells[row - 1, col].Formula.StartsWith("Datum") )
                            {
                                tupleList.Add(new Tuple<int, int>(worksheet.Index, col));
                                //Log.Write(Log.Category.ExcelRead, 1903081317, string.Format("Kopiere Letztjahreswerte in Monatsdatei:\tBlatt: {0}\t Spalte: {1}\tWert: {2}[{3}]", worksheet.Index, col, worksheet.Cells[row, col].Value, worksheet.Cells[row - 1, col].Formula) );
                            }
                        }

                        
                    }
                }

            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelWrite, -001211324, string.Format("Fehler beim Lesen der TagNames aus Excel-Monats-Datei-Vorlage für Übernahme der Vorjahreswerte: {0} \r\n\t\t  Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3} \r\n\t\t StackTrace: {4}", 0, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
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
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelWrite, -903071417, string.Format("Fehler beim Kopieren der Vorjahreswerte in die Excel-Monats-Datei : {0} \r\n\t\t  Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3} \r\n\t\t StackTrace: {4}", 0, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
            }
        }
    
        /// <summary>
        /// Wartet auf InTouch (D.S. $Minute), bildet ggf. Mittelwerte, schreibt in Tagesdatei, setzt Mittelwerte/Minutenzähler zurück.
        /// </summary>
        /// <param name="xlDayFilePath">Pfad zur Tagesdatei.</param>
        /// <param name="items">Liste der zu lesenden TagNames.</param>
        public static void XlWriteToDayFile(string xlDayFilePath, List<List<string>> items)
        {
            Log.Write(Log.Category.MethodCall, 1907261323, string.Format("XlWriteToDayFile({0},{1})", xlDayFilePath, "*content*"));

            //Warte auf Berechnung in InTouch (D.S. $Minute) und Bestimme Mittelwerte.
            if (Program.AppStartedBy == "Task")
            {
                Log.Write(Log.Category.InTouchVar, 1907161616, "Warte auf InTouch-Skripte.");
                Tools.Wait(Tools.WaitForScripts);
            }

            //     Beispiel: true  =                          10:56    ,     4               ,(60 - 5 ) = 55   
            //     Beispiel: false =                          10:53    ,     4               ,(60 - 5 ) = 55    
            bool fullHourTimeframe = !Excel.Between(DateTime.Now.Minute, Excel.XlPosOffsetMin, 60 - Excel.XlNegOffsetMin);

            if (fullHourTimeframe)
            {
                // Wenn Zeit um Stundensprung, ermittle Mittelwerte in InTouch; Setzte ExBestStdWerte
                InTouch.WriteDiscTag("ExBestStdWerte", true);
            }

            XlDayFileWriteRange(xlDayFilePath, items);

            Tools.Wait(1);

            if (fullHourTimeframe)
            {
                //Edit: Im D. S. $Hour wird gesetzt   ExLöscheStdMin = 1;
                InTouch.WriteDiscTag("ExLöscheStdMin", true);
            }
       
            //Viertelstundenwerte auch bei Stundensprung zurücksetzen!
            InTouch.WriteDiscTag("ExLösche15StdMin", true);
            
        }

        /// <summary>
        /// Schreibt eine Zeile in die Excel-Datei xlFilePath, wenn das Tabellenblatt mindestens eine Spalte "Uhrzeit" enthält".
        /// </summary>
        /// <param name="xlFilePath"></param>
        /// <param name="content">2-Dimensionale Liste der Tabellenblätter und Variablen</param>
        public static void XlDayFileWriteRange(string xlFilePath, List<List<string>> content)
        {
            Log.Write(Log.Category.MethodCall, 1907261324, string.Format("XlDayFileWriteRange({0},{1})", xlFilePath, "*content*"));

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
                            Log.Write(Log.Category.ExcelRead, 2001211611, string.Format("Blatt {0} aus {1} hat keine Zeitspalte.", worksheetNo, Path.GetFileNameWithoutExtension(xlFilePath) ) );
                        }
                        else
                        {
                            #region Zeile in Tabellenblatt Kühlraum oder TK-Raum schreiben
                            int row = XlSetRowAndCol(timeCols, out int col);
                            if (col < 1 || row < 1)
                            {
                                // Es muss in dieses Blatt nichts geschrieben werden.
                                Log.Write(Log.Category.ExcelRead, 2001211612, string.Format("Keine Einträge für Blatt {0} (Zeile:{1},Spalte:{2}) aus {3}", worksheetNo, row, col, timeCols.Length));
                                continue;
                            }

                            foreach (string item in wsContent) //Zeile füllen
                            {
                                if (item.Length < 1 || item == null)
                                {
                                    //Wenn der Spalte keine Variable zugeordnet wurde
                                    worksheet.Cells[row, col].Value = null;
                                }
                                else
                                {
                                    if (item != "DoNotChangeCell")
                                    {
                                        var result = InTouch.ReadTag(item);
                                        // Wenn Variable nicht vorhanden ist, wird von Intouch float.MaxValue ausgegeben.
                                        if (Convert.ToSingle(result) != float.MaxValue)
                                        {
                                            worksheet.Cells[row, col].Value = result;
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

                    Log.Write(Log.Category.Info, 1902080909, string.Format("In {0} wurden in {1} Tabellenblätter {2} Werte bearbeitet.", Path.GetFileName(xlFilePath), content.Count, content.SelectMany(list => list).Distinct().Count()));
                }
            }
            catch(InvalidOperationException)
            {
                Log.Write(Log.Category.ExcelWrite, -002190920, string.Format("Die Excel-Datei {0} konnte nicht beschrieben werden. Sie ist vermutlich durch ein anderes Programm geöffnet.", Path.GetFileNameWithoutExtension(xlFilePath) ) );
                Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelWrite, -902011300, string.Format("Fehler beim schreiben in die Excel-Datei: {0} \r\n\t\t\t\t Typ: {1} \r\n\t\t\t\t Fehlertext: {2}  \r\n\t\t\t\t InnerException: {3} \r\n\t\t\t\t StackTrace: {4}", xlFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Schreibe in Monatstabelle die Tagesummen von allen Tagen des Monats.
        /// </summary>
        public static void XlNewMonthFileWriteRow(string xlMonthFilePath, string xlDayFilesDir)
        {
            Log.Write(Log.Category.MethodCall, 1907261325, string.Format("XlNewMonthFileWriteRow({0}, {1})", xlMonthFilePath, xlDayFilesDir));

            //Datei für Monatstabelle erzeugen            
            if (!XlTryCreateWorkbookFromTemplate(Excel.XlTemplateMonthFilePath, xlMonthFilePath, Excel.XlMonthFileFirstRowToWrite)) return;
            Tools.Wait(2);
            
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

                foreach (FileInfo xlDayFileInfo in xlDayFileList)
                {
                    //Überspringe temporäre Dateien
                    if (xlDayFileInfo.Name.StartsWith("~")) continue;

                    //Nehme den Tag aus dem Dateinamen, da Erstelldatum nicht eindeutig ist (T01... wird erst am 2. Tag um 0:15 Uhr erzeugt.)!
                    int.TryParse(xlDayFileInfo.Name.Substring(2, 2), out int dayNo); // Dateiname-Muster: M_ddmmyyyy.xlsx
                    //Log.Write(Log.Category.FileSystem, 2001201441, "Lese Tabelle von Tag Nr. " + dayNo);
                    if (dayNo < DateTime.Now.Day || DateTime.Now.Day == 1) // am 1. des Monats alle Dateien des Vormonats zusammenfassen.
                    {
                        List<List<string>> valuesDay = XlReadRowValues(xlDayFileInfo.FullName, sumRowNo, true); 
                        allDayValues.Add(dayNo, valuesDay);
                    }
                }
                #endregion

                FileInfo file = new FileInfo(xlMonthFilePath);
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    foreach (var dayValues in allDayValues) // Tagesdateien in Monatsordner
                    {
                        int worksheetNo = 0;
                        int row = dayValues.Key + Excel.XlMonthFileFirstRowToWrite - 1;
                        Log.Write(Log.Category.ExcelRead, 1905021631, string.Format("Schreibe Zeile {0:00} in Monatsdatei {1}", row,  Path.GetFileNameWithoutExtension(xlMonthFilePath) ));

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
                                        Log.Write(Log.Category.ExcelRead, 1910220801, string.Format("Der TagName '{0}' konnte nicht im Blatt {1}, Zeile {2}, Spalte {3} in der Tagesvorlage gefunden werden und wird übersprungen.", item, worksheetNo, row, col));
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
                                Log.Write(Log.Category.ExcelRead, -910220727, "Fehler in foreach (string item in WorksheetsItemsM) \r\n\t" + exTest.Message + " \r\n\t" + exTest.StackTrace);
                                Program.AppErrorOccured = true;
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

                Log.Write(Log.Category.ExcelWrite, 1902271532, "Habe in Monatsdatei geschrieben.");

                //PDF erstellen.
                Pdf.CreatePdf(xlMonthFilePath);
            }
            catch (System.ArgumentException exArg)
            {
                Log.Write(Log.Category.ExcelRead, -904241240, "Die Namen der Tagesdateien im Monatsordner sind nicht eindeutig für\r\n\t\t\t\t" +xlMonthFilePath + "\r\n\t\t\t\tDie Monatsdatei kann nicht erstellt werden.\r\n\t\t\t\t" + exArg.Message);
                Program.AppErrorOccured = true;
            }
            catch (IndexOutOfRangeException index_ex)
            {
                Log.Write(Log.Category.ExcelRead, -905081725, "Arrayfehler in Vorlagedatei. Mögliche Ursache: Farbhintergründe im Wertefeld in T_vorl.xlsx oder M_vorl.xlsx\r\nFehlermeldung: " + index_ex.Message + "\r\n\t" + index_ex.InnerException + "\r\n\t" + index_ex.StackTrace );
                Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelWrite, -902011251, string.Format("Fehler beim erstellen der Excel-Monats-Datei : xlMonthFilePath {0} \r\n\t\t\t\tTyp: {1} \r\n\t\t\t\tFehlertext: {2}  \r\n\t\t\t\tInnerException: {3} \r\n\t\t\t\tStackTrace: {4}", xlMonthFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Schreibt eine Zeile in die Excel-Datei xlFilePath, wenn das Tabellenblatt keine Spalte "Uhrzeit" enthält.
        /// </summary>
        /// <param name="xlFilePath">Pfad zur Excel-Mappe</param>
        /// <param name="content">Liste der Tabellenblätte rmit Liste der zu schreibenden Variablen</param>
        public static void XlWriteShockFreezer(string xlFilePath, List<List<string>> content)
        {
            Log.Write(Log.Category.MethodCall, 1907261326, string.Format("XlWriteShockFreezer({0}, *content*)", xlFilePath));

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
                        Log.Write(Log.Category.ExcelShock, 1902141620, string.Format("In Blatt {0} gibt es {1} Zeilen. Erste Leerzeile: {2}", worksheetNo, rangeMaxRows, firstEmptyRow));

                        //Liste alle Werte der Spalte Schockkühler-Nr. auf.
                        ExcelRange shockNoCol = worksheet.Cells[Excel.XlDayFileFirstRowToWrite, (int)XlShockCol.ShockNo, firstEmptyRow, (int)XlShockCol.ShockNo];
                        List<int> shockNoList = new List<int>();
                        
                        foreach (var cell in shockNoCol)
                        {
                            shockNoList.Add(cell.GetValue<int>());
                        }
                        #endregion

                        #region Lese InTouch-Werte

                        NativeMethods intouch = new NativeMethods(0, 0);

                        int SchockNr = intouch.ReadInteger("SchockNr");
                        if (SchockNr < 1)
                        {
                            Log.Write(Log.Category.InTouchVar, 1903151201, "InTouch-Tag \"SchockNr\": Es wurde kein Schockkkühler definiert.");
                            break; // Es wurde kein Schockkühler gesetzt.
                        }
                        #endregion

                        #region Starte/beende Schockkühlvorgang; schreibe in Excel-Tabelle 

                        // Finde Zeile mit letztem Auftreten von SchockNr
                        int currentRowNo = shockNoList.LastIndexOf(SchockNr) + XlDayFileFirstRowToWrite;
                        if (currentRowNo < XlDayFileFirstRowToWrite) currentRowNo = firstEmptyRow; // Dieser Schockkühler wurde noch nicht aufgezeichnet: gehe in erste leere Zeile.

                        bool startTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Value == null;
                        bool endTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.EndTime].Value == null;

                        // Wenn Start- und End-Zeit schon gefüllt sind, gehe in erste leere Zeile
                        if (!startTimeIsEmpty && !endTimeIsEmpty)
                        {
                            currentRowNo = firstEmptyRow;
                            startTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Value == null;
                            endTimeIsEmpty = worksheet.Cells[currentRowNo, (int)XlShockCol.EndTime].Value == null;
                            //Log.Write(0, -14, string.Format("B) Änderung: currentRowNo: >{0}<", currentRowNo));
                        }

                        //LeseTagNames für Schockkühler-Temperaturen
                        string[] ProductTempTagNames = wsContent[(int)XlShockCol.ProductStartTemp - 1].Split(';');
                        string[] RoomTempTagNames = wsContent[(int)XlShockCol.RoomStartTemp - 1].Split(';');

                        // Wenn Anfangszeit leer ist: neuen Schockkühlvorgang starten
                        if (startTimeIsEmpty)
                        {
                            string ChargenNrS = intouch.ReadString("ChargenNrS");
                            int SchockRProgX = intouch.ReadInteger("SchockRProg" + SchockNr);
                            int SchockSollDauer = intouch.ReadInteger("SchockSollDauer");
                            float ProductStartTemp = intouch.ReadFloat(ProductTempTagNames[SchockNr - 1]);
                            float RoomStartTemp = intouch.ReadFloat(RoomTempTagNames[SchockNr - 1]);

                            Log.Write(Log.Category.ExcelShock, 1902141621, string.Format("Neuer Vorgang wird in Zeile {0} geschrieben: Schockkühler {1}, Charge {2}, Programm {3}, Soll-Dauer {4} min", currentRowNo, SchockNr, ChargenNrS, SchockRProgX, SchockSollDauer));

                            worksheet.Cells[currentRowNo, (int)XlShockCol.ChargeNo].Value = ChargenNrS;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ShockNo].Value = SchockNr;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Value = DateTime.Now.ToLongTimeString();
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ShockProgramNo].Value = ( SchockRProgX == int.MaxValue ) ? (object)"TagName?" : SchockRProgX ;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.MaxDuration].Value = (SchockSollDauer == int.MaxValue) ? (object)"TagName?" : SchockSollDauer;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.ProductStartTemp].Value = (ProductStartTemp == float.MaxValue) ? (object)"TagName?" : ProductStartTemp;
                            worksheet.Cells[currentRowNo, (int)XlShockCol.RoomStartTemp].Value = (RoomStartTemp == float.MaxValue) ? (object)"TagName?" : RoomStartTemp;
                        }

                        // Wenn Anfangszeit gesetzt und Endzeit leer ist: Schockkühlvorgang beenden
                        if (!startTimeIsEmpty && endTimeIsEmpty)
                        {
                            float ProductEndTemp = intouch.ReadFloat(ProductTempTagNames[SchockNr - 1]);
                            float RoomEndTemp = intouch.ReadFloat(RoomTempTagNames[SchockNr - 1]);
                            float SchockMinX = intouch.ReadFloat("SchockMin" + SchockNr);
                            float SchockMaxX = intouch.ReadFloat("SchockMax" + SchockNr);
                            float SchockMinXK = intouch.ReadFloat("SchockMin" + SchockNr + "K");
                            float SchockMaxXK = intouch.ReadFloat("SchockMax" + SchockNr + "K");

                            Log.Write(Log.Category.ExcelShock, 1902141623, string.Format("Vorgang in Zeile {0} beendet: Schockkühler {1}, ", currentRowNo, SchockNr));

                            worksheet.Cells[currentRowNo, (int)XlShockCol.EndTime].Value = DateTime.Now.ToLongTimeString();
                            TimeSpan startTime = DateTime.Parse(worksheet.Cells[currentRowNo, (int)XlShockCol.StartTime].Text).TimeOfDay;
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
                    Log.Write(Log.Category.ExcelShock, -902111629, string.Format("Fehler beim Schreiben in die Excel-Datei (Schockkühler): \r\n\t\t Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2} \r\n\t\t StackTrace: {3} \r\n\t\t GetBaseException().Message: {4}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace, ex.GetBaseException().Message));
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
        public static bool Between(this int num, int lower, int upper, bool inclusive = true)
        {
            Log.Write(Log.Category.MethodCall, 1907261327, string.Format("Between({0},{1},{2},{3})",num, lower, upper,inclusive));

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
        public static int XlGetRowInColByValue(string xlFilePath, int worksheetno, int col, string searchValue)
        {
            Log.Write(Log.Category.MethodCall, 1907261328, string.Format("XlGetRowInColByValue({0},{1},{2},{3},)", xlFilePath, worksheetno, col, searchValue));

            try
            {
                if (!File.Exists(xlFilePath))
                {
                    Log.Write(Log.Category.ExcelRead, -902261330, string.Format("Datei nicht gefunden: {0}", xlFilePath));
                    Program.AppErrorOccured = true;
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
                    for (int row = worksheet.Dimension.End.Row; row >= Excel.XlMonthFileFirstRowToWrite; row--)
                    {
                        if (worksheet.Cells[row, col].Value.ToString() == searchValue.ToString())
                        {
                            return row;
                        }
                    }
                }

                Log.Write(Log.Category.ExcelRead, 1902261334, string.Format("Wert >{0}< wurde nicht gefunden in Datei {1}, Blatt {2}, Spalte {3}.", searchValue, xlFilePath, worksheetno, col));
                return -1;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelRead, -902261215, string.Format("Fehler beim finden der Zeile in Excel-Datei: {0} \r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}  \r\n\t\t StackTrace: {4}", xlFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
                return -1;
            }
        }

        /// <summary>
        /// Liest eine komplette Reihe in einem Excel-Sheet
        /// </summary>
        /// <param name="xlFilePath">Pfad zum Excel-Sheet</param>
        /// <param name="worksheetNo">Lfd. Nr. des Tabellenblatts</param>
        /// <param name="row">Nummer der zu lesenden Zeile </param>
        /// <returns>Liste der Zellenwerte in dieser Reihe. Listen-Eintrag "DoNotChangeCell" markiert, dass Zellwerte beim Schreiben der Liste nicht geändert werden sollen.</returns>
        public static List<List<string>> XlReadRowValues(string xlFilePath, int row, bool recordBGColoredCells = false)
        {
            Log.Write(Log.Category.MethodCall, 1907261329, string.Format("XlReadRowValues({0},{1},{2})", xlFilePath, row, recordBGColoredCells));

            try
            {
                if (!File.Exists(xlFilePath))
                {
                    Log.Write(Log.Category.ExcelRead, -902250952, string.Format("Datei nicht gefunden: {0}", xlFilePath));
                    Program.AppErrorOccured = true;
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
                        excelData.Add(new List<String>()); //Adds new sub List
                        ++wsCount;

                        //loop all columns in a row
                        for (int col = worksheet.Dimension.Start.Column; col <= worksheet.Dimension.End.Column; col++)
                        {
                            var val = worksheet.Cells[row, col].Value;

                            //add the cell data to the List 
                            if (val == null)
                            {
                                excelData[wsCount].Add("");
                            }
                            else
                            {
                                //Wenn keine Hintergrundfarbe gesetzt ist
                                if (worksheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb == null || recordBGColoredCells)
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
                Log.Write(Log.Category.ExcelRead, -002191008, string.Format("Die Excel-Datei {0} konnte nicht beschrieben werden. Sie ist vermutlich durch ein anderes Programm geöffnet.", Path.GetFileNameWithoutExtension( xlFilePath ) ) );
                Program.AppErrorOccured = true;
                return null;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.ExcelRead, -902011257, string.Format("Fehler beim Lesen aus Excel-Datei: {3} \r\n\t\t\t\t Typ: {0} \r\n\t\t\t\t Fehlertext: {1}  \r\n\t\t\t\t InnerException: {2}  \r\n\t\t\t\t StackTrace: {4}", ex.GetType().ToString(), ex.Message, ex.InnerException, xlFilePath, ex.StackTrace));
                Program.AppErrorOccured = true;
                return null;
            }

        }
        
        /// <summary>
        /// XLSX-Filename + Ordner für Jahr/Monat + Pfad erzeugen.
        /// </summary>
        /// <param name="OffsetDays">Tage von heute (negative Werte liegen in der Vergangenheit).</param>
        /// <returns>Pfad zur Excel-Tages- oder Monats-Datei oder Ordner.</returns>
        public static string CeateXlFilePath(int OffsetDays = 0, bool MonthFile = false, int OffsetYears = 0, bool directoryOnly = false)
        {
            Log.Write(Log.Category.MethodCall, 1907261330, string.Format("CeateXlFilePath({0},{1},{2},{3})", OffsetDays, MonthFile, OffsetYears , directoryOnly ));

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
        private static void XlDeleteRowValuesExceptBackgroundcolor(string xlWorkingWorkbookFilePath, int row)
        {
            Log.Write(Log.Category.MethodCall, 1907261331, string.Format("XlDeleteRowValuesExceptBackgroundcolor({0},{1})", xlWorkingWorkbookFilePath, row));

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
                            //Wenn keine Hintergrundfarbe gesetzt ist
                            if (worksheet.Cells[row, col].Style.Fill.BackgroundColor.Rgb == null)
                            {
                                //Setze Schriftfarbe auf schwarz.
                                worksheet.Cells[row, col].Style.Font.Color.SetColor(System.Drawing.Color.Black);

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
        internal static void XlWriteDateToNamedRange(string xlWorkbookFilePath, string xlNamedRangeName)
        {
            Log.Write(Log.Category.MethodCall, 1907261332, string.Format("XlWriteDateToNamedRange({0},{1})", xlWorkbookFilePath, xlNamedRangeName));

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
                        Log.Write(Log.Category.Info, 1903081251, string.Format("Bezeichnung \"{1}\" in Blatt 1 in {0} gefunden bei {2}.", xlWorkbookFilePath, xlNamedRangeName, namedRange.Address));
                    }
                    catch
                    {
                        //Suche  Named Ranges in Workbook                        
                        try
                        {
                            ExcelNamedRange namedRange = excelPackage.Workbook.Names[xlNamedRangeName];
                            namedRange.Value = DateTime.Now.ToLocalTime();

                            Log.Write(Log.Category.Info, 1903081252, string.Format("Bezeichnung \"{1}\" in Workbook in {0} gefunden bei {2}.", xlWorkbookFilePath, xlNamedRangeName, namedRange.Address));
                        }
                        catch
                        {
                            //Named Range nicht gefunden.
                            ExcelRange range = excelPackage.Workbook.Worksheets[1].Cells["B4"];
                            excelPackage.Workbook.Names.Add(xlNamedRangeName, range);
                            range.Value = DateTime.Now.ToLocalTime();
                            Log.Write(Log.Category.ExcelRead, 1903081250, string.Format("Bezeichnung \"{1}\" nicht in Workbook {0} gefunden und deshalb neu erstellt.", xlWorkbookFilePath, xlNamedRangeName));
                        }
                    }

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
        public static int[] XlTimeColCount(ExcelWorksheet worksheet)
        {
            Log.Write(Log.Category.MethodCall, 1907261333, string.Format("XlTimeColCount({0})",worksheet.Name));

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
        private static int XlSetRowAndCol(int[] timeCols, out int column)
        {
            Log.Write(Log.Category.MethodCall, 1907261334, string.Format("XlSetRowAndCol({0},{1}) - timeCols.Length ={2}", timeCols, "out int column", timeCols.Length));

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
        internal static void KillExcel(Process process, string filePath)
        {
            Log.Write(Log.Category.MethodCall, 1907261335, string.Format("KillExcel({0},{1})",process,filePath));

            if (process.MainWindowTitle.Contains(Path.GetFileName(filePath)))
            {
                Log.Write(Log.Category.FileSystem, -903261138, string.Format("Die Datei {0} ist bereits geöffnet und wird jetzt geschlossen.", Path.GetFileName(filePath)));
                process.Kill();
                process.WaitForExit();
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