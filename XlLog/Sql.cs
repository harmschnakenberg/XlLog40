using OfficeOpenXml;
using System;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using OfficeOpenXml.Style;
using System.Drawing;

namespace Kreutztraeger
{
    class Sql //Fehlernummern siehe Log.cs 11YYZZ
    {        
        public static string XmlDir { get; set; } = @"D:\Into_110\PROJEKTNAME\XML";
        public static string DataSource { get; set; } = @".\SQLExpress";
        static readonly string SqlConnString  = string.Format( @"Data Source={0}; Initial Catalog = WWALMDB; User ID=wwAdmin; Password=wwAdmin", DataSource);     
        //static readonly string SqlConnString = string.Format(@"ODBC; DRIVER=SQL Server; SERVER=.\sqlexpress; UID=wwuser; PWD=wwuser; DATABASE=WWALMDB", DataSource); // UNGÜLTIG!

        static string DBGruppe = "A00_General";
        static string DBGruppeComment = "unbekannt";
        static string DBStartTime;
        static DateTime StartTime = DateTime.Now.AddMonths(-1);
        static string DBEndTime;
        static DateTime EndTime = DateTime.Now;
        static int DBvonPrio = 1;
        static int DBbisPrio = 900;

        /// <summary>
        /// Erzeugt eine SQL-Abfrage der Alarmdatenbank und speichert das Ergebnis in einer Excel-Tabelle.
        /// Die Excel-Tabelle wird in ein PDF gewandelt. keepExcelFile bestimmt, ob die Excel-Datei gelöscht wird.
        /// </summary>
        /// <param name="keepExcelFile">true = erzeugt Exceldatei und PDF; false = nur PDF.</param>
        internal static void AlmListToExcel(bool keepExcelFile) //Fehlernummern siehe Log.cs 1101ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 110101, string.Format("AlmListToExcel({0})", keepExcelFile));

            try
            {

                string sqlQuery = SqlQueryString();
                if (sqlQuery.Length < 1) return;

                DataTable dt = SqlQyery(SqlConnString, sqlQuery);

                string almPrintFilePath = Path.Combine(Excel.XlArchiveDir, "Listen");
                Directory.CreateDirectory(almPrintFilePath);
                almPrintFilePath = Path.Combine(almPrintFilePath, string.Format("AlmDruck_{0:yyyy-MM-dd_HH-mm-ss}.xlsx", EndTime) );

                TryCreateNewExcelAlmFile(almPrintFilePath);

                FillAlmListFile(almPrintFilePath, dt);

                Tools.Wait(1);

                Pdf.CreatePdf(almPrintFilePath);

                if (!keepExcelFile) File.Delete(almPrintFilePath);
            }
            catch (SqlException sql_ex)
            {
                Log.Write(Log.Cat.SqlQuery, Log.Prio.Error, 110102, string.Format("Fehler in SQL-Syntax: \r\n\t\t\t{0}:\r\n\t\t\t{1}\r\n\t\t\t{2}\r\n", DataSource, SqlQueryString(), sql_ex.Message ) );
                Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.SqlQuery, Log.Prio.Error, 110103, string.Format("Fehler beim Auslesen der AlmDatenbank: Typ: {0} \r\n\t\t\t\tFehlertext: {1}  \r\n\t\t\t\tInnerException: {2}  \r\n\t\t\t\tStackTrace: {3}\r\n\r\n\r\n CONSTRING: {4}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace, SqlConnString));
                Program.AppErrorOccured = true;
            }
        }

     
        /// <summary>
        /// Erzeugt die Sql-Abfrage Zeichenfolge
        /// </summary>
        /// <returns>Sql-Abfrage Zeichenfolge</returns>
        private static string SqlQueryString() //Fehlernummern siehe Log.cs 1102ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 110201, string.Format("SqlQueryString()"));
            //ToDo: Testen; Was, wenn InTouchTagName nicht existiert?

            try
            {
                DBGruppe = (string)InTouch.ReadTag("DBGruppe");
                if (DBGruppe.Trim().Length < 1) DBGruppe = "A00_General";                
                DBGruppeComment = (string)InTouch.ReadTag(DBGruppe + ".Comment");            
                DBStartTime = (string)InTouch.ReadTag("DBStartTime");              
                DBEndTime = (string)InTouch.ReadTag("DBEndTime");
                DBvonPrio = (int)InTouch.ReadTag("DBvonPrio");
                DBbisPrio = (int)InTouch.ReadTag("DBbisPrio");

                Log.Write(Log.Cat.SqlQuery, Log.Prio.Info, 110202, string.Format("SQL-Abfrage mit: DBGruppe {0}\r\n\t\t\t\t, DBGruppeComment {1}\r\n\t\t\t\t, DBStartTime {2}\r\n\t\t\t\t, DBEndTime {3}\r\n\t\t\t\t, DBvonPrio {4}\r\n\t\t\t\t, DBbisPrio {5}\r\n\t\t\t\t", DBGruppe, DBGruppeComment, DBStartTime, DBEndTime, DBvonPrio, DBbisPrio));

                // GROUP BY - string
                string grptxt = GetGrpString();

                //Fehler, wenn GetGrpString() einen Leerstring erzeugt.
                if (grptxt.Length < 3) grptxt = "1=1";

                string usCulture = "en-US";
                bool time_ok = false;

                time_ok = DateTime.TryParse(DBStartTime, System.Globalization.CultureInfo.GetCultureInfo(usCulture), System.Globalization.DateTimeStyles.None, out StartTime);
                if (!time_ok) StartTime = DateTime.Now.AddMonths(-1);

                time_ok = DateTime.TryParse(DBEndTime, System.Globalization.CultureInfo.GetCultureInfo(usCulture), System.Globalization.DateTimeStyles.None, out EndTime);
                if (!time_ok) EndTime = DateTime.Now;

                // TEST Prüfe eingelesenen Werte
                //Log.Write(Log.Category.SqlQuery, 1903141202, "\r\nConString: " + ConnString + "\r\nDBGruppe: " + DBGruppe + "\r\nDBGruppeComment: " + DBGruppeComment + "\r\nDBStartTime: " + DBStartTime + "\t" + StartTime.ToString("g") + "\r\nDBStartTime: " + DBEndTime + "\t" + EndTime.ToString("g") + "\r\n DBvonPrio: " + DBvonPrio + "\r\n DBbisPrio: " + DBbisPrio + "\r\n");

                //# TT.MM..YYYY HH:mm
                string vonzeittxt = " AND v_AE.EventStamp > '" + StartTime.ToString("s") + "'";
                string biszeittxt = " AND v_AE.EventStamp < '" + EndTime.ToString("s") + "'";

                string SQLString = "SELECT " +
                 // "v_AE.EventStamp AS 'Datum'," +
                 "v_AE.EventStamp AS 'Zeit'," +
                 "v_AE.Priority AS 'Prio', " +
                 "v_AE.Operator AS 'Benutzer', " +
                 "v_AE.Description AS 'Beschreibung', " +
                 "v_AE.Value AS 'Wert', " +
                 "v_AE.CheckValue AS 'Alt Wert', " +
                 "v_AE.AlarmState AS 'Zustand'," +
                 "v_AE.Area AS 'Gruppe' " +
                 "FROM WWALMDB.dbo.v_AlarmEventHistory v_AE WHERE " +
                 grptxt + vonzeittxt + biszeittxt +
                 " AND (v_AE.Priority >= " + DBvonPrio +
                 " AND v_AE.Priority <= " + DBbisPrio + ") " +
                 "ORDER BY v_AE.EventStamp DESC";

                Log.Write(Log.Cat.SqlQuery, Log.Prio.Info, 110204, "SQL-Abfrage:\r\n" + SQLString);

                return SQLString;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.SqlQuery, Log.Prio.Error, 110203, string.Format("Fehler beim Zusammenstellen der AlmDatenbank-Abfrage: \r\n\t\t Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}  \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
                return string.Empty;
            }
        }

        /// <summary>
        /// Erzeugt den Teil für den SQL-Abfragestring, in dem die anzuzeigenden Alarmgruppen festgelegt werden. 
        /// </summary>
        /// <returns>für SQL-Abfrage Teilstring 'WHERE...'</returns>
        private static string GetGrpString() //Fehlernummern siehe Log.cs 1103ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 110301, string.Format("GetGrpString()"));

            try
            {
                if (!Directory.Exists(XmlDir))
                {
                    Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 110302, "Der XML-Ordner >" + XmlDir + "< konnte nicht gefunden werden.");
                    //Program.AppErrorOccured = true;
                    return null;
                }

                string xmlFilePath = Path.Combine(XmlDir, DBGruppe + ".grp");

                if (!File.Exists(xmlFilePath))
                {
                    Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 110303, string.Format("Die Alarmgruppendatei >{0}< konnte nicht gefunden werden.", xmlFilePath));
                    //Program.AppErrorOccured = true;
                    return null;
                }

                string[] queryGroups = File.ReadAllLines(xmlFilePath);

                //Alarme
                string[] almGroups = queryGroups.Where(x => x.StartsWith("A")).ToArray();
                string AGrpCmd = "(v_AE.AlarmState<>'ACK      ') AND (v_AE.Area In ('" + String.Join("', '", almGroups) + "'))";

                //Schalter
                string[] switchGroups = queryGroups.Where(x => x.StartsWith("S")).ToArray();
                string SGrpCmd = "(v_AE.AlarmState NOT Like 'ACK%') AND (v_AE.Area In ('" + String.Join("', '", switchGroups) + "'))";

                string GrpCmd = string.Empty;

                if (AGrpCmd.Count() > 0 && SGrpCmd.Count() == 0) GrpCmd = AGrpCmd;
                if (AGrpCmd.Count() > 0 && SGrpCmd.Count() > 0) GrpCmd = "(" + AGrpCmd + " OR " + SGrpCmd + ")";
                if (AGrpCmd.Count() == 0 && SGrpCmd.Count() > 0) GrpCmd = SGrpCmd;
                if (AGrpCmd.Count() == 0 && SGrpCmd.Count() == 0) GrpCmd = "(v_AE.AlarmState<>'ACK      ')";

                return GrpCmd;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.SqlQuery, Log.Prio.Error, 110304, string.Format("Fehler in GetGrpString() - Auslesen der Alarmdatenbank: \r\n\t\tTyp: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}  \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
                return null;
            }
        }

        /// <summary>
        /// Abfrage der (Alarm-)datenbank und Ausgabe als DataTable
        /// </summary>
        /// <param name="ConnString">Sql-Connection String</param>
        /// <param name="SqlQuery">Sql-Abfrage</param>
        /// <returns>DataTable mit Abfrageergebnis</returns>
        internal static DataTable SqlQyery(string ConnString, string SqlQuery) //Fehlernummern siehe Log.cs 1104ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 110401, string.Format(" SqlQyery({0},{1})", ConnString, SqlQuery));

            DataTable dataTable = new DataTable();

            using (SqlConnection con = new SqlConnection(ConnString))
            {
             
                try // Normale Benutzeranmeldung
                {
                    con.Open();
                }
                catch (System.Data.SqlClient.SqlException) // Windows-Authentifizierung
                {
                    Log.Write(Log.Cat.SqlQuery, Log.Prio.Info, 110402, "Nutze Windows Authentifizierung an SQL Server.");

                    System.Data.SqlClient.SqlConnectionStringBuilder builder = new System.Data.SqlClient.SqlConnectionStringBuilder
                    {
                        ["data source"] = System.Environment.MachineName,
                        ["database"] = "WWALMDB",
                        ["Integrated Security"] = "SSPI"
                    };

                    con.ConnectionString = builder.ConnectionString;
                    con.Open();
                }
                catch
                {
                    Log.Write(Log.Cat.SqlQuery, Log.Prio.Error, 110402, string.Format("Verbindung zu Sql Server >{0}< konnte nicht aufgebaut werden.", DataSource));
                }

                using (SqlCommand cmd = new SqlCommand(SqlQuery, con))
                {                   
                    // create data adapter
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    // this will query your database and return the result to your datatable
                    da.Fill(dataTable);
                }
            }
            return dataTable;
        }

        /// <summary>
        /// Erzeugt eine neue Excel-Mappe mit Formatierungen für eine Alarmliste.
        /// Überschreibt ggf. existierende Datei mit gleichem Pfad.
        /// </summary>
        /// <param name="xlFilePath">Pfad der zu erzeugenden Datei.</param>
        private static void TryCreateNewExcelAlmFile(string xlFilePath) //Fehlernummern siehe Log.cs 1105ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 110501, string.Format("TryCreateNewExcelAlmFile({0})", xlFilePath));

            if (File.Exists(xlFilePath)) File.Delete(xlFilePath);

            FileInfo file = new FileInfo(xlFilePath);

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    //Create new Workbook & worksheet                    
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("AlmDruck");

                    //Zeitspalte
                    worksheet.Column(1).Width = 14;
                    worksheet.Column(1).Style.Numberformat.Format = "dd.MM.yyyy hh:mm";
                    worksheet.Column(1).Style.Font.Name = "Arial";
                    worksheet.Column(1).Style.Font.Size = 9;
                    
                    //Prio
                    worksheet.Column(2).Width = 4;
                    worksheet.Column(2).Style.Font.Name = "Arial";
                    worksheet.Column(2).Style.Font.Size = 9;

                    //Benutzer
                    worksheet.Column(3).Width = 11;
                    worksheet.Column(3).Style.Font.Name = "Arial";
                    worksheet.Column(3).Style.Font.Size = 9;

                    //Beschreibung
                    worksheet.Column(4).Width = 55;
                    worksheet.Column(4).Style.Font.Name = "Arial";
                    worksheet.Column(4).Style.Font.Size = 9;

                    //Wert
                    worksheet.Column(5).Width = 4;
                    worksheet.Column(5).Style.Font.Name = "Arial";
                    worksheet.Column(5).Style.Font.Size = 9;

                    //Altwert
                    worksheet.Column(6).Width = 7;
                    worksheet.Column(6).Style.Font.Name = "Arial";
                    worksheet.Column(6).Style.Font.Size = 9;

                    ////Zustand
                    //worksheet.Column(7).Width = 11;
                    //worksheet.Column(7).Style.Font.Name = "Arial";
                    //worksheet.Column(7).Style.Font.Size = 9;

                    //Gruppe
                    worksheet.Column(8).Width = 13;
                    worksheet.Column(8).Style.Font.Name = "Arial";
                    worksheet.Column(8).Style.Font.Size = 9;

                    //Header
                    worksheet.HeaderFooter.FirstHeader.CenteredText = "&16&\"Arial,Regular Bold\" Alarmdatenbank";
                    worksheet.PrinterSettings.RepeatRows = worksheet.Cells["1:1"];

                    //Seitenlayout
                    worksheet.PrinterSettings.PaperSize = ePaperSize.A4;
                    worksheet.PrinterSettings.Orientation = eOrientation.Landscape;
                    //worksheet.PrinterSettings.LeftMargin = (decimal) 0.1;
                    //worksheet.PrinterSettings.RightMargin = (decimal)0.1;
                    worksheet.PrinterSettings.HorizontalCentered = true;
                    worksheet.PrinterSettings.FitToPage = true;
                    worksheet.PrinterSettings.FitToWidth = 1;
                    worksheet.PrinterSettings.FitToHeight = 0;
                    worksheet.View.PageLayoutView = true;
                    excelPackage.Workbook.Calculate();
                    excelPackage.Save();
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 110502, string.Format("Datei für Alarmausdruck konnte nicht erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}  \r\n\t\t StackTrace: {4}", xlFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
            }
        }

        /// <summary>
        /// Füllt die VorlageDatei xlFilePath mit der DataTable aus der SQL-Abfrage
        /// </summary>
        /// <param name="xlFilePath"></param>
        /// <param name="dt">DataTable aus SQL-Abfrage SqlQuery()</param>
        private static void FillAlmListFile(string xlFilePath, DataTable dt) // //Fehlernummern siehe Log.cs 1106ZZ
        {
            if (!File.Exists(xlFilePath))
            {
                Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 110602, string.Format("Die Datei {0} für AlmDruck existiert nicht.", xlFilePath));
                return;
            }

            FileInfo file = new FileInfo(xlFilePath);

            try
            {
                using (ExcelPackage excelPackage = new ExcelPackage(file))
                {
                    //Kommentare
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets["AlmDruck"];

                    worksheet.Cells[1, 1].AddComment("von " + StartTime.ToString("g"), (string)InTouch.ReadTag("$Operator"));
                    worksheet.Cells[1, 2].AddComment(" bis " + EndTime.ToString("g"), (string)InTouch.ReadTag("$Operator"));
                    worksheet.Cells[1, 8].AddComment(DBGruppe, (string)InTouch.ReadTag("$Operator"));
                    //worksheet.Cells[1, 1].AddComment(DBGruppe + " (" + DBGruppeComment + ") von " + StartTime.ToString("g") + " bis " + EndTime.ToString("g"), (string)InTouch.ReadTag("$Operator"));

                    //Überschriften
                    worksheet.Cells[1, 1].Value = dt.Columns[0].Caption; //Zeit
                    worksheet.Cells[1, 2].Value = dt.Columns[1].Caption; //Prio
                    worksheet.Cells[1, 3].Value = dt.Columns[2].Caption; //Benutzer
                    worksheet.Cells[1, 4].Value = dt.Columns[3].Caption; //Beschreibung
                    worksheet.Cells[1, 5].Value = dt.Columns[4].Caption; //Wert
                    worksheet.Cells[1, 6].Value = dt.Columns[5].Caption; //Altwert
                    //worksheet.Cells[1, 7].Value = dt.Columns[6].Caption; //Zustand
                    worksheet.Cells[1, 8].Value = dt.Columns[7].Caption; //Gruppe
                    worksheet.Cells[1, 1, 1, 8].Style.Font.Bold = true;

                    //Werte
                    int row = 1;
                    Color almColor = Color.Black;

                    foreach (DataRow dataRow in dt.Rows)
                    {
                        row++;
                        object[] array = dataRow.ItemArray;

                        string zustand = array[dt.Columns["Zustand"].Ordinal].ToString();
                        string gruppe = array[dt.Columns["Gruppe"].Ordinal].ToString();

                        almColor = SetFontColor(zustand, gruppe[0]);

                        for (int col = 0; col < array.Length; col++)
                        {
                            worksheet.Cells[row, col + 1].Value = array[col];
                            worksheet.Cells[row, col + 1].Style.Font.Color.SetColor(almColor);
                            worksheet.Cells.Style.Border.Bottom.Style = ExcelBorderStyle.Hair;
                        }
                    }
                    
                    string description = "&8&\"Arial\" Obergruppe: " + DBGruppeComment + " von " + StartTime.ToString("g") + " bis " + EndTime.ToString("g") + ", Priorität " + DBvonPrio.ToString() + " bis " + DBbisPrio.ToString();
                    worksheet.HeaderFooter.FirstFooter.LeftAlignedText = description;
                    worksheet.HeaderFooter.EvenFooter.LeftAlignedText = description;
                    worksheet.HeaderFooter.OddFooter.LeftAlignedText = description;

                    string pageNo = string.Format("{0} / {1}", ExcelHeaderFooter.PageNumber, ExcelHeaderFooter.NumberOfPages);
                    worksheet.HeaderFooter.FirstFooter.RightAlignedText = pageNo;
                    worksheet.HeaderFooter.EvenFooter.RightAlignedText = pageNo;
                    worksheet.HeaderFooter.OddFooter.RightAlignedText = pageNo;
                    
                    excelPackage.Workbook.Calculate();
                    excelPackage.Save();
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.ExcelWrite, Log.Prio.Error, 110503, string.Format("In die Datei für Alarmausdruck konnte nicht geschrieben werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}  \r\n\t\t StackTrace: {4}", xlFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));

            }
        }

        /// <summary>
        /// Textfarbe in Abhänigkeit von Alarmstatus und Alarm/Schalter
        /// </summary>
        /// <param name="status">Spalte 'Zustand' aus SQL-Abfrage</param>
        /// <param name="group">Erstes Zeichen aus Spalte Gruppe aus SQL-Abfrage</param>
        /// <returns>Textfarbe</returns>
        private static Color SetFontColor(string status, char group) //Fehlernummern siehe Log.cs 1107ZZ
        {
            /*
                [Zustand] mögliche Werte im Feld
                "UNACK_ALM", Alarm gekommen / Schalter nicht Auto
                "         ", Ereignisse / Sollwerte
                "ACK_ALM  ", Quittierung aktiver Alarm
                "ACK_RTN  ", Alarm gegangen / Schalter Auto

                Alarme gekommen:     rot
                Alarme gegangen:     grün
                Alarme quittiert:    blau
                Schalter nicht Auto: hell rot
                Schalter Auto:       hell grün
                Ereignisse:          schwarz
                Sollwerte:           schwarz
            */


            if (status.Equals("UNACK_ALM") && group.Equals('A'))
            {
                return Color.Red;
            }
            else if (status.Equals("ACK_RTN  ") && group.Equals('A'))
            {
                return Color.Green;
            }
            else if (status.Equals("ACK_ALM  ") && group.Equals('A'))
            {
                return Color.Blue;
            }
            else if (status.Equals("UNACK_RTN") && group.Equals('S'))
            {
                return Color.LightGreen;
            }
            else if (status.Equals("UNACK_ALM") && group.Equals('S'))
            {
                return Color.Orange;
            }
            else
            {
                return Color.Black;
            }
        }

      
    }


}

