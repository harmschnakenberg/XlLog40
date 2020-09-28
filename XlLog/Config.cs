using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;


namespace Kreutztraeger
{
    class Config //Fehlernummern siehe Log.cs 02YYZZ
    {
        private const string ConfigFileName = "XlConfig.ini";

        /// <summary>
        /// Erstellt eine Konfig-INI mit Default-Werten.
        /// </summary>
        /// <param name="ConfigFileName">Name der Konfig-Datei</param>
        private static void CreateConfig(string ConfigFileName) //Fehlernummern siehe Log.cs 0201ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.LogAlways, 020101, string.Format("CreateConfig({0})", ConfigFileName));
                                                    
            string configPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ConfigFileName);
            using (StreamWriter w = File.AppendText(configPath))
            {
                try
                {
                    w.WriteLine("[öäü " + w.Encoding.EncodingName + ", Build " + System.Reflection.Assembly.GetExecutingAssembly().GetName().Version +"]\r\n" +
                                "\r\n[Intern]\r\n" +
                                ";DebugWord=" + Log.DebugWord + "\r\n" +
                                ";WaitToClose=" + Tools.WaitToClose + "\r\n" +
                                ";WaitForScripts=" + Tools.WaitForScripts + "\r\n" +
                                ";StartTaskIntervallMinutes=" + Scheduler.StartTaskIntervallMinutes + "\r\n" +
                                // ";DataSource=" + Sql.DataSource + "\r\n" +
                                ";;DataSource=" + Environment.MachineName + "\r\n" +
                                
                                "\r\n[InTouch]\r\n" +
                                ";InTouchDiscFlag=" + Program.InTouchDiscXlLogFlag + "\r\n" +
                                ";InTouchDiscAlarm=" + Program.InTouchDiscAlarm + "\r\n" +
                                ";InTouchDiscTimeOut=" + Program.InTouchDiscTimeOut + "\r\n" +
                                ";InTouchDiscSetCalculations=" + Program.InTouchDiscSetCalculations + "\r\n" +
                                ";InTouchDiscResetHourCounter=" + Program.InTouchDiscResetHourCounter + "\r\n" +
                                ";InTouchDiscResetQuarterHourCounter=" + Program.InTouchDiscResetQuarterHourCounter + "\r\n" +
                                ";InTouchDIntErrorNumber=" + Program.InTouchDIntErrorNumber + "\r\n" +
                                ";InTouchIntPrintBitMaskDay=" + Print.InTouchIntPrintBitMaskDay + "\r\n" +
                                ";InTouchIntPrintBitMaskMonth=" + Print.InTouchIntPrintBitMaskMonth + "\r\n" +

                                "\r\n[Pfade]\r\n" +
                                ";XlArchiveDir=" + Excel.XlArchiveDir + "\r\n" +
                                ";;XmlDir=" + TryFindXmlDir() + "\r\n" +

                                "\r\n[Vorlagen]\r\n" +
                                ";XlTemplateDayFilePath=" + Excel.XlTemplateDayFilePath + "\r\n" +
                                ";XlTemplateMonthFilePath=" + Excel.XlTemplateMonthFilePath + "\r\n" +
                                ";XlPassword=" + Excel.XlPassword + "\r\n" +
                                ";XlDayFileFirstRowToWrite=" + Excel.XlDayFileFirstRowToWrite + "\r\n" +
                                ";XlMonthFileFirstRowToWrite=" + Excel.XlMonthFileFirstRowToWrite + "\r\n" +
                                ";XlPosOffsetMin=" + Excel.XlPosOffsetMin + "\r\n" +
                                ";XlNegOffsetMin=" + Excel.XlNegOffsetMin + "\r\n" +

                                "\r\n[PDF]\r\n" +
                                ";XlImmediatelyCreatePdf=0\r\n" +
                                ";PdfConvertStartHour=" + Pdf.PdfConvertStartHour + "\r\n" +
                                ";PdfConverterPath=" + Pdf.PdfConverterPath + "\r\n" +
                                ";;PdfConverterPath=D:\\XlLog\\XlOffice2Pdf.exe\r\n" +
                                ";PdfConverterArgs=" + Pdf.PdfConverterArgs + "\r\n" +
                                //";PdfConverterArgs=*Quelle* *Ziel*\r\n" +

                                "\r\n[Druck]\r\n" +
                                ";PrintStartHour=" + Print.PrintStartHour + "\r\n" +
                                ";PrintAppPath=" + Print.PrintAppPath + "\r\n" +
                                ";;PrintAppPath=D:\\XlLog\\XlOfficePrint.exe\r\n" +
                                ";PrintAppArgs=" + Print.PrinterAppArgs + "\r\n" +
                                ";;PrintAppArgs=\"*Quelle*\" \"HP OfficeJet Pro 8210\" pages=*Seiten*\r\n"
                                ); ;
                }
                catch (Exception ex)
                {
                    Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 020102, string.Format("Die Konfigurationsdatei konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", configPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                    Console.WriteLine("FEHLER beim Erstellen von {0}. Siehe Log.", configPath);
                    //Program.AppErrorOccured = true;
                }
            }
        }

        /// <summary>
        /// Lädt Werte aus der Konfig-INI.
        /// </summary>
        internal static void LoadConfig()
        {
            //Console.WriteLine("LoadConfig() gestartet.");
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 020103, string.Format("LoadConfig()"));

            string appDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string configPath = Path.Combine(appDir, ConfigFileName);

            try
            {

                if (!File.Exists(configPath))
                {
                    CreateConfig(ConfigFileName);
                    Console.WriteLine("Neue Config.ini angelegt unter " + configPath);
                }
                else
                { 
                    string configAll = System.IO.File.ReadAllText(configPath, System.Text.Encoding.UTF8); 
                    char[] delimiters = new char[] { '\r', '\n' };
                    string[] configLines = configAll.Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                    Dictionary<string, string> dict = new Dictionary<string, string>();
                    foreach (string line in configLines)
                    {                        
                        if (line[0] != ';' && line[0] != '[')
                        {
                            string[] item = line.Split('=');
                            string val = item[1].Trim();
                            if (item.Length > 2)
                            {
                                for (int n = 2; n < item.Length; n++)
                                {
                                    val += "=" + item[n].Trim();
                                }
                            }
                            dict.Add(item[0].Trim(), val);
                        }
                    }

                    if (dict.Count == 0) return;

                    //Dateipfade
                    string configVal = TagValueFromConfig(dict, "XlTemplateDayFilePath");
                    if (File.Exists(configVal))
                        Excel.XlTemplateDayFilePath = configVal;

                    configVal = TagValueFromConfig(dict, "XlTemplateMonthFilePath");
                    if (File.Exists(configVal))
                        Excel.XlTemplateMonthFilePath = configVal;

                    //Ordnerpfade
                    configVal = TagValueFromConfig(dict, "XlArchiveDir");
                    if (Directory.Exists(configVal))
                        Excel.XlArchiveDir = configVal;

                    configVal = TagValueFromConfig(dict, "XmlDir");
                    if (Directory.Exists(configVal))
                        Sql.XmlDir = configVal;

                    configVal = TagValueFromConfig(dict, "PdfConverterPath");
                    if (File.Exists(configVal))
                        Pdf.PdfConverterPath = configVal;

                    configVal = TagValueFromConfig(dict, "PrintAppPath");
                    if (File.Exists(configVal))
                        Print.PrintAppPath = configVal;

                    //Integer
                    configVal = TagValueFromConfig(dict, "XlDayFileFirstRowToWrite");
                    if (int.TryParse(configVal, out int i))
                        Excel.XlDayFileFirstRowToWrite = i;

                    configVal = TagValueFromConfig(dict, "DebugWord");
                    if (int.TryParse(configVal, out i))
                        Log.DebugWord = i;

                    configVal = TagValueFromConfig(dict, "XlPosOffsetMin");
                    if (int.TryParse(configVal, out i))
                        Excel.XlPosOffsetMin = i;

                    configVal = TagValueFromConfig(dict, "XlNegOffsetMin");
                    if (int.TryParse(configVal, out i))
                        Excel.XlNegOffsetMin = i;

                    configVal = TagValueFromConfig(dict, "PdfConvertStartHour");
                    if (int.TryParse(configVal, out i))
                        Pdf.PdfConvertStartHour = i;

                    configVal = TagValueFromConfig(dict, "WaitToClose");
                    if (int.TryParse(configVal, out i))
                        Tools.WaitToClose = i;

                    configVal = TagValueFromConfig(dict, "WaitForScripts");
                    if (int.TryParse(configVal, out i))
                        Tools.WaitForScripts = i;

                    configVal = TagValueFromConfig(dict, "PrintStartHour");
                    if (int.TryParse(configVal, out i))
                        Print.PrintStartHour = i;

                    configVal = TagValueFromConfig(dict, "XlImmediatelyCreatePdf");
                    if (int.TryParse(configVal, out i))
                    {
                        if (i > 0) Excel.XlImmediatelyCreatePdf = true;
                        else Excel.XlImmediatelyCreatePdf = false;
                    }

                    configVal = TagValueFromConfig(dict, "StartTaskIntervallMinutes");
                    if (int.TryParse(configVal, out i))
                        Scheduler.StartTaskIntervallMinutes = i;

                    //String
                    configVal = TagValueFromConfig(dict, "InTouchDiscFlag");
                    if (configVal != null)
                        Program.InTouchDiscXlLogFlag = dict["InTouchDiscFlag"];

                    configVal = TagValueFromConfig(dict, "InTouchDiscAlarm");
                    if (configVal != null)
                        Program.InTouchDiscAlarm = dict["InTouchDiscAlarm"];

                    configVal = TagValueFromConfig(dict, "InTouchDiscTimeOut");
                    if (configVal != null)
                        Program.InTouchDiscTimeOut = dict["InTouchDiscTimeOut"];

                    configVal = TagValueFromConfig(dict, "PdfConverterArgs");
                    if (configVal != null)
                        Pdf.PdfConverterArgs = configVal;

                    configVal = TagValueFromConfig(dict, "PrintAppArgs");
                    if (configVal != null)
                        Print.PrinterAppArgs = configVal;
                    
                    configVal = TagValueFromConfig(dict, "XlPassword");
                    if (configVal != null)
                    {
                        if (configVal.StartsWith("\"") && configVal.EndsWith("\""))
                        {
                            string encrypt = configVal.Substring(1, configVal.LastIndexOf("\"") - 1);
                            Excel.XlPasswordEncrypted = encrypt;
                            Excel.XlPassword = EncryptDecrypt(encrypt, 200);
                        }
                        else
                        {
                            Excel.XlPassword = configVal;
                        }
                    }

                    configVal = TagValueFromConfig(dict, "InTouchDIntErrorNumber");
                    if (configVal != null)                    
                        Program.InTouchDIntErrorNumber = configVal;                    

                    configVal = TagValueFromConfig(dict, "InTouchIntPrintBitMaskDay");
                    if (configVal != null)                    
                        Print.InTouchIntPrintBitMaskDay = configVal;
                    
                    configVal = TagValueFromConfig(dict, "InTouchIntPrintBitMaskMonth");
                    if (configVal != null)                    
                        Print.InTouchIntPrintBitMaskMonth = configVal;

                    configVal = TagValueFromConfig(dict, "DataSource");
                    if (configVal != null)
                        Sql.DataSource = configVal;
                    
                }

            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 020104, string.Format("Fehler beim Lesen der Konfigurationsdatei: \r\n\t\t{0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", configPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                Console.WriteLine("FEHLER beim Lesen von {0}. Siehe Log.", configPath);
                //Program.AppErrorOccured = true;
            }
        }

        private static string TagValueFromConfig(Dictionary<string, string> dict, string TagName)
        {
            if (dict.TryGetValue(TagName, out string val))
            {
                return val;
            }
            else return null;
        }

        /// <summary>
        /// Passwortentschlüsselung
        /// </summary>
        /// <param name="szPlainText"></param>
        /// <param name="szEncryptionKey"></param>
        /// <returns></returns>
        private static string EncryptDecrypt(string szPlainText, int szEncryptionKey)
        {
            StringBuilder szInputStringBuild = new StringBuilder(szPlainText);
            StringBuilder szOutStringBuild = new StringBuilder(szPlainText.Length);
            char Textch;
            for (int iCount = 0; iCount < szPlainText.Length; iCount++)
            {
                Textch = szInputStringBuild[iCount];
                Textch = (char)(Textch ^ szEncryptionKey);
                szOutStringBuild.Append(Textch);
            }
            return szOutStringBuild.ToString();
        }


        /// <summary>
        /// Sucht den Ordner XML im InTouch-Projekt unter
        /// rootDir\Into*\*\XML
        /// </summary>
        /// <param name="rootDir">Stammverzeichnis bzw. Laufwerksbezeichnung wenn nicht im gleichen Laufwerk wie XlLog.exe</param>
        private static string TryFindXmlDir(string rootDir = null)
        {
            if (rootDir == null) rootDir = Directory.GetDirectoryRoot(Assembly.GetExecutingAssembly().Location);

            try
            {
                string[] IntoDirs = Directory.GetDirectories(rootDir, @"Into*"); // Direktes suchen im root-Verzeichnis gibt Fehlermeldung      
                DirectoryInfo[] ProjectDirs = new DirectoryInfo(IntoDirs[0]).GetDirectories().OrderByDescending(o => o.LastWriteTime).ToArray(); // Jüngste Ordner = Log und aktueller Projektordner?

                string xmlDir = Sql.XmlDir;

                foreach (DirectoryInfo dir in ProjectDirs)
                {
                    if (dir.Name != "Log")
                    {
                        xmlDir = dir.GetDirectories("XML")[0].FullName;
                        break;
                    }
                }

                Sql.XmlDir = xmlDir;
                //string xmlDir = Directory.GetDirectories(IntoDirs[0], @"XML", SearchOption.AllDirectories)[0];
                return xmlDir;
            }
            catch
            {
                return Sql.XmlDir;
            }                   
        }
    }
}
