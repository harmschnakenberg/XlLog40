using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;

namespace Kreutztraeger
{
    class Config
    {
        private const string ConfigFileName = "XlConfig.ini";

        /// <summary>
        /// Erstellt eine Konfig-INI mit Default-Werten.
        /// </summary>
        /// <param name="ConfigFileName">Name der Konfig-Datei</param>
        private static void CreateConfig(string ConfigFileName)
        {
            Log.Write(Log.Category.MethodCall, 1907261400, string.Format("CreateConfig({0})", ConfigFileName));

            string configPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), ConfigFileName);
            using (StreamWriter w = File.AppendText(configPath))
            {
                try
                {
                    w.WriteLine("[" + w.Encoding.ToString() + "]\r\n" +
                                "[Allgemein]\r\n" +
                                ";DebugWord=" + Log.DebugWord + "\r\n" + 
                                ";WaitToClose=" + Tools.WaitToClose + "\r\n" + 
                                ";WaitForScripts=" + Tools.WaitForScripts + "\r\n" +

                                "\r\n[InTouch]\r\n" +
                                ";InTouchDiscFlag=" + Program.XlLogFlag + "\r\n" + 
                                ";InTouchDiscAlarm="+ Program.InTouchDiscAlarm + "\r\n" +
                                ";InTouchDiscTimeOut=" + Program.InTouchDiscTimeOut + "\r\n" + 

                                "\r\n[Pfade]\r\n" +
                                ";XlArchiveDir=" + Excel.XlArchiveDir + "\r\n" +
                                ";XmlDir=" + Sql.XmlDir + "\r\n" + 

                                "\r\n[Vorlagen]\r\n" +
                                ";XlTemplateDayFilePath=" + Excel.XlTemplateDayFilePath + "\r\n" +
                                ";XlTemplateMonthFilePath=" + Excel.XlTemplateMonthFilePath + "\r\n" +
                                ";XlPassword=" + Excel.XlPassword + "\r\n" +
                                ";XlDayFileFirstRowToWrite=" + Excel.XlDayFileFirstRowToWrite + "\r\n" +
                                ";XlMonthFileFirstRowToWrite=" + Excel.XlMonthFileFirstRowToWrite + "\r\n" +
                                ";XlPosOffsetMin=" + Excel.XlPosOffsetMin + "\r\n" +
                                ";XlNegOffsetMin=" + Excel.XlNegOffsetMin + "\r\n" +

                                "\r\n[PDF]\r\n" +
                                ";PdfConvertStartHour=" + Pdf.PdfConvertStartHour + "\r\n" +
                                ";PdfConverterPath=" + Pdf.PdfConverterPath + "\r\n" +
                                ";PdfConverterArgs=" + Pdf.PdfConverterArgs + "\r\n\r\n" +
                                
                                ";PdfConverterPath=D:\\XlLog\\LibreOfficePortable\\LibreOfficeCalcPortable.exe\r\n" +
                                ";PdfConverterArgs=- calc - invisible - convert - to pdf \"*Quelle*\" -outdir \"*Ziel*\"\r\n" +
                                ";PdfConverterPath=D:\\XlLog\\XlOffice2Pdf.exe\r\n" +
                                ";PdfConverterArgs=*Quelle* *Ziel*\r\n" +
                                ";PdfConverterPath=D:\\XlLog\\Xl2Pdf.exe\r\n" +
                                ";PdfConverterArgs=*Quelle* *Ziel*\r\n" +

                                "\r\n[Druck]\r\n" +
                                ";PrintStartHour=" + Print.PrintStartHour + "\r\n" +
                                ";PrintAppPath=" + Print.PrintAppPath + "\r\n" +
                                ";PrintAppArgs=" + Print.PrinterAppArgs + "\r\n\r\n" +

                                ";PrintAppPath=D:\\XlLog\\XlOfficePrint.exe\r\n" + 
                                ";PrintAppPath=D:\\XlLog\\PdfToPrinter.exe\r\n" +
                                ";PrintAppArgs=\"*Quelle*\" \"HP OfficeJet Pro 8210\" pages=*Seiten*\r\n" 
                                
                                );
                }
                catch (Exception ex)
                {
                    Log.Write(Log.Category.FileSystem, -902060750, string.Format("Die Konfigurationsdatei konnte nicht gefunden oder erstellt werden: {0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", configPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                    Console.WriteLine("FEHLER beim Erstellen von {0}. Siehe Log.", configPath);
                    Program.AppErrorOccured = true;
                }
            }
        }

        /// <summary>
        /// Lädt Werte aus der Konfig-INI.
        /// </summary>
        internal static void LoadConfig()
        {
            //Console.WriteLine("LoadConfig() gestartet.");
            Log.Write(Log.Category.MethodCall, 1911281128, string.Format("LoadConfig()"));

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
                            if (item.Length > 2) val += "=" + item[2].Trim();
                            dict.Add(item[0].Trim(), val);
                        }
                    }

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

                    //String
                    configVal = TagValueFromConfig(dict, "InTouchDiscFlag");
                    if (configVal != null)
                        Program.XlLogFlag = dict["InTouchDiscFlag"];

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
                    
                    configVal = TagValueFromConfig(dict, "PrintStartHour");
                    if (int.TryParse(configVal, out i))
                        Print.PrintStartHour = i;

                    configVal = TagValueFromConfig(dict, "XlPassword");
                    if (int.TryParse(configVal, out i))
                        Print.PrintStartHour = i;
                    
                }

            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.FileSystem, -902060750, string.Format("Fehler beim Lesen der Konfigurationsdatei: \r\n\t\t{0}\r\n\t\t Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", configPath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                Console.WriteLine("FEHLER beim Lesen von {0}. Siehe Log.", configPath);
                Program.AppErrorOccured = true;
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

    }
}
