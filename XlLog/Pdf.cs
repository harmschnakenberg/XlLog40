using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
//using System.Runtime.InteropServices;

namespace Kreutztraeger
{
    internal static class Pdf
    {

        /// <summary>
        /// Vor dieser Stunde soll nicht in PDF umgewandelt werden. Wert < 0 || Wert > 24 = nie 
        /// </summary>
        public static int PdfConvertStartHour = 0;
        public static string PdfConverterPath { get; set; } = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "Xlsx2Pdf.exe"); // @"D:\XlLog\LibreOfficePortable\LibreOfficeCalcPortable.exe";
        public static string PdfConverterArgs { get; set; } = "*Quelle* *Ziel*"; //"-calc -invisible -convert-to pdf \"*Quelle*\" -outdir \"*Ziel*\"";

        /// <summary>
        /// Prüft Queldatei und Converter für xlsx -> pdf. Startet die Umwandlung von Excel-Datei in PDF.
        /// </summary>
        /// <param name="xlFilePath">Pfad zur Excel-Datei, die in PDF umgewnadelt werden soll.</param>
        /// <param name="MinHour"> Werte >= 0 geben die früheste Tagesstunde an, in der umgewandelt wird (Zeitsteuerung PDF-Druck). Werte kleiner 0 => sofort umwandeln.</param>
        public static void CreatePdf(string xlFilePath)
        {
            Log.Write(Log.Category.MethodCall, 1907261350, string.Format("CreatePdf({0})", xlFilePath));

            try
            {
                if (!File.Exists(xlFilePath))
                {
                    Log.Write(Log.Category.FileSystem, 1903111805, string.Format("Es kann kein PDF erzeugt werden. Die Datei {0} konnte nicht gefunden werden.", xlFilePath));
                    //kein Fehler
                    return;
                }

                if (Path.GetFileName(xlFilePath).StartsWith("~") )
                {
                    Log.Write(Log.Category.FileSystem, 2002191605, string.Format("Temporäre Datei {0} wird übersprungen.", xlFilePath));
                    //kein Fehler
                    return;
                }

                if (!File.Exists(Pdf.PdfConverterPath))
                {
                    Log.Write(Log.Category.FileSystem, -904241035, string.Format("Der PDF-Converter {0} konnte nicht gefunden werden.", PdfConverterPath));
                    Program.AppErrorOccured = true;
                }
                else
                {
                    //Wenn xlFilePath keine temporäre Datei ist und Zeit mindestens MinHour  
                    if (!Path.GetFileNameWithoutExtension(xlFilePath).StartsWith("~") )
                    {
                        CreatePdFWithConverter(xlFilePath);
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.PdfWrite, -902011251, string.Format("Fehler beim erstellen der PDF-Datei : Typ: \r\n{0} \r\n\t\t Fehlertext: \r\n{1}  \r\n\t\t InnerException: \r\n{2}\r\n\t\t StackTrace: \r\n{3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace));
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// Erzeugt eine PDF aus der angegebenen Excel-Datei. Wird aufgerufen von 'public static void CreatePdf(string xlFilePath)' 
        /// </summary>
        /// <param name="xlFilePath">Pfad zur Excel-Datei, die in PDF umgewnadelt werden soll.</param>
        /// <returns>Gibt 'true' zurück, wenn die Datei ohne Fehler erzeugt wurde, ansonsten 'false'.</returns>
        private static bool CreatePdFWithConverter(string xlFilePath)
        {
            Log.Write(Log.Category.MethodCall, 1907261351, string.Format("CreatePdFWithConverter({0})", xlFilePath));

            string LO_CalcCommand = PdfConverterArgs.Replace("*Quelle*", xlFilePath).Replace("*Ziel*", Path.GetDirectoryName(xlFilePath)); // "-calc -invisible -convert-to pdf \"" + xlFilePath + "\" -outdir \"" + Path.GetDirectoryName(xlFilePath) + "\"";

            Log.Write(Log.Category.PdfWrite, 1907241709, "PDF-erzeugen mit: " + PdfConverterPath + " " + LO_CalcCommand);

            ProcessStartInfo start = new ProcessStartInfo
            {
                FileName = Pdf.PdfConverterPath, // Specify exe name.
                UseShellExecute = false,
                CreateNoWindow = false,
                WindowStyle = ProcessWindowStyle.Normal,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                Arguments = LO_CalcCommand
            };

            using (Process process = Process.Start(start))
            {
                // Read the error stream first and then wait.
                string error = process.StandardError.ReadToEnd();
                // Bei SFB Holdorf wurde ohne Timeout bis zu 6 Stunden gewartet! Timeout 60 sec.!
                process.WaitForExit(60000);

                //Check for an error
                if (!string.IsNullOrEmpty(error))
                {
                    Log.Write(Log.Category.PdfWrite, -902250800, string.Format("Die Anwendung {0} konnte {1} nicht in PDF umwandeln.\r\n\t\t\t\tFehlertext: {2}", Pdf.PdfConverterPath, xlFilePath, error));
                    Program.AppErrorOccured = true;
                    return false;
                }
                else
                {
                    Log.Write(Log.Category.Info, 1902250802, string.Format("Die PDF-Datei {0} wurde erzeugt.", Path.ChangeExtension(xlFilePath, ".pdf")));                    
                    return true;
                }                
            }
        }

        /// <summary>
        /// Prüft, ob in DirPath für alle *.xlsx-Dateien eine gleichnamige .pdf vorhanden ist und erstellt ggf. fehlende PDFs außer vom aktuellen Tag.
        /// </summary>
        /// <param name="DirPath">Ordner in dem *xlsx-Dateien ohne dazugehörige PDFs liegen können.</param>
        internal static void CreatePdf4AllXlsxInDir(string DirPath)
        {
            Log.Write(Log.Category.MethodCall, 1907261352, string.Format("CreatePdf4AllXlsxInDir({0})", DirPath));
            Console.WriteLine("Bitte warten. PDF-Dateien werden erstellt...");

            foreach (string filePath in Directory.GetFiles(DirPath, "*.xlsx"))
            {
                //Erzeuge kein PDF vom aktuellen Tag
                if (filePath == Excel.CeateXlFilePath()) continue;

                string pdfPath = Path.ChangeExtension(filePath, ".pdf");
                if (!File.Exists(pdfPath) && !Path.GetFileName(filePath).StartsWith("~") )
                {
                    Log.Write(Log.Category.PdfWrite, 1903250919, "Erzeuge Datei " + pdfPath);
                    CreatePdf(filePath);
                }
            }
        }

        /// <summary>
        /// direkte PDF-Erzeugen aus InTouch z.B. aus PDF-Betrachter.
        /// </summary>
        public static void CreatePdfFromCmd()
        {
            Log.Write(Log.Category.MethodCall, 1907261353, string.Format("CreatePdfFromCmd()"));

            if (Program.CmdArgs.Length > 1)
            {
                string xlFilePath = Program.CmdArgs[1];
                if (File.Exists(xlFilePath))
                {
                    Log.Write(Log.Category.PdfWrite, 1903181406, "Erzeuge PDF-Datei aus " + xlFilePath);
                    Pdf.CreatePdf(xlFilePath);
                }
                else
                {
                    Log.Write(Log.Category.FileSystem, -903181407, "Es konnte keine PDF-Datei erstellt werden. Ungültige Quelldatei: " + xlFilePath);
                    Program.AppErrorOccured = true;
                }
            }
            else
            {
                Log.Write(Log.Category.PdfWrite, -903181408, "Es konnte keine PDF-Datei erstellt werden. Fehlendes 2. Argument für Quellpfad (xslx)");
                Program.AppErrorOccured = true;
            }
        }

    }
}
