using System;
using System.Diagnostics;
using System.IO;
using System.Threading;

namespace Kreutztraeger
{
    class Scheduler
    {

        static readonly int intervallMinutes = 15;
        static readonly string taskName = string.Format(@"XlLog_vbs{0}", intervallMinutes);

        /// <summary>
        /// Prüft, ob eine *.vbs-Datei und ein SchedulerTask für die automatische Ausführung vorhanden sind und erstellt diese ggf. 
        /// </summary>
        internal static void CeckOrCreateTaskScheduler()
        {
            try
            {
                string taskPath = System.Reflection.Assembly.GetAssembly(typeof(Program)).Location;  // Diese exe

                //Erstelle *.vbs-Dateien für versteckte Ausführung
                string vbsFilePath = Path.Combine(Path.GetDirectoryName(taskPath), "XlLogSchock.vbs");
                string fileContent = string.Format("Set WshShell = WScript.CreateObject(\"WScript.Shell\")\r\nWshShell.Run \"{0} -Schock\",0,True", taskPath);
                CreateVbsScript(taskPath, vbsFilePath, fileContent);

                vbsFilePath = Path.Combine(Path.GetDirectoryName(taskPath), "XlLogAlmDruck.vbs");
                fileContent = string.Format("Set WshShell = WScript.CreateObject(\"WScript.Shell\")\r\nWshShell.Run \"{0} -AlmDruck\",0,True", taskPath);
                CreateVbsScript(taskPath, vbsFilePath, fileContent);

                vbsFilePath = Path.ChangeExtension(taskPath, "vbs");
                fileContent = string.Format("Set WshShell = WScript.CreateObject(\"WScript.Shell\")\r\nWshShell.Run \"{0} -Task\",0,True", taskPath);
                CreateVbsScript(taskPath, vbsFilePath, fileContent);

                if (!CheckForSchedulerTask(taskName))
                {
                    if (!CreateSchedulerTask(intervallMinutes, taskName, vbsFilePath))
                    {
                        Log.Write(Log.Category.Scheduler, -902011218, "Es konnte kein neuer Task erstellt werden.");
                        Program.AppErrorOccured = true;
                    }
                }
                else
                {
                    // Log.Write(Log.DebugCategory.WriteToConsole, 1902011219, "Task ist im TaskScheduler vorhanden.");
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.Scheduler, -902011217, string.Format("Fehler beim bearbeiten des TaskSchedulers: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// VBS-Datei für versteckte Ausführung der exe
        /// </summary>
        /// <param name="exeFilePath"></param>
        private static void CreateVbsScript(string exeFilePath, string vbsFilePath, string fileContent)
        {
            // Um die Anwendung unter dem zurzeit angemeldeten Benutzer versteckt zu starten, wird der Umweg über eine *.vbs-Datei gewählt.
            // Bei Starten im TaskScheduler mit fester Benutzeranmeldung (= Versteckt) werden InTouch-Tags nicht mehr erreicht.
            // Schockkühler soll direkt aus InTouch per vbs gestartet werden mit Parameter -Schock
            if (!File.Exists(vbsFilePath))
            {
                using (StreamWriter w = File.AppendText(vbsFilePath))
                {
                    try
                    {
                        w.WriteLine(fileContent);
                    }
                    catch (Exception ex)
                    {
                        Log.Write(Log.Category.FileSystem, -902061641, string.Format("Fehler beim erstellen der VBS-Skript-Datei {0}: Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", vbsFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException));
                        Console.WriteLine("FEHLER VBS-Skrpt-Datei: {0}", vbsFilePath);
                        Program.AppErrorOccured = true;
                    }
                }
            }
            
        }

        /// <summary>
        /// Erstellt einen TaskSheduler Task.
        /// </summary>
        /// <param name="intervallMinutes"></param>
        /// <param name="scheduledTaskName"></param>
        /// <param name="taskPath"></param>
        /// <returns>true, wenn Task fehlerfrei erzeugt wurde.</returns>
        internal static bool CreateSchedulerTask(int intervallMinutes, string scheduledTaskName, string taskPath)
        {
            #region  Zur vollen Viertelstunde starten
            int min = 0;
            if (DateTime.Now.Minute <= 15) min = 15;
            else if (DateTime.Now.Minute <= 30) min = 30;
            else if (DateTime.Now.Minute <= 45) min = 45;

            // schtasks kann nur /ST HH:mm ohne Sekunden. Zum sekundengenauen Start sind Sekundenwerte über XML-Datei möglich. Noch nicht ausprobiert.
            string startTime = DateTime.Now.AddMinutes(min - DateTime.Now.Minute).ToShortTimeString();
            #endregion
                         
            // InTouch-Werte können nur durch den Benutzer empfangen werden, der grade angemeldet ist. Daher kein /RU + /RP möglich!
            string schtasksCommand = String.Format("/Create /SC Minute /MO {0} /TN \\KKT\\{1} /TR \"{2}\" /ST {3}", intervallMinutes, scheduledTaskName, taskPath, startTime);

            Log.Write(Log.Category.Scheduler, 1902040956, string.Format("Neuer Task wird erstellt. Intervall: {0} min; Startzeit: {1} ",  intervallMinutes, startTime));

            ProcessStartInfo start = new ProcessStartInfo
            {
                FileName = "schtasks.exe", // Specify exe name.
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                Arguments = schtasksCommand
            };

            using (Process process = Process.Start(start))
            {
                // Read the error stream first and then wait.
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                //Check for an error
                if (!String.IsNullOrEmpty(error))
                {
                    //Console.WriteLine("Fehler beim erstellen des SchedulerTasks {0} - Fehlertext: {1}", scheduledTaskName, error);
                    return false;
                }
                else
                {
                    //Console.WriteLine("Task wurde fehlerfrei erstellt.");
                    return true;
                }
            }
        }

        /// <summary>
        /// Fragt ab, ob der Task taskname (ohne Ordnerpfad) vorhanden ist.
        /// Keine Überprüfung, ob Task aktiv!
        /// </summary>
        /// <param name="taskname"></param>
        /// <returns></returns>
        private static bool CheckForSchedulerTask(string taskname)
        {
            ProcessStartInfo start = new ProcessStartInfo
            {
                FileName = "schtasks.exe", // Specify exe name.
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                Arguments = "/query /TN \\KKT\\" + taskname,
                RedirectStandardOutput = true
            };

            try
            {
                using (Process process = Process.Start(start))
                {
                    // Read in all the text from the process with the StreamReader.
                    using (StreamReader reader = process.StandardOutput)
                    {
                        string stdout = reader.ReadToEnd();
  
                        if (stdout.Contains(taskname))
                        {
                            return true;
                        }
                        else
                        {    
                            return false;
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.Scheduler, -902011221, string.Format("Fehler beim prüfen des TaskSchedulers: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                Program.AppErrorOccured = true;
                return false;
            }

        }


        /// <summary>
        /// Löscht den angegebenen Task im Ordner \KKT\
        /// </summary>
        /// <param name="taskname"></param>
        internal static void DeleteSchedulerTask()
        {            
            string schtasksCommand = String.Format("/Delete /TN \\KKT\\{0} /F", taskName);

            Log.Write(Log.Category.Scheduler, 2001301027, string.Format("Task {0} wird gelöscht.", taskName));

            ProcessStartInfo start = new ProcessStartInfo
            {
                FileName = "schtasks.exe", // Specify exe name.
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                Arguments = schtasksCommand
            };

            using (Process process = Process.Start(start))
            {
                // Read the error stream first and then wait.
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                //Check for an error
                if (!String.IsNullOrEmpty(error))
                {
                    Log.Write(Log.Category.Scheduler, -001301029, string.Format("Task {0} konnte nicht gelöscht werden. {1}", taskName, error));
                    //Console.WriteLine("Fehler beim erstellen des SchedulerTasks {0} - Fehlertext: {1}", scheduledTaskName, error);                    
                }
            }
        }

    }
}
