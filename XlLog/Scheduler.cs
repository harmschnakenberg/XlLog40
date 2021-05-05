using System;
using System.Diagnostics;
using System.IO;

namespace Kreutztraeger
{
    class Scheduler //Fehlernummern siehe Log.cs 10YYZZ
    {

        public static int StartTaskIntervallMinutes { get; set; }  = 15;
        static readonly string taskName = string.Format(@"XlLog_vbs{0}", StartTaskIntervallMinutes);
        public static bool UseTaskScheduler { get; set; } = true;

        /// <summary>
        /// Prüft, ob eine *.vbs-Datei und ein SchedulerTask für die automatische Ausführung vorhanden sind und erstellt diese ggf. 
        /// </summary>
        internal static void CeckOrCreateTaskScheduler() //Fehlernummern siehe Log.cs 1001ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 100101, string.Format("CeckOrCreateTaskScheduler()"));
            
            if (!UseTaskScheduler) return;

            try
            {
                string taskPath = System.Reflection.Assembly.GetAssembly(typeof(Program)).Location;  // Diese exe

                //Erstelle *.vbs-Dateien für versteckte Ausführung
                //neu 22.04.2020 *.vbs-Dateien entfallen, da unsichtbare Ausführung aus InTouch möglich.

                string vbsFilePath = Path.ChangeExtension(taskPath, "vbs");
                string fileContent = string.Format("Set WshShell = WScript.CreateObject(\"WScript.Shell\")\r\nWshShell.Run \"{0} -Task\",0,True", taskPath);
                CreateVbsScript(taskPath, vbsFilePath, fileContent);

                if (!CheckForSchedulerTask(taskName))
                {
                    if (!CreateSchedulerTask(StartTaskIntervallMinutes, taskName, vbsFilePath))
                    {
                        Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 100102, "Es konnte kein neuer Task erstellt werden.");
                        //Program.AppErrorOccured = true;
                    }
                }
                else
                {
                    // Log.Write(Log.DebugCategory.WriteToConsole, 1902011219, "Task ist im TaskScheduler vorhanden.");
                }
            }
            catch (Exception ex)
            {
                Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 100103, string.Format("Fehler beim bearbeiten des TaskSchedulers: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                //Program.AppErrorOccured = true;
            }
        }

        /// <summary>
        /// VBS-Datei für versteckte Ausführung der exe
        /// </summary>
        /// <param name="exeFilePath"></param>
        private static void CreateVbsScript(string exeFilePath, string vbsFilePath, string fileContent) //Fehlernummern siehe Log.cs 1002ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 100201, string.Format("CreateVbsScript({0}, {1}, {2} ", exeFilePath, vbsFilePath, fileContent));
            // Um die Anwendung unter dem zurzeit angemeldeten Benutzer versteckt zu starten, wird der Umweg über eine *.vbs-Datei gewählt.
            // Bei Starten im TaskScheduler mit fester Benutzeranmeldung (= Versteckt) werden InTouch-Tags nicht mehr erreicht.
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
                        Log.Write(Log.Cat.FileSystem, Log.Prio.Error, 100202, string.Format("Fehler beim erstellen der VBS-Skript-Datei {0}: Typ: {1} \r\n\t\t Fehlertext: {2}  \r\n\t\t InnerException: {3}", vbsFilePath, ex.GetType().ToString(), ex.Message, ex.InnerException));
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
        internal static bool CreateSchedulerTask(int intervallMinutes, string scheduledTaskName, string taskPath) //Fehlernummern siehe Log.cs 1004ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 100401, string.Format("CreateSchedulerTask({0}, {1}, {2})", intervallMinutes, scheduledTaskName, taskPath));

            #region  Zur vollen Viertelstunde starten
            int min = 0;
            if (intervallMinutes < 60)
            {
                if (DateTime.Now.Minute <= 15) min = 15;
                else if (DateTime.Now.Minute <= 30) min = 30;
                else if (DateTime.Now.Minute <= 45) min = 45;
            }
            //Bei stündlicher Wiederholung nur zur vollen Stunde starten.
            

            // schtasks kann nur /ST HH:mm ohne Sekunden. Zum sekundengenauen Start sind Sekundenwerte über XML-Datei möglich. Noch nicht ausprobiert.
            string startTime = DateTime.Now.AddMinutes(min - DateTime.Now.Minute).ToShortTimeString();
            #endregion
                         
            // InTouch-Werte können nur durch den Benutzer empfangen werden, der grade angemeldet ist. Daher kein /RU + /RP möglich!
            string schtasksCommand = String.Format("/Create /SC Minute /MO {0} /TN \\KKT\\{1} /TR \"wscript \\\"{2}\"\" /ST {3}", intervallMinutes, scheduledTaskName, taskPath, startTime);

            Console.WriteLine("schtasks.exe " + schtasksCommand);
            Log.Write(Log.Cat.Scheduler, Log.Prio.LogAlways, 100402, string.Format("Neuer Task wird erstellt. Intervall: {0} min; Startzeit: {1} ",  intervallMinutes, startTime));

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
        private static bool CheckForSchedulerTask(string taskname) //Fehlernummern siehe Log.cs 1005ZZ
        {
            Log.Write(Log.Cat.MethodCall, Log.Prio.Info, 100501, string.Format("CheckForSchedulerTask({0})", taskname));

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
                Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 100502, string.Format("Fehler beim prüfen des TaskSchedulers: Typ: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}", ex.GetType().ToString(), ex.Message, ex.InnerException));
                Program.AppErrorOccured = true;
                return false;
            }

        }


        /// <summary>
        /// Löscht den angegebenen Task im Ordner \KKT\
        /// </summary>
        /// <param name="taskname"></param>
        internal static void DeleteSchedulerTask() //Fehlernummern siehe Log.cs 1006ZZ
        {            
            string schtasksCommand = String.Format("/Delete /TN \\KKT\\{0} /F", taskName);

            Log.Write(Log.Cat.Scheduler, Log.Prio.LogAlways, 100602, string.Format("Task {0} wird gelöscht.", taskName));

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
                    Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 100603, string.Format("Task {0} konnte nicht gelöscht werden. {1}", taskName, error));
                    //Console.WriteLine("Fehler beim erstellen des SchedulerTasks {0} - Fehlertext: {1}", scheduledTaskName, error);                    
                }
            }
        }

    }
}
