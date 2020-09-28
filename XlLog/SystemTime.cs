using System;
using System.Runtime.InteropServices;

namespace Kreutztraeger
{
    //Quelle: https://dotnet-snippets.de/snippet/setzen-der-systemzeit/58

    public class SetNewSystemTime //Fehlernummern siehe Log.cs 12YYZZ
    {
        struct str_Zeit
        {
            public short Jahr;
            public short Monat;
            public short TagInDerWoche;
            public short Tag;
            public short Stunde;
            public short Minute;
            public short Sekunde;
            public short Millisekunde;
        }

        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool SetSystemTime(ref str_Zeit neueZeit);

        private static void SetzeSystemzeit(DateTime NeueZeit) //Fehlernummern siehe Log.cs 1201ZZ
        {
            NeueZeit = NeueZeit.ToUniversalTime();

            str_Zeit Zeit = new str_Zeit
            {
                Jahr = (short)NeueZeit.Year,
                Monat = (short)NeueZeit.Month,
                TagInDerWoche = (short)NeueZeit.DayOfWeek,
                Tag = (short)NeueZeit.Day,
                Stunde = (short)NeueZeit.Hour,
                Minute = (short)NeueZeit.Minute,
                Sekunde = (short)NeueZeit.Second,
                Millisekunde = (short)NeueZeit.Millisecond
            };

            bool result = SetSystemTime(ref Zeit);

            //If the function succeeds, the return value is nonzero.
            if (result)
            {
                Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 120102,
                    "Der Benutzer >" + InTouch.ReadTag("$Operator") + "< hat die Systemzeit nach >" + NeueZeit.ToLocalTime() + "< (" + NeueZeit + " UTC) umgestellt.");
            }
            else
            {
                Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 120103,
                "Der Benutzer >" + InTouch.ReadTag("$Operator") + "< konnte die Systemzeit nicht nach >" + NeueZeit + "< umstellen. Fehler aus kernel32.dll");
            }
        }

        private static void SetzeSystemzeit2(DateTime NeueZeit)
        {
            string powerShellCommand = String.Format("Set-Date -Date {0}", NeueZeit);

            System.Diagnostics.ProcessStartInfo start = new System.Diagnostics.ProcessStartInfo
            {
                FileName = "powershell", // Specify exe name.
                UseShellExecute = false,
                CreateNoWindow = true,
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                Arguments = powerShellCommand
            };

            using (System.Diagnostics.Process process = System.Diagnostics.Process.Start(start))
            {
                // Read the error stream first and then wait.
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();

                //Check for an error
                if (!String.IsNullOrEmpty(error))
                {
                    Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 120105,
                        "Der Benutzer >" + InTouch.ReadTag("$Operator") + "< konnte die Systemzeit nicht nach >" + NeueZeit + "< umstellen. Fehler aus PowerShell");
                }
                else
                {
                    Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 120106,
                        "Der Benutzer >" + InTouch.ReadTag("$Operator") + "< hat die Systemzeit per PowerShell nach >" + NeueZeit + "< umgestellt.");
                }
            }
        }


        /// <summary>
        /// Interpretiert newDateTime als Zeit. Bei Erfolg wird die Systemzeit auf diesen Wert gestellt und der Task XlLog_vbs15 gelöscht und neu erstellt.
        /// </summary>
        /// <param name="newDateTime"></param>
        public static void SetNewSystemtimeAndScheduler(string newDateTime) //Fehlernummern siehe Log.cs 1202ZZ
        {
            var ci = new System.Globalization.CultureInfo("de-DE");
            if (DateTime.TryParseExact(newDateTime, "dd.MM.yyyy HH:mm:ss", ci, System.Globalization.DateTimeStyles.None , out DateTime dateTime))
            {
                SetzeSystemzeit(dateTime);

                Scheduler.DeleteSchedulerTask();

                Scheduler.CeckOrCreateTaskScheduler();
            }
            else
            {
                Log.Write(Log.Cat.Scheduler, Log.Prio.Error, 120202,  "Neue Systemzeit konnte nicht gesetzt werden. Uhrzeit nicht erkannt in Zeichenfolge >" + newDateTime + "<");
            }
        }

    }

}
 