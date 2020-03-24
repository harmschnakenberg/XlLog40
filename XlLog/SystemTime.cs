using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;

namespace Kreutztraeger
{
    //Quelle: https://dotnet-snippets.de/snippet/setzen-der-systemzeit/58

    public class SetNewSystemTime
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

        private static void SetzeSystemzeit(DateTime NeueZeit)
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

            SetSystemTime(ref Zeit);
        }

        /// <summary>
        /// Interpretiert newDateTime als Zeit. Bei Erfolg wird die Systemzeit auf diesen Wert gestellt und der Task XlLog_vbs15 gelöscht und neu erstellt.
        /// </summary>
        /// <param name="newDateTime"></param>
        public static void SetNewSystemtimeAndScheduler(string newDateTime)
        {
            if (DateTime.TryParse(newDateTime, out DateTime dateTime))
            {
                SetzeSystemzeit(dateTime);

                Scheduler.DeleteSchedulerTask();

                Scheduler.CeckOrCreateTaskScheduler();
            }
            else
            {
                Log.Write(Log.Category.Scheduler, 2001301013, "Neue Systemzeit konnte nicht gesetzt werden. Uhrzeit nicht erkannt in >" + newDateTime + "<");
            }
        }

    }

}