using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace Kreutztraeger
{
    class Tools
    {
        /// <summary>
        /// Sekunden, die das CMD-Fenster offen bleibt, wenn durch Benutzer gestartet.  
        /// </summary>
        internal static int WaitToClose { get; set; } = 20;

        /// <summary>
        /// Sekunden, die das Programm auf Skripte in InTouch wartet (nur Ausführung durch Task)
        /// </summary>
        internal static int WaitForScripts { get; set; } = 10;

        /// <summary>
        /// Wartet und zählt die Sekunden in der Konsole runter.
        /// </summary>
        /// <param name="seconds">Sekunden, die gewartet werden sollen.</param>
        internal static void Wait(int seconds)
        {
            while (seconds > 0)
            {
                Console.Write(seconds.ToString("00"));
                --seconds;
                Thread.Sleep(1000);
                Console.Write("\b\b");
            }
        }

        public static bool IsBitSet(int b, int pos)
        {
            return (b & (1 << pos)) != 0;
        }

        public class Tuple<T1, T2>
        {
            public T1 First { get; private set; }
            public T2 Second { get; private set; }
            internal Tuple(T1 first, T2 second)
            {
                First = first;
                Second = second;
            }
        }

        public static class Tuple
        {
            public static Tuple<T1, T2> New<T1, T2>(T1 first, T2 second)
            {
                var tuple = new Tuple<T1, T2>(first, second);
                return tuple;
            }
        }
    }
}
