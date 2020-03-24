using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;

namespace Kreutztraeger
{
    /*
     * LISTE HARDCODIERTER INTOUCH-VARIABLEN:
     * - ExT_Druck
     * - ExM_Druck
     * - $Operator
     * - DBGruppe
     * - DBStartTime
     * - DBEndTime 
     * - DBvonPrio 
     * - DBbisPrio
     * - ExBestStdWerte
     * - ExLöscheStdMin
     * - ExLösche15StdMin
     * - SchockNr
     * - ChargenNrS
     * - SchockRProg + SchockNr
     * - SchockSollDauer
     * - 
     * 
     */

    public static class InTouch
    {      
        /// <summary>
        /// Lese tagName aus InTouch (view.exe) 
        /// </summary>
        /// <param name="tagName">Name des zu lesenden Tags.</param>
        /// <returns>object mit Tag-Wert. Muss durch cast in gewünschten Datentyp gewandelt werden. Später ggf. per Delegaten ansprechen, um cast zu vermeiden?</returns>
        public static object ReadTag(string tagName)
        {
            object result = null;

            try
            {
                NativeMethods intouch = new NativeMethods(0, 0);

                switch (intouch.GetTagType(tagName))
                {
                    case 1: //PT_DISCRETE:
                        result = intouch.ReadDiscrete(tagName);
                        break;
                    case 2: //PT_INTEGER:
                        result = intouch.ReadInteger(tagName);
                        break;
                    case 3: //PT_REAL:
                        result = intouch.ReadFloat(tagName);
                        break;
                    case 4: //PT_STRING:
                        result = intouch.ReadString(tagName); 
                        break;
                }

                return result;
            }
            catch (DllNotFoundException dllex)
            {
                Log.Write(Log.Category.InTouchDB, -902011207, string.Format("InTouch-Bibliotheken nicht bereit: {0}", dllex.Message));
                Program.AppErrorOccured = true;
                return null;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.InTouchDB, -902011210, string.Format("Fehler beim Lesen aus InTouch für TagName >{4}<: \r\n\t\tTyp: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}  \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace, tagName));
                Program.AppErrorOccured = true;                
                return null;
            }

        }

        /// <summary>
        /// Setzt InTouch-Bits, um Excel-Alarmmeldungen zu steuern.        
        /// </summary>
        /// <param name="ErrorOccured">true = Fehler beim Ausführen dieses Programms aufgetreten</param>
        internal static void SetExcelAliveBit(bool AppErrorOccured)
        {

            if (AppErrorOccured)
            {
                //Ein Fehler ist in diesem Programm aufgetreten.                
                //Setze Meldung "Exceltabellen werden nicht gefüllt.."
                WriteDiscTag(Program.InTouchDiscAlarm , true);
            }
            else
            {
                //Es ist kein Fehler aufgetreten.                 
                //Rücksetzen von TimeoutBit.
                WriteDiscTag(Program.InTouchDiscTimeOut, false);
                //Rücksetzen von Alarm-Bit
                WriteDiscTag(Program.InTouchDiscAlarm, false);               
            }
        }

        /// <summary>
        /// Beschreibt eine Discrete Variable in InTouch
        /// </summary>
        /// <param name="tagName">Variablenname</param>
        /// <param name="value">Zuweisungswert</param>
        internal static void WriteDiscTag (string tagName, bool value)
        {
            try
            {
                short val = 0;
                if (value) val = 1;

                NativeMethods intouch = new NativeMethods(0, 0);

                intouch.WriteDiscrete(tagName, val);
               
            }
            catch (DllNotFoundException dllex)
            {
                Log.Write(Log.Category.InTouchDB, -905071615, string.Format("InTouch-Bibliotheken nicht bereit zum Schreiben: {0}", dllex.Message));
                Program.AppErrorOccured = true;
            }
            catch (Exception ex)
            {
                Log.Write(Log.Category.InTouchDB, -905071616, string.Format("Fehler beim Schreiben in InTouch für TagName >{4}<: \r\n\t\tTyp: {0} \r\n\t\t Fehlertext: {1}  \r\n\t\t InnerException: {2}  \r\n\t\t StackTrace: {3}", ex.GetType().ToString(), ex.Message, ex.InnerException, ex.StackTrace, tagName));
                Program.AppErrorOccured = true;
            }
        }

    }

    public class NativeMethods
    {
        //für 64-Bit Betriebssystem
        const string ptaccPath = @"C:\Program Files (x86)\Wonderware\InTouch\ptacc.dll";
        const string wwheapPath = @"C:\Program Files (x86)\Common Files\ArchestrA\wwheap.dll";

        //für 32-Bit Betriebssystem
        //const string ptaccPath = @"C:\Program Files\Wonderware\InTouch\ptacc.dll";
        //const string wwheapPath = @"C:\Program Files\Common Files\ArchestrA\wwheap.dll";

        public static string PtaccPath { get; } = ptaccPath;
        public static string WwheapPath { get; } = wwheapPath;

        // wwHeap_Register: Verstoß gegen Benennungsregel ignoieren, sonst Fehler!
        [DllImport(wwheapPath)]
        static extern bool wwHeap_Register(int hWnd, ref short NotifyMessage);

        // wwHeap_Unregister: Verstoß gegen Benennungsregel ignoieren, sonst Fehler!
        [DllImport(wwheapPath)]
        static extern bool wwHeap_Unregister();

        [DllImport(ptaccPath)]
        static extern int PtAccInit(int hWnd, short nExtra);

        [DllImport(ptaccPath)]
        static extern int PtAccShutdown(int accid);

        [DllImport(ptaccPath)]
        static extern int PtAccOK(int accid);

        //nicht Mashalling auf Unicode, sonst Variablen nicht erreichbar
        [DllImport(ptaccPath)]
        static extern int PtAccHandleCreate(int accid, string nme);

        [DllImport(ptaccPath)]
        static extern int PtAccHandleActivate(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern int PtAccHandleDeactivate(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern int PtAccHandleDelete(int accid, int hPt);

        //nicht Mashalling auf Unicode, sonst Variablen nicht erreichbar
        [DllImport(ptaccPath)]
        static extern int PtAccActivate(int accid, string ptname);

        [DllImport(ptaccPath)]
        static extern int PtAccType(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern int PtAccDeactivate(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern int PtAccDelete(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern int PtAccSetExtraInt(int accid, int hPt, short nOffset, short nValue);

        [DllImport(ptaccPath)]
        static extern int PtAccGetExtraInt(int accid, int hPt, short nOffset);

        [DllImport(ptaccPath)]
        static extern int PtAccSetExtraLong(int accid, int hPt, short nOffset, int lValue);

        [DllImport(ptaccPath)]
        static extern int PtAccGetExtraLong(int accid, int hPt, short nOffset);

        [DllImport(ptaccPath)]
        static extern int PtAccReadD(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern int PtAccReadI(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern float PtAccReadR(int accid, int hPt);

        [DllImport(ptaccPath)]
        static extern double PtAccReadA(int accid, int hPt);

        //nicht Mashalling auf Unicode, sonst Variablen nicht erreichbar
        [DllImport(ptaccPath, CallingConvention = CallingConvention.Winapi)]
        static extern int PtAccReadM(int accid, int hPt, StringBuilder nm, int nMax);
        //static extern int PtAccReadM(int accid, int hPt, string nm, int nMax);

        [DllImport(ptaccPath)]
        static extern int PtAccWriteD(int accid, int hPt, short value);

        [DllImport(ptaccPath)]
        static extern int PtAccWriteI(int accid, int hPt, int value);

        [DllImport(ptaccPath)]
        static extern int PtAccWriteR(int accid, int hPt, float value);

        [DllImport(ptaccPath)]
        static extern int PtAccWriteA(int accid, int hPt, double value);

        //nicht Mashalling auf Unicode, sonst Variablen nicht erreichbar
        [DllImport(ptaccPath)]
        static extern int PtAccWriteM(int accid, int hPt, string value);

       
        public bool ExitFlag = false;
        private readonly short xStorage;
        readonly int hWnd;
        short NotifyMessage = 0;
        int intResult;
        float floatResult;
        double doubleResult;
        string stringResult;

        public NativeMethods(int _hWnd, short _xStorage)
        {
            // 'Set defaults
            hWnd = _hWnd;
            xStorage = _xStorage;
        }

        /// <summary>
        /// Gibt die Variablenart einer InTouch-Variablen anhand des Namens aus.
        /// </summary>
        /// <param name="TagName">InTouch TagName</param>
        /// <returns>0 = Fehler, 1 = int (Discrete), 2 = Integer, 3 = float (Real)</returns>
        public int GetTagType(string TagName)
        {
            // Returns the current value of an InTouch INTEGER variable.
            int accid;
            int hPt;            
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(hWnd, xStorage);
            if (accid == 0)
            {
                ExitFlag = true;
            }
            else
            {
                // This code reads in the tagnames 
                hPt = PtAccActivate(accid, TagName);
                //Log.Write(0, -1, string.Format("Intouch-TagName >{0}< hPt = >{1}<", TagName, hPt));

                intResult = PtAccType(accid, hPt);
                PtAccShutdown(accid);
            }
            return intResult;
        }

        public int ReadInteger(string TagName)
        {
            // Returns the current value of an InTouch INTEGER variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(hWnd, xStorage);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code reads in the tagnames 
                hPt = PtAccActivate(accid, TagName);
                intResult = PtAccReadI(accid, hPt);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }
            return intResult;
        }

        public bool ReadDiscrete(string TagName)
        {
            // Returns the current value of an InTouch DISCRETE variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(hWnd, xStorage);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code reads in the tagnames 
                hPt = PtAccActivate(accid, TagName);
                intResult = PtAccReadD(accid, hPt);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }

            if (intResult == 0) return false;
            else return true;
        }

        public float ReadFloat(string TagName)
        {
            // Returns the current value of an InTouch FLOATING point variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(hWnd, xStorage);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code reads in the tagnames 
                hPt = PtAccActivate(accid, TagName);
                floatResult = PtAccReadR(accid, hPt);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }
            return floatResult;
        }

        public double ReadDIF(string TagName)
        {
            // Returns the current value of an InTouch DISCRETE, INTEGER or FLOATING point variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(hWnd, xStorage);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code reads in the tagnames 
                hPt = PtAccActivate(accid, TagName);
                doubleResult = PtAccReadA(accid, hPt);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }
            return doubleResult;
        }

        public string ReadString(string strTagName)
        {
            // Returns the current value of an InTouch STRING variable.
            int accid;
            int hPt;
            StringBuilder strStringTag = new StringBuilder(131);    //neu 21.03.2019: Stringbuilder, da Übergabe an C++; Quelle: https://stackoverflow.com/questions/20752001/passing-strings-from-c-sharp-to-c-dll-and-back-minimal-example      

            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(hWnd, xStorage);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code reads in the tagnames 
                hPt = PtAccActivate(accid, strTagName);
                if (hPt == 0)
                {
                    Log.Write(Log.Category.InTouchVar, 1903151310, "Message Variable nicht gefunden: " + strTagName);
                    return null; //TagName nicht gefunden. neu 10.04.2019 HaSch
                }

                // intResult gleich stringResult.Length
                intResult = PtAccReadM(accid, hPt, strStringTag, strStringTag.Capacity);

                //Log.Write(Log.Category.InTouchVar, 1903151311, string.Format("\r\n####### ReadString \r\n\tIntouch-TagName >{0}< -> \r\n\tLänge Platzhalter 131 \r\n\tLänge Ergebnis >{1}< \r\n\tInhalt >{2}< \r\n\tLänge Inhalt: {4}\r\n\thPt >{3}<\r\n####### ", strTagName, intResult, strStringTag, hPt, strStringTag.Length) );

                stringResult = strStringTag.ToString();

                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }
            return stringResult;
        }

        public void WriteInteger(string TagName, int Value)
        {
            // Sets a new value into an InTouch INTEGER variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(0, 0);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code writes value to "TagName"
                hPt = PtAccActivate(accid, TagName);
                PtAccWriteI(accid, hPt, Value);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }

        }

        public void WriteDiscrete(string TagName, short Value)
        {
            // Sets a new value into an InTouch DISCRETE variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(0, 0);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code writes value to "TagName"
                hPt = PtAccActivate(accid, TagName);
                PtAccWriteD(accid, hPt, Value);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }

        }

        public void WriteFloat(string TagName, float Value)
        {
            // Sets a new value into an InTouch FLOATING point variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(0, 0);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code writes value to "TagName"
                hPt = PtAccActivate(accid, TagName);
                PtAccWriteR(accid, hPt, Value);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }

        }

        public void WriteDIF(string TagName, double Value)
        {
            // Sets a new value into an InTouch DISCRETE, INTEGER or FLOATING point variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(0, 0);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code writes value to "TagName"
                hPt = PtAccActivate(accid, TagName);
                PtAccWriteA(accid, hPt, Value);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }

        }

        public void WriteString(string TagName, string Value)
        {
            // Sets a new value into an InTouch STRING variable.
            int accid;
            int hPt;
            // This calls wwHeap_Register - This must be called before any PtAcc calls.
            wwHeap_Register(0, ref NotifyMessage);
            // This initializes Ptacc.dll
            accid = PtAccInit(0, 0);
            if ((accid == 0))
            {
                ExitFlag = true;
            }
            else
            {
                // This code writes value to "TagName"
                hPt = PtAccActivate(accid, TagName);
                PtAccWriteM(accid, hPt, Value);
                // This shuts down Ptacc.dll
                PtAccShutdown(accid);
            }

        }

    }

}
