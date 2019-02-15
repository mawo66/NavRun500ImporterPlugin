using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.IO;

using ZoneFiveSoftware.Common.Visuals;
using ZoneFiveSoftware.Common.Visuals.Fitness;
using ZoneFiveSoftware.Common.Data;
using ZoneFiveSoftware.Common.Data.Fitness;
using ZoneFiveSoftware.Common.Data.GPS;

namespace NavRun500ImporterPlugin
{
    class NavRunInterpreter
    {
        private List<UInt32> adrlist;
        //private StreamWriter debugWriter;


        [StructLayout(LayoutKind.Sequential, Pack=1)]
        private struct ActivitySummary
        {
            public UInt16  usAnzahlSamples;
            public byte    ucAnzahlRunden;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 6)]
            public byte[]  pucStartzeit;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 3)]
            public byte[]  pucTrainingszeit;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[]  pucUnbekannt;
            public UInt32  ulDistanz;
            public UInt16  usGeschwDurchschn;
            public UInt16  usGeschwMax;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[]  pucUnbekannt3;
            public byte    ucHerzfreqDurchschn;
            public byte    ucHerzfreqMax;
            public byte    ucHerzfreqMin;
            public byte    pucUnbekannt4;
            public UInt32  ulKalorien;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 3)]
            public byte[]  pucZeitUnterhalbZone;
            public byte    pucUnbekannt6;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 3)]
            public byte[]  pucZeitInZone;
            public byte    pucUnbekannt7;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 3)]
            public byte[]  pucZeitOberhalbZone;
        }

        [StructLayout(LayoutKind.Sequential, Pack=1)]
        private struct FullSample
        {
            public byte   ucXXX;
            public byte   ucSat;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 6)]
            public byte[] pucDatumZeit;
            public Int32  slLaengengrad;
            public Int32  slBreitengrad;
            public Int16  ssHoehe;
            public UInt16 usRichtung;
            public UInt16 usDistanz;
            public UInt16 usGeschw;
            public byte   ucHerzFreq;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct ShortSample
        {
            public byte   ucXXX;
            public byte   ucSat;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 2)]
            public byte[] pucZeit;
            public Int32  slLaengengradDiff;
            public Int32  slBreitengradDiff;
            public Int16  ssHoeheDiff;
            public UInt16 usRichtung;
            public UInt16 usDistanzDiff;
            public UInt16 usGeschw;
            public byte   ucHerzFreq;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct EmptySample
        {
            public byte ucXXX;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 6)]
            public byte[] pucDatumZeit;
            public byte ucHerzFreq;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct NoSample
        {
            public byte ucXXX;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 2)]
            public byte[] pucZeit;
        }

        [StructLayout(LayoutKind.Sequential, Pack = 1)]
        private struct RundenInfo
        {
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 4)]
            public byte[]  pucZwischenzeit;
            public byte  ucHerzfreqDurchschn;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 3)]
            public byte[]  pucUnbekannt;
            public UInt32  ulDistanz;
            public UInt16 usGeschw;
            [MarshalAsAttribute(UnmanagedType.ByValArray, SizeConst = 2)]
            public byte[]  pucUnbekannt2;
        }

        /// <summary>
        /// Konstruktor.
        /// </summary>
        public NavRunInterpreter()
        {
            adrlist = new List<UInt32>();
            //debugWriter = new StreamWriter("Debug.log");
        }


        /// <summary>
        /// Ermittelt die Anzahl der Aktivitäten im Speicherabbild.
        /// </summary>
        /// <param name="_Data">Speicherabbild.</param>
        /// <returns>Anzahl der gefundenen Aktivitäten.</returns>
        public int GetNumberOfActivities(ref byte[] _Data)
        {
            int ActCount=0;
            int AktAdr = 0x100;
            bool bFertig=false;

            if (_Data.Length < 0x1000) return (0);

            if (_Data[AktAdr] == 0xFF) return (0);

            ActCount++;
            adrlist.Add((UInt32)_Data[AktAdr] << 12);

            while (!bFertig)
            {
                while (_Data[AktAdr] != 0xFF) AktAdr++;

                AktAdr++;
                if (_Data[AktAdr] != 0xFF)
                {
                    ActCount++;
                    adrlist.Add((UInt32)_Data[AktAdr] << 12);
                }
                else
                {
                    bFertig = true;
                }
            }

            return (ActCount);
        }


        /// <summary>
        /// Ermittelt alle Daten zu einer Aktivität.
        /// </summary>
        /// <param name="_Data">Speicherabbild.</param>
        /// <param name="_ActivityNumber">Nummer der Aktivität.</param>
        /// <param name="_ImportResults">Referenz auf Objekt für Rückgabe.</param>
        /// <returns>true bei Erfolg.</returns>
        public bool ImportActivity(ref byte[] _Data, int _ActivityNumber, IImportResults _ImportResults)
        {
            ActivitySummary Summary;
            FullSample FSample;
            ShortSample SSample;
            EmptySample ESample;
            NoSample NSample;
            IActivity activity;
            RundenInfo Runde;
            UInt32 ulAdr;
            UInt32 ulCount=0;

            if (_ActivityNumber >= adrlist.ToArray().Length) return (false);

            //debugWriter.WriteLine("Importiere Nr. " + _ActivityNumber.ToString());

            try
            {
                Summary = (ActivitySummary)ByteArrayToStruct(_Data, adrlist[_ActivityNumber], typeof(ActivitySummary));
            }
            catch (Exception)
            {
                //debugWriter.WriteLine("Exception beim Umwandeln in Struktur.");
                return (false);
            }

            // Startzeit
            DateTime aktDateTime = new DateTime(2000 + Summary.pucStartzeit[5], 
                                                Summary.pucStartzeit[4],
                                                Summary.pucStartzeit[3],
                                                Summary.pucStartzeit[2],
                                                Summary.pucStartzeit[1],
                                                Summary.pucStartzeit[0]);

            // Aktivität anlegen
            activity = _ImportResults.AddActivity(aktDateTime);
            activity.HasStartTime = true;

            // Zusammenfassung
            activity.AverageHeartRatePerMinuteEntered = Summary.ucHerzfreqDurchschn;
            activity.TotalCalories = (float)(Summary.ulKalorien / 100.0);
            activity.TotalDistanceMetersEntered = (float)(Summary.ulDistanz / 10.0);
            TimeSpan trainingTime = new TimeSpan(Summary.pucTrainingszeit[2],
                                                 Summary.pucTrainingszeit[1],
                                                 Summary.pucTrainingszeit[0]);
            activity.TotalTimeEntered = trainingTime;

            //Runden
            DateTime rundenStart;
            DateTime rundenEnde;
            rundenStart = aktDateTime;
            ulAdr = adrlist[_ActivityNumber] + 0x40;
            for (byte r = 0; r < Summary.ucAnzahlRunden; r++)
            {
                Runde = (RundenInfo)ByteArrayToStruct(_Data, ulAdr, typeof(RundenInfo));

                TimeSpan zwischenzeit = new TimeSpan(0,
                                                     Runde.pucZwischenzeit[0],
                                                     Runde.pucZwischenzeit[1],
                                                     Runde.pucZwischenzeit[2],
                                                     10 * ConvertBCDToHex(Runde.pucZwischenzeit[3]));

                rundenEnde = aktDateTime.Add(zwischenzeit);
                TimeSpan dauer = rundenEnde.Subtract(rundenStart);

                ILapInfo lap = activity.Laps.Add(rundenStart, dauer);
                lap.TotalDistanceMeters = (float)(Runde.ulDistanz / 10.0);
                lap.AverageHeartRatePerMinute = Runde.ucHerzfreqDurchschn;

                rundenStart = rundenEnde;
                ulAdr += (UInt32)Marshal.SizeOf(Runde);
            }

            Int32 aktLaenge=0, aktBreite=0;
            Int16 aktHoehe=0;
            byte aktHerz=0;
            bool bValidHerz, bValidGPS;

            ulAdr = adrlist[_ActivityNumber] + 0x1000;

            for (UInt32 i = 0; i < Summary.usAnzahlSamples; i++)
            {
                bValidHerz = false;
                bValidGPS = false;

                if ((_Data[ulAdr] == 0x80) || (_Data[ulAdr] == 0x00))
                {
                    // vollständiges Sample
                    FSample = (FullSample)ByteArrayToStruct(_Data, ulAdr, typeof(FullSample));

                    aktDateTime = new DateTime(2000 + FSample.pucDatumZeit[0],
                                               FSample.pucDatumZeit[1],
                                               FSample.pucDatumZeit[2],
                                               FSample.pucDatumZeit[3],
                                               FSample.pucDatumZeit[4],
                                               FSample.pucDatumZeit[5]);

                    aktLaenge = FSample.slLaengengrad;
                    aktBreite = FSample.slBreitengrad;
                    aktHoehe = FSample.ssHoehe;
                    aktHerz = FSample.ucHerzFreq;
                    if ((FSample.ucSat & 0x0F) != 0)
                    {
                        bValidGPS = true;
                    }
                    if (aktHerz > 0)
                    {
                        bValidHerz = true;
                    }
                    ulAdr += (UInt32)Marshal.SizeOf(FSample);
                    ulCount++;
                }
                else if (_Data[ulAdr] == 0x01)
                {
                    // verkürztes Sample
                    SSample = (ShortSample)ByteArrayToStruct(_Data, ulAdr, typeof(ShortSample));

                    aktDateTime = SetNewTime(aktDateTime, SSample.pucZeit[0], SSample.pucZeit[1]);
                    aktLaenge += SSample.slLaengengradDiff;
                    aktBreite += SSample.slBreitengradDiff;
                    aktHoehe += SSample.ssHoeheDiff;
                    aktHerz = SSample.ucHerzFreq;
                    if ((SSample.ucSat & 0x0F) != 0)
                    {
                        bValidGPS = true;
                    }
                    if (aktHerz > 0)
                    {
                        bValidHerz = true;
                    }

                    ulAdr += (UInt32)Marshal.SizeOf(SSample);
                    ulCount++;
                }
                else if (_Data[ulAdr] == 0x03)
                {
                    // leeres Sample (nur Zeit und Hertfreq)
                    ESample = (EmptySample)ByteArrayToStruct(_Data, ulAdr, typeof(EmptySample));
                    aktDateTime = new DateTime(2000 + ESample.pucDatumZeit[0],
                                                      ESample.pucDatumZeit[1],
                                                      ESample.pucDatumZeit[2],
                                                      ESample.pucDatumZeit[3],
                                                      ESample.pucDatumZeit[4],
                                                      ESample.pucDatumZeit[5]);

                    aktHerz = ESample.ucHerzFreq;
                    if (aktHerz > 0)
                    {
                        bValidHerz = true;
                    }
                    ulAdr += (UInt32)Marshal.SizeOf(ESample);
                    ulCount++;
                }
                else if (_Data[ulAdr] == 0x02)
                {
                    // nur kurze Zeit
                    NSample = (NoSample)ByteArrayToStruct(_Data, ulAdr, typeof(NoSample));
                    aktDateTime = SetNewTime(aktDateTime, NSample.pucZeit[0], NSample.pucZeit[1]);
                    ulAdr += (UInt32)Marshal.SizeOf(NSample);
                }
                else if (_Data[ulAdr] == 0xFF)
                {
                    //debugWriter.WriteLine("Ende (Abbruch).");
                    break;
                }
                else
                {
                    //debugWriter.WriteLine("unbekanntes Sample " + String.Format("0x{0:X02}", _Data[ulAdr]) + " Adr=" + String.Format("0x{0:X08}", ulAdr));
                    break;
                }

                // Herzfrequenzsample eintragen, wenn vorhanden
                if (bValidHerz)
                {
                    if (activity.HeartRatePerMinuteTrack == null)
                    {
                        activity.HeartRatePerMinuteTrack = new NumericTimeDataSeries();
                    }
                    activity.HeartRatePerMinuteTrack.Add(aktDateTime, (float)aktHerz);
                }

                // GPS-Sample eintragen, wenn vorhanden
                if (bValidGPS)
                {
                    if (activity.GPSRoute == null)
                    {
                        activity.GPSRoute = new GPSRoute();
                        //debugWriter.WriteLine("New GPS-Route");
                    }
                    GPSPoint gps = new GPSPoint((float)(aktBreite / 10000000.0),
                                                (float)(aktLaenge / 10000000.0),
                                                (float)(aktHoehe));
                    activity.GPSRoute.Add(aktDateTime, gps);

                    //debugWriter.WriteLine("Point " + ulCount.ToString() + ": " + aktDateTime.ToString() + "   " + gps.LongitudeDegrees.ToString() + " " + gps.LatitudeDegrees.ToString());
                }
            }

            //debugWriter.Flush();
            return (true);
        }


        /// <summary>
        /// Ermittelt einen neuen Zeitwert aus einen alten Zeitwert und neuen Minuten-
        /// und Sekundeninformationen.
        /// </summary>
        /// <param name="_DateTime">Alter Zeitwert.</param>
        /// <param name="_NewMin">Neuer Minutenwert.</param>
        /// <param name="_NewSec">Neuer Sekundenwert.</param>
        /// <returns>Neuer Zeitwert.</returns>
        private DateTime SetNewTime(DateTime _DateTime, int _NewMin, int _NewSec)
        {
            DateTime newDateTime;
            
            int year = _DateTime.Year;
            int month = _DateTime.Month;
            int day = _DateTime.Day;
            int hour = _DateTime.Hour;
            int minute = _DateTime.Minute;
            int second = _DateTime.Second;

            if (_NewMin == minute)
            {
                // Minute stimmt noch -> Sekunde aktualisieren
                second = _NewSec;
                newDateTime = new DateTime(year, month, day, hour, minute, second);
            }
            else
            {
                // Überlauf
                second = _NewSec;
                newDateTime = new DateTime(year, month, day, hour, minute, second);
                newDateTime = newDateTime.AddMinutes(1);
            }
            return (newDateTime);
        }


        /// <summary>
        /// Ermittelt den Wert einer in BCD-Darstellung vorhandenen Zahl.
        /// </summary>
        /// <param name="bcd">Zahl in BCD Darstellung.</param>
        /// <returns>Wert.</returns>
        private int ConvertBCDToHex(int bcd)
        {
            int result;
            int digit1 = bcd >> 4;
            int digit2 = bcd & 0x0f;

            result = (10 * digit1) + digit2;
 
            return (result);
        }


        /// <summary>
        /// Kopiert Daten aus einem Byte-Array in eine entsprechende Strukture (struct). Die Struktur muss ein sequenzeilles Layout besitzen. ( [StructLayout(LayoutKind.Sequential)] 
        /// </summary>
        /// <param name="array">Das Byte-Array das die daten enthält</param>
        /// <param name="offset">Offset ab dem die Daten in die Struktur kopiert werden sollen.</param>
        /// <param name="structType">System.Type der Struktur</param>
        /// <returns></returns>
        private object ByteArrayToStruct(byte[] array, uint offset, Type structType)
        {
            if (structType.StructLayoutAttribute.Value != LayoutKind.Sequential)
                throw new ArgumentException("structType ist keine Struktur oder nicht Sequentiell.");

            int size = Marshal.SizeOf(structType);
            if (array.Length < (offset + size))
                throw new ArgumentException("Byte-Array hat die falsche Länge.");

            byte[] tmp = new byte[size];
            Array.Copy(array, offset, tmp, 0, size);

            GCHandle structHandle = GCHandle.Alloc(tmp, GCHandleType.Pinned);
            object structure = Marshal.PtrToStructure(structHandle.AddrOfPinnedObject(), structType);
            structHandle.Free();

            return structure;
        }

    }
}
