using System;
using System.Threading;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;
using System.IO.Ports;

using ZoneFiveSoftware.Common.Visuals;
using ZoneFiveSoftware.Common.Visuals.Fitness;
using ZoneFiveSoftware.Common.Data;
using ZoneFiveSoftware.Common.Data.Fitness;
using ZoneFiveSoftware.Common.Data.GPS;

namespace NavRun500ImporterPlugin
{
    class NavRunCom
    {
        private SerialPort m_Port;

        #region NavRunCom Constants
        private static readonly byte NAV_TELSTART1 = 0xA0;
        private static readonly byte NAV_TELSTART2 = 0xA2;
        private static readonly byte NAV_TELEND1 = 0xB0;
        private static readonly byte NAV_TELEND2 = 0xB3;
        private static readonly byte TEL_READVERSION = 0x10;
        private static readonly byte TEL_READMEMORY = 0x12;
        #endregion

        /// <summary>
        /// Konstruktor.
        /// </summary>
        public NavRunCom()
        {
            m_Port = null;
        }

        /// <summary>
        /// Sucht die Schnittstelle, an der ein NavRun 500 angeschlossen ist
        /// </summary>
        /// <returns>true wenn Gerät gefunden wurde.</returns>
        public bool Open()
        {
            m_Port = OpenPort();
            if (m_Port == null)
            {
                return (false);
            }
            else
            {
                return (true);
            }
        }

        /// <summary>
        /// Schliesst die Schnittstelle zum NavRun 500.
        /// </summary>
        public void Close()
        {
            if (m_Port != null)
            {
                m_Port.Close();
                m_Port = null;
            }
        }

        /// <summary>
        /// Sucht im Header die maximale Page, in der Daten vorhanden sind, die ausgelesen werden müssen.
        /// Im Header steht ab Adr 0x100 z.B. 01 02 03 04 FF 05 06 FF FF FF ...
        /// D.h. wenn mindenstens 2mal FF auftritt folgen keine Daten mehr.
        /// </summary>
        /// <param name="_Data">Das Byte-Array das den Header enthält.</param>
        /// <returns>Nummer der letzten Page.</returns>
        private byte GetMaxPage(byte[] _Data)
        {
            int AktAdr = 0x100;

            if (_Data.Length < 0x1000) return (0);
            if (_Data[AktAdr] == 0xFF) return (0);

            while ((AktAdr < 0xFFFF) && 
                   ((_Data[AktAdr] != 0xFF) || (_Data[AktAdr+1] != 0xFF)) )
            {
                AktAdr++;
            }

            return (_Data[AktAdr-1]);
        }

        /// <summary>
        /// Liest alle notwendigen Daten aus dem NavRun 500 aus.
        /// </summary>
        /// <param name="_Monitor">Referenz auf Monitorobjekt zur Aktualisierung der Statusanzeige.</param>
        /// <returns>Byte-Array mit Speicherabbild.</returns>
        public byte[] ReadData(IJobMonitor _Monitor)
        {
            UInt32 ulAdr;
            List<byte> bytelist = new List<byte>();
            byte[] data;
            byte[] memblock = new byte[0x80];
            byte ucPageCount, ucBlockCount, ucMaxPage;
            float fPercentStep, fPercent;

            // Header (1.Page) einlesen
            Debug.WriteLine("Lese Header ...");
            ulAdr = 0;
            for (ucBlockCount = 0; ucBlockCount < 32; ucBlockCount++)
            {
                data = ReadDeviceMemory(ulAdr, 0x80);

                // Fehlerprüfung
                if (data == null)
                {
                    Debug.WriteLine("Fehler beim Lesen: Adr=" + String.Format("0x{0:X08}",ulAdr) + " return null!");
                    return (null);
                }
                else if( data.Length < 0x81 )
                {
                    Debug.WriteLine("Fehler beim Lesen: Adr=" + String.Format("0x{0:X08}", ulAdr) + " return len=" + String.Format("0x{0:X02}", data.Length));
                    return (null);
                }

                // zum Speicher hinzufügen
                for (int i = 0; i < 0x80; i++)
                {
                    bytelist.Add(data[i+1]);
                }
                ulAdr += 0x80;
            }
            _Monitor.PercentComplete = (float)0.05;

            // verwendete Pages ermitteln
            ucMaxPage = GetMaxPage(bytelist.ToArray());
            Debug.WriteLine("maximale Page: " + String.Format("0x{0:X02}", ucMaxPage));

            if (ucMaxPage > 0)
            {
                fPercent = (float)0.1;
                fPercentStep = (float)0.95 / (ucMaxPage * 32);

                ulAdr = 0x1000;
                for (ucPageCount = 1; ucPageCount <= ucMaxPage; ucPageCount++)
                {
                    Debug.WriteLine("Lese Page " + String.Format("0x{0:X08}", ulAdr) + "...");
                    for (ucBlockCount = 0; ucBlockCount < 32; ucBlockCount++)
                    {
                        data = ReadDeviceMemory(ulAdr, 0x80);

                        // Fehlerprüfung
                        if (data == null)
                        {
                            Debug.WriteLine("Fehler beim Lesen: Adr=" + String.Format("0x{0:X08}", ulAdr) + " return null!");
                            return (null);
                        }
                        else if (data.Length < 0x81)
                        {
                            Debug.WriteLine("Fehler beim Lesen: Adr=" + String.Format("0x{0:X08}", ulAdr) + " return len=" + String.Format("0x{0:X02}", data.Length));
                            return (null);
                        }

                        // zum Speicher hinzufügen
                        for (int i = 0; i < 0x80; i++)
                        {
                            bytelist.Add(data[i + 1]);
                        }

                        ulAdr += 0x80;
                        fPercent += fPercentStep;
                        _Monitor.PercentComplete = fPercent;
                    }
                }
            }
            else
            {
                Debug.WriteLine("Keine weiteren Pages verwendet!");
               _Monitor.PercentComplete = 1;
            }

            Debug.WriteLine("gelesene Daten: " + String.Format("0x{0:X08}", bytelist.ToArray().Length) + " bytes");
            return (bytelist.ToArray());
        }


        /// <summary>
        /// Liest einen Speicherblock aus dem NavRun 500 aus.
        /// </summary>
        /// <param name="_ulAdr">Startadresse.</param>
        /// <param name="_ucSize">Anzahl an Bytes.</param>
        /// <returns>Byte-Array mit Speicherblock.</returns>
        private byte[] ReadDeviceMemory(UInt32 _ulAdr, byte _ucSize)
        {
            if (m_Port == null) return(null);

            byte[] packet = new byte[5];

            packet[0] = TEL_READMEMORY;
            packet[1] = (byte)_ulAdr;
            packet[2] = (byte)(_ulAdr >> 8);
            packet[3] = (byte)(_ulAdr >> 16);
            packet[4] = (byte)_ucSize;

            byte[] response = SendTel(m_Port, packet, 2);

            return(response);
        }


        /// <summary>
        /// Prüft, ob an einer bestimmten Schnittstelle ein NavRun 500 angeschlossen ist.
        /// Es wird ein Telegramm zur Versionsabfrage gesendet, dass korrekt beantwortet
        /// werden muss.
        /// </summary>
        /// <param name="_Port">Zu testende Schnittstelle.</param>
        /// <returns>true bei Erfolg.</returns>
        private bool ValidNavRunPort(SerialPort _Port)
        {
            _Port.ReadTimeout = 1000;
            _Port.Open();

            byte[] packet = new byte[1] { TEL_READVERSION };

            byte[] response = SendTel(_Port, packet, 1);

            if (response == null)
            { 
                return(false);
            }
            else
            {
                return(true);
            }
        }

        /// <summary>
        /// Sendet ein Telegramm zu einer Schnittstelle und liest die Antwort ein.
        /// </summary>
        /// <param name="_Port">Schnittstelle.</param>
        /// <param name="_Data">Daten.</param>
        /// <param name="_Repeat">Anzahl der Wiederholungen bei Fehler.</param>
        /// <returns>Byte-Array mit Antwortdaten.</returns>
        private byte[] SendTel(SerialPort _Port, byte[] _Data, int _Repeat)
        {
            UInt16 usDataSize = (UInt16)_Data.Length;
            UInt16 usChecksum = 0, usRxChecksum = 0;
            List<byte> bytelist = new List<byte>();
            int b=-1;
            byte x;
            int state = 0;
            bool bFertig = false;
            bool bRetry = false;

            if ((_Port == null) || (usDataSize == 0)) return (null);

            byte[] sendbuf = new byte[usDataSize + 8];

            // Checksumme berechnen
            for (UInt16 i = 0; i < usDataSize; i++)
            {
                usChecksum += (UInt16)(_Data[i]);
            }

            // Telegramm zusammenbauen
            sendbuf[0] = NAV_TELSTART1;
            sendbuf[1] = NAV_TELSTART2;
            sendbuf[2] = (byte)(usDataSize >> 8);
            sendbuf[3] = (byte)usDataSize;
            _Data.CopyTo(sendbuf, 4);
            sendbuf[usDataSize + 4] = (byte)(usChecksum >> 8);
            sendbuf[usDataSize + 5] = (byte)usChecksum;
            sendbuf[usDataSize + 6] = NAV_TELEND1;
            sendbuf[usDataSize + 7] = NAV_TELEND2;

            // senden
            _Repeat++;
            while( _Repeat > 0 )
            {
              _Port.Write(sendbuf, 0, sendbuf.Length);
              Thread.Sleep(2 * sendbuf.Length);

              // Antwort empfangen
              state = 0;
              usChecksum = 0;
              bFertig = false;

              while (!bFertig)
              {
                try
                {
                  b = _Port.ReadByte();
                }
                catch (Exception)
                {
                  if (!bRetry)
                  {
                    bRetry = true;
                    Thread.Sleep(10);
                    continue;
                  }
                  else
                  {
                    b = -1;
                  }
                }

                if (b <= -1) break;

                x = (byte)b;

                switch (state)
                {
                  case 0:
                    if (x == NAV_TELSTART1)
                    {
                      state = 1;
                    }
                    break;

                  case 1:
                    if (x == NAV_TELSTART2)
                    {
                      state = 2;
                    }
                    break;

                  case 2:
                    // Länge Hi empfangen
                    usDataSize = x;
                    state = 3;
                    break;

                  case 3:
                    // Länge Lo empfangen
                    usDataSize <<= 8;
                    usDataSize |= x;
                    if (usDataSize == 0)
                    {
                      state = 5;
                    }
                    else
                    {
                      state = 4;
                    }
                    break;

                  case 4:
                    // Daten
                    bytelist.Add(x);
                    usChecksum += (UInt16)(x);
                    usDataSize--;
                    if (usDataSize == 0)
                    {
                      state = 5;
                    }
                    break;

                  case 5:
                    // Checksumme Hi empfangen
                    usRxChecksum = x;
                    state = 6;
                    break;

                  case 6:
                    // Checksumme Lo empfangen
                    usRxChecksum <<= 8;
                    usRxChecksum |= x;
                    state = 7;
                    break;

                  case 7:
                    if (x == NAV_TELEND1)
                    {
                      state = 8;
                    }
                    break;

                  case 8:
                    if (x == NAV_TELEND2)
                    {
                      state = 0;
                      bFertig = true;
                      _Repeat = 0;
                    }
                    break;

                  default:
                    state = 0;
                    break;
                }
              }

              if (bFertig && (usChecksum == usRxChecksum))
              {
                // alles OK
                return (bytelist.ToArray());
              }
              else
              {
                _Repeat--;
              }
            }

            return (null);
        }


        /// <summary>
        /// Sucht alle seriellen Schnittstellen des Rechners nach einen NavRun 500 ab.
        /// </summary>
        /// <returns>Gültiges Schnittstellen Objekt, wenn die Suche erfolgreich war, sonst null.</returns>
        private SerialPort OpenPort()
        {
            string[] comports = System.IO.Ports.SerialPort.GetPortNames();
            foreach(string i in comports)
            {
                Debug.WriteLine("NavRunCom: Versuche " + i);
                SerialPort port = null;
                try
                {
                    port = new SerialPort(i, 115200);
                    if( ValidNavRunPort(port) )
                    {
                        m_Port = port;
                        //m_Port.DataReceived += new System.IO.Ports.SerialDataReceivedEventHandler(ReceiveData);
                        return port;
                    }
                    else if (port != null)
                    {
                        Debug.WriteLine("Keine NavRun.");
                        port.Close();
                    }
                }
                catch (Exception)
                {
                    Debug.WriteLine("Fehler!");
                    if (port != null)
                    {
                        port.Close();
                    }
                }
            }
            return(null);
        }

    }
}
