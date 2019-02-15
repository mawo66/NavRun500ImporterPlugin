using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Reflection;

using ZoneFiveSoftware.Common.Visuals;
using ZoneFiveSoftware.Common.Visuals.Fitness;
using ZoneFiveSoftware.Common.Data;
using ZoneFiveSoftware.Common.Data.Fitness;
using ZoneFiveSoftware.Common.Data.GPS;

namespace NavRun500ImporterPlugin
{
    class NavRunDevice : IFitnessDevice
    {
        public NavRunDevice()
        {
            this.id = new Guid("21d2c9b0-4682-11e0-9207-0800200c9a66");
            this.image = Properties.Resources.Image_Uhr;
            this.name = "Ultrasport NavRun 500";
        }

        public string Configure(string configurationInfo)
        {
            return "";
        }

        public string ConfiguredDescription(string configurationInfo)
        {
            return Name;
        }

        public Guid Id
        {
            get { return id; }
        }

        public System.Drawing.Image Image
        {
            get { return image; }
        }

        public bool Import(string configurationInfo, IJobMonitor monitor, IImportResults importResults)
        {
            byte[] DevData = null;
            bool bRet = false;
            int iNumActivities=-1;
            bool bFound = false;
            string dllpath = Assembly.GetExecutingAssembly().Location;
            
            // Pfad der DLL ermitteln (zu Debugzwecken)
            try
            {
                dllpath = dllpath.Remove(dllpath.LastIndexOf('\\') + 1);
            }
            catch (Exception)
            {
                dllpath = "";
            }

            Debug.WriteLine("Import: configurationInfo = " + configurationInfo);

            NavRunCom device = new NavRunCom();
            monitor.PercentComplete = 0;


            if (File.Exists(dllpath + "navruntestdata.bin"))
            {
                // zu Debugzwecken aus Testdatei importieren
                monitor.StatusText = Properties.Resources.Status_ImportFromFile;
                BinaryReader binReader;
                FileInfo fileinfo = new FileInfo(dllpath + "navruntestdata.bin");
                long filesize = fileinfo.Length;

                binReader = new BinaryReader(File.Open(dllpath + "navruntestdata.bin", FileMode.Open));
                DevData = binReader.ReadBytes((int)filesize);
                binReader.Close();
                monitor.PercentComplete = 1;
                bFound = true;
            }
            else
            {
                // Gerät suchen und Daten lesen
                monitor.StatusText = Properties.Resources.Status_SucheGeraet;
                if (device.Open())
                {
                    monitor.StatusText = Properties.Resources.Status_LeseDaten;
                    DevData = device.ReadData(monitor);
                    bFound = true;
                }
            }

            if (bFound)
            {
                if ( (DevData == null) || (DevData.Length < 0x1000) )
                {
                    // Fehler beim Auslesen
                    monitor.ErrorText = Properties.Resources.Error_LeseDaten;
                }
                else
                {
                    try
                    {
                        BinaryWriter binWriter;
                        binWriter = new BinaryWriter(File.Open(dllpath + "rawdata.bin", FileMode.Create));
                        binWriter.Write(DevData);
                        binWriter.Close();
                    }
                    catch (Exception)
                    { 
                    }

                    NavRunInterpreter interpreter = new NavRunInterpreter();
                    iNumActivities = interpreter.GetNumberOfActivities(ref DevData);
                    if (iNumActivities == 0)
                    {
                        // Auslesen OK, aber keine Daten da
                        monitor.PercentComplete = 1;
                        bRet = true;
                    }
                    else
                    {
                        // Auslesen OK und Daten vorhanden
                        for (int i = 0; i < iNumActivities; i++)
                        {
                            interpreter.ImportActivity(ref DevData, i, importResults);
                        }
                        monitor.PercentComplete = 1;
                        bRet = true;
                    }
                }
                device.Close();
            }
            else
            {
                monitor.ErrorText = Properties.Resources.Error_NichtGefunden;
            }

            return bRet;
        }

        public string Name
        {
            get { return name; }
        }

        #region Private members
        private Guid id;
        private Image image;
        private string name;
        #endregion
    }
}
