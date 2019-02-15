using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;

using ZoneFiveSoftware.Common.Visuals.Fitness;


namespace NavRun500ImporterPlugin
{
    class Plugin : IPlugin
    {
        public Plugin()
        {
            instance = this;
        }

        #region IPlugin Members

        public Guid Id
        {
            get { return new Guid("ceb3610c-6caf-46ec-b6bb-7c857ca075af"); }
        }

        public IApplication Application
        {
            get { return application; }
            set { application = value; }
        }

        public string Name
        {
            get { return "NavRun500 Importer Plugin"; }
        }

        public string Version
        {
            get { return GetType().Assembly.GetName().Version.ToString(3); }
        }

        public void ReadOptions(XmlDocument xmlDoc, XmlNamespaceManager nsmgr, XmlElement pluginNode)
        {
        }

        public void WriteOptions(XmlDocument xmlDoc, XmlElement pluginNode)
        {
        }

        #endregion

        public static Plugin Instance
        {
            get { return instance; }
        }

        private static Plugin instance = null;

        private IApplication application;
    }
}
