using System;
using System.Collections.Generic;
using System.Text;
using ZoneFiveSoftware.Common.Visuals.Fitness;

namespace NavRun500ImporterPlugin
{
    class ExtendFitnessDevices : IExtendFitnessDevices
    {
        public IList<IFitnessDevice> FitnessDevices
        {
            get { return new IFitnessDevice[] { new NavRunDevice() }; }
        }
    }
}
