using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Common.Service
{
    public class CommonService : IService
    {
        public void Init()
        {
            CommHelper.AppLogger.Info("Init something ...");
        }
    }
}
