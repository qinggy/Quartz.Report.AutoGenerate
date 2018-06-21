using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Common.Service
{
    public class CommonService : IService
    {
        public void BeforeExecute()
        {
            CommHelper.AppLogger.Info("Start execute a job ...");
        }

        public void AfterExecute()
        {
            CommHelper.AppLogger.Info("End execute a job ...");
        }
    }
}
