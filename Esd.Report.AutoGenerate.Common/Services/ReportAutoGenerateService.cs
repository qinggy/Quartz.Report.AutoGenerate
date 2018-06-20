using Esd.Report.AutoGenerate.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Common.Service
{
    public class ReportAutoGenerateService : IService
    {
        public void AutoGenerate(JobDetail jobDetail)
        {
            CommHelper.AppLogger.Info($"当前Task的编号为{jobDetail.Id}");
        }
    }
}
