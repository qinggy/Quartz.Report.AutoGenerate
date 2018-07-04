using Esd.Report.AutoGenerate.Application;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Service
{
    public class ReportAutoGenerateService : IService
    {
        public void AutoGenerate(JobDetail jobDetail)
        {
            CommHelper.AppLogger.Info($"当前Task的编号为{jobDetail.Id}");
        }
    }
}
