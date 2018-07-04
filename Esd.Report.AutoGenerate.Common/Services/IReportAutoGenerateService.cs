using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Services
{
    public interface IReportAutoGenerateService : IService
    {
        void AutoGenerate(JobDetail jobDetail);
    }
}
