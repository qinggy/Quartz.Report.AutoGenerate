using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Services
{
    public interface ICommonService : IService
    {
        bool BeforeExecute(IJobExecutionContext context);

        bool AfterExecute(IJobExecutionContext context);
    }
}
