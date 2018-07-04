using Esd.Report.AutoGenerate.Application;
using Esd.Report.AutoGenerate.Application.Service;
using Esd.Report.AutoGenerate.Application.Services;
using Newtonsoft.Json;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate
{
    public abstract class BaseJob : IJob
    {
        private readonly ICommonService commonService;
        protected BaseJob()
        {
            commonService = new CommonService();
        }

        public void Execute(IJobExecutionContext context)
        {
            if (commonService.BeforeExecute(context))
                ExecuteJob(context);
            commonService.AfterExecute(context);
        }

        public abstract void ExecuteJob(IJobExecutionContext context);
    }
}
