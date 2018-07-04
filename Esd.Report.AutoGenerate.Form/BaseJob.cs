using Esd.Report.AutoGenerate.Application;
using Esd.Report.AutoGenerate.Application.Service;
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
        private readonly CommonService CommonService;
        protected BaseJob()
        {
            CommonService = new CommonService();
        }

        public void Execute(IJobExecutionContext context)
        {
            if (CommonService.BeforeExecute(context))
                ExecuteJob(context);
            CommonService.AfterExecute(context);
        }

        public abstract void ExecuteJob(IJobExecutionContext context);
    }
}
