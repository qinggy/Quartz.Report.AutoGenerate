using Esd.Report.AutoGenerate.Service;
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
        protected BaseJob(CommonService _commonService)
        {
            CommonService = _commonService;
        }

        public void Execute(IJobExecutionContext context)
        {
            CommonService.Init();

            ExecuteJob(context);
        }

        public abstract void ExecuteJob(IJobExecutionContext context);
    }
}
