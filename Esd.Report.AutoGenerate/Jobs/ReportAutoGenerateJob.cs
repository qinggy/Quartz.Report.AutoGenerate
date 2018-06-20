using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Esd.Report.AutoGenerate.Common;
using Esd.Report.AutoGenerate.Common.Service;
using Newtonsoft.Json;
using Quartz;

namespace Esd.Report.AutoGenerate.Jobs
{
    public class ReportAutoGenerateJob : BaseJob
    {
        private readonly ReportAutoGenerateService reportAutoGenerateService;
        public ReportAutoGenerateJob(ReportAutoGenerateService _reportAutoGenerateService, CommonService _commonService)
            : base(_commonService)
        {
            reportAutoGenerateService = _reportAutoGenerateService;
        }

        public override void ExecuteJob(IJobExecutionContext context)
        {
            var dataMap = context.JobDetail.JobDataMap;
            var jsonJobEntity = dataMap.GetString("JobEntity");
            var jobEntity = JsonConvert.DeserializeObject<JobDetail>(jsonJobEntity);
            reportAutoGenerateService.AutoGenerate(jobEntity);
        }
    }
}
