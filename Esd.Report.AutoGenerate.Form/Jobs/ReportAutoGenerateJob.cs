using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Esd.Report.AutoGenerate.Application;
using Esd.Report.AutoGenerate.Application.Service;
using Esd.Report.AutoGenerate.Application.Services;
using Newtonsoft.Json;
using Quartz;

namespace Esd.Report.AutoGenerate.Jobs
{
    public class ReportAutoGenerateJob : BaseJob
    {
        private readonly IReportAutoGenerateService reportAutoGenerateService;

        public ReportAutoGenerateJob()
        {
            reportAutoGenerateService = new ReportAutoGenerateService();
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
