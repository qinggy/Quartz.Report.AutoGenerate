using Esd.Report.AutoGenerate.Application.Services;
using Newtonsoft.Json;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Service
{
    public class CommonService : ICommonService
    {
        public bool BeforeExecute(IJobExecutionContext context)
        {
            var dataMap = context.JobDetail.JobDataMap;
            var jsonJobEntity = dataMap.GetString("JobEntity");
            var jobEntity = JsonConvert.DeserializeObject<JobDetail>(jsonJobEntity);
            CommHelper.AppLogger.Info($"Start execute {jobEntity.Name} job ...");

            return true;
        }

        public bool AfterExecute(IJobExecutionContext context)
        {
            var dataMap = context.JobDetail.JobDataMap;
            var jsonJobEntity = dataMap.GetString("JobEntity");
            var jobEntity = JsonConvert.DeserializeObject<JobDetail>(jsonJobEntity);
            CommHelper.AppLogger.Info($"End execute {jobEntity.Name} job ...");

            return true;
        }
    }
}
