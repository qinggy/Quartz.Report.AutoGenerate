using Newtonsoft.Json;
using Quartz;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Esd.Report.AutoGenerate.Application.Service
{
    public class CommonService : IService
    {
        public bool BeforeExecute(IJobExecutionContext context)
        {
            //var dataMap = context.JobDetail.JobDataMap;
            //var jsonJobEntity = dataMap.GetString("JobEntity");
            //var jobEntity = JsonConvert.DeserializeObject<JobDetail>(jsonJobEntity);
            CommHelper.AppLogger.Info($"Start execute job ...");

            return true;
        }

        public bool AfterExecute(IJobExecutionContext context)
        {
            //var dataMap = context.JobDetail.JobDataMap;
            //var jsonJobEntity = dataMap.GetString("JobEntity");
            //var jobEntity = JsonConvert.DeserializeObject<JobDetail>(jsonJobEntity);
            CommHelper.AppLogger.Info($"End execute job ...");

            return true;
        }
    }
}
