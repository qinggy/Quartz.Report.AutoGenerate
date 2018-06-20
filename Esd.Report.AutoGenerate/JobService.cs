using Autofac;
using Autofac.Extras.Quartz;
using Esd.Report.AutoGenerate.Jobs;
using Esd.Report.AutoGenerate.Common.Model;
using Newtonsoft.Json;
using Quartz;
using Quartz.Spi;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Topshelf.Quartz;
using Topshelf.ServiceConfigurators;
using Esd.Report.AutoGenerate.Common;

namespace Esd.Report.AutoGenerate
{
    public class JobService
    {
        private const string DatabaseFile = "database.config";
        public static readonly string ServiceName;
        public static readonly string DisplayServiceName;
        public static readonly string ServiceDescription;
        public static readonly DatabaseExpression databaseExpression;
        private static List<JobDetail> JobList
            = new List<JobDetail>();
        public static IContainer Container;

        static JobService()
        {
            ServiceName = ConfigurationManager.AppSettings["ServiceName"];
            DisplayServiceName = ConfigurationManager.AppSettings["DisplayServiceName"];
            ServiceDescription = ConfigurationManager.AppSettings["ServiceDescription"];
            databaseExpression = DatabaseFile.XmlToObject<DatabaseExpression>();

            InitContainer();
            CommHelper.AppLogger.Info("初始化Job");
        }

        private static void InitQuartzJob()
        {
            JobList = new Database().DbContext.Database.SqlQuery<JobDetail>(databaseExpression.QuerySqlJobWaitForExecute).ToList();
        }

        public static void InitSchedule(ServiceConfigurator<JobService> svc)
        {
            svc.UsingQuartzJobFactory(Container.Resolve<IJobFactory>);
            InitQuartzJob();

            foreach (var job in JobList)
            {
                svc.ScheduleQuartzJob(q =>
                {
                    q.WithJob(JobBuilder.Create<ReportAutoGenerateJob>()
                        .WithIdentity(job.Name, ServiceName)
                        .UsingJobData("JobEntity", JsonConvert.SerializeObject(job))
                        .Build);

                    q.AddTrigger(() => TriggerBuilder.Create()
                        .WithCronSchedule(job.CornLike)
                        .Build());

                    CommHelper.AppLogger.InfoFormat("任务 {0} 已完成调度设置", job.Name);
                });
            }

            CommHelper.AppLogger.Info("调度任务 初始化完毕");
        }


        private static void InitContainer()
        {
            var builder = new ContainerBuilder();
            builder.RegisterModule(new QuartzAutofacFactoryModule());
            builder.RegisterModule(new QuartzAutofacJobsModule(typeof(JobService).Assembly));
            builder.RegisterType<JobService>().AsSelf();

            var execDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            var files = Directory.GetFiles(execDir, "Esd.Report.AutoGenerate.*.dll", SearchOption.TopDirectoryOnly);
            if (files.Length > 0)
            {
                var assemblies = new Assembly[files.Length];
                for (var i = 0; i < files.Length; i++)
                    assemblies[i] = Assembly.LoadFile(files[i]);

                builder.RegisterAssemblyTypes(assemblies)
                    .Where(t => t.GetInterfaces().ToList().Contains(typeof(IService)))
                    .AsSelf()
                    .InstancePerLifetimeScope();
            }

            Container = builder.Build();
            CommHelper.AppLogger.Info("初始化完成");
        }

        public bool Start()
        {
            CommHelper.AppLogger.Info("服务启动");
            return true;
        }

        public bool Stop()
        {
            Container.Dispose();
            CommHelper.AppLogger.Info("服务已关闭");
            return false;
        }

    }
}
