using Autofac;
using Autofac.Extras.Quartz;
using Esd.Report.AutoGenerate.Application;
using Esd.Report.AutoGenerate.Application.Model;
using Esd.Report.AutoGenerate.Jobs;
using Esd.Report.AutoGenerate.Service;
using Newtonsoft.Json;
using Quartz;
using Quartz.Impl;
using Quartz.Spi;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Topshelf;
using Topshelf.Quartz;
using Topshelf.ServiceConfigurators;

namespace Esd.Report.AutoGenerate
{
    public class ServiceRunner : ServiceControl, ServiceSuspend
    {
        static ServiceRunner()
        {
            ServiceName = AppSettings.ServiceName;
            DisplayServiceName = AppSettings.DisplayServiceName;
            ServiceDescription = AppSettings.ServiceDescription;
            databaseExpression = DatabaseFile.XmlToObject<DatabaseExpression>();

            //InitContainer();
            CommHelper.AppLogger.Info("初始化Job");
        }

        #region ServiceComponents
        private readonly IScheduler scheduler;

        public ServiceRunner()
        {
            var properties = new NameValueCollection
            {
                [AppSettings.poolType] = AppSettings.PoolType,
                [AppSettings.threadCount] = AppSettings.ThreadCount,
                [AppSettings.threadPriority] = AppSettings.ThreadPriority,
                [AppSettings.misfire] = AppSettings.MissFire,
                [AppSettings.storeType] = AppSettings.StoreType
            };
            if (AppSettings.IfServiceExport)
            {
                properties.Set(AppSettings.instanceName, AppSettings.InstanceName);
                properties.Set(AppSettings.exportType, AppSettings.ExportType);
                properties.Set(AppSettings.exportPort, AppSettings.ExportPort);
                properties.Set(AppSettings.bindName, AppSettings.BindName);
                properties.Set(AppSettings.channelType, AppSettings.ChannelType);
            }
            var schedulerFactory = new StdSchedulerFactory(properties);
            scheduler = schedulerFactory.GetScheduler();
            InitSchedule();
        }

        public bool Start(HostControl hostControl)
        {
            scheduler.Start();
            CommHelper.AppLogger.Info("服务启动");
            return true;
        }

        public bool Stop(HostControl hostControl)
        {
            Container.Dispose();
            scheduler.Clear();
            scheduler.Shutdown(false);
            CommHelper.AppLogger.Info("服务已关闭");
            return true;
        }

        public bool Continue(HostControl hostControl)
        {
            scheduler.ResumeAll();
            CommHelper.AppLogger.Info("重新启动");
            return true;
        }

        public bool Pause(HostControl hostControl)
        {
            scheduler.PauseAll();
            CommHelper.AppLogger.Info("服务暂停");
            return true;
        }

        #endregion

        private const string DatabaseFile = "database.config";
        public static readonly string ServiceName;
        public static readonly string DisplayServiceName;
        public static readonly string ServiceDescription;
        public static readonly DatabaseExpression databaseExpression;
        private static List<JobDetail> JobList
            = new List<JobDetail>();
        public static IContainer Container;

        private static void InitQuartzJob()
        {
            JobList = new Database().DbContext.Database.SqlQuery<JobDetail>(databaseExpression.QuerySqlJobWaitForExecute).ToList();
        }

        public static void InitSchedule(ServiceConfigurator<ServiceRunner> svc)
        {
            svc.UsingQuartzJobFactory(Container.Resolve<IJobFactory>);
            InitQuartzJob();

            foreach (var job in JobList)
            {
                svc.ScheduleQuartzJob(q =>
                {
                    q.WithJob(JobBuilder.Create<ReportAutoGenerateJob>()
                        .WithIdentity(job.Name, AppSettings.JobGroupName)
                        .UsingJobData("JobEntity", JsonConvert.SerializeObject(job))
                        .Build);

                    q.AddTrigger(() => TriggerBuilder.Create()
                        .WithIdentity($"{job.Name}_trigger", AppSettings.JobGroupName)
                        .WithCronSchedule(job.CornLike)
                        .Build());

                    CommHelper.AppLogger.InfoFormat("任务 {0} 已完成调度设置", job.Name);
                });
            }

            CommHelper.AppLogger.Info("调度任务 初始化完毕");
        }

        public void InitSchedule()
        {
            InitQuartzJob();
            foreach (var jobEntity in JobList)
            {
                var jobKey = JobKey.Create($"{jobEntity.Name}_{jobEntity.Id.ToString()}_job", AppSettings.JobGroupName);
                if (!scheduler.CheckExists(jobKey))
                {
                    var job = JobBuilder.Create<ReportAutoGenerateJob>()
                        .WithIdentity(jobKey)
                        .UsingJobData("JobEntity", JsonConvert.SerializeObject(jobEntity))
                        .Build();

                    var triggerKey = new TriggerKey($"{jobEntity.Name}_{jobEntity.Id.ToString()}_trigger", AppSettings.JobGroupName);
                    var trigger = TriggerBuilder.Create()
                        .WithIdentity(triggerKey)
                        .WithCronSchedule(jobEntity.CornLike)
                        .Build();

                    scheduler.ScheduleJob(job, trigger);
                    CommHelper.AppLogger.InfoFormat("任务 {0} 已完成调度设置", jobEntity.Name);
                }
            }

            CommHelper.AppLogger.Info("调度任务 初始化完毕");
        }

        private static void InitContainer()
        {
            var builder = new ContainerBuilder();
            builder.RegisterModule(new QuartzAutofacFactoryModule());
            builder.RegisterModule(new QuartzAutofacJobsModule(typeof(ServiceRunner).Assembly));
            builder.RegisterType<ServiceRunner>().AsSelf();

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
    }
}
