using Autofac;
using Esd.Report.AutoGenerate.Service;
using log4net;
using System;
using System.IO;
using System.Reflection;
using Topshelf;
using Topshelf.Autofac;

namespace Esd.Report.AutoGenerate
{
    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(config =>
            {
                config.UseLog4Net();
                config.Service<ServiceRunner>();
                config.RunAsLocalSystem();
                config.UseAutofacContainer(JobService.Container);
                config.EnablePauseAndContinue();
                config.SetServiceName(JobService.ServiceName);
                config.SetDisplayName(JobService.DisplayServiceName);
                config.SetDescription(JobService.ServiceDescription);

                config.Service<JobService>(setting =>
                {
                    JobService.InitSchedule(setting);
                    setting.ConstructUsingAutofacContainer();
                    setting.WhenStarted(o => o.Start());
                    setting.WhenStopped(o => o.Stop());
                });
            });
        }
    }
}
