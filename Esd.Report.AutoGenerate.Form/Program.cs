using Autofac;
using Esd.Report.AutoGenerate.Application;
using log4net;
using System;
using System.IO;
using System.Reflection;
using Topshelf;
using Topshelf.Autofac;
using Topshelf.ServiceConfigurators;

namespace Esd.Report.AutoGenerate
{
    class Program
    {
        static void Main(string[] args)
        {
            HostFactory.Run(config =>
            {
                config.Service<ServiceRunner>(s =>
                {
                    //ServiceRunner.InitSchedule(s);
                    //s.ConstructUsingAutofacContainer();
                    s.ConstructUsing(() => new ServiceRunner());
                    s.WhenStarted((sr, hct) => sr.Start(hct));
                    s.WhenStopped((sr, hct) => sr.Stop(hct));
                    s.WhenContinued((sr, hct) => sr.Continue(hct));
                    s.WhenPaused((sr, hct) => sr.Pause(hct));
                });

                config.RunAsLocalSystem();
                config.StartAutomatically();
                //config.UseAutofacContainer(ServiceRunner.Container);
                config.SetServiceName(ServiceRunner.ServiceName);
                config.SetDisplayName(ServiceRunner.DisplayServiceName);
                config.SetDescription(ServiceRunner.ServiceDescription);
                config.EnablePauseAndContinue();
                config.EnableServiceRecovery(action =>
                {
                    action.RestartService(1);
                    action.OnCrashOnly();

                    action.RestartService(3);
                    action.OnCrashOnly();

                    action.RestartService(5);
                    action.OnCrashOnly();
                });
                config.UseLog4Net();
            });

            Console.ReadLine();
        }
    }
}
