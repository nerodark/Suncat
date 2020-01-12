using System;
using System.ServiceProcess;
using System.Threading;

namespace SuncatService
{
    public static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        public static void Main(string[] args)
        {
            if (Environment.UserInteractive)
            {
                var service = new SuncatService();
                service.OnDebug(args);
            }
            else
            {
                ServiceBase[] ServicesToRun;
                ServicesToRun = new ServiceBase[]
                {
                    new SuncatService()
                };
                ServiceBase.Run(ServicesToRun);
            }
        }
    }
}
