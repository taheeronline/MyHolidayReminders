using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace MyHolidayReminders
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        static void Main()
        {
            Service1 service = new Service1();
            service.CanHandleSessionChangeEvent = true;
            service.CanPauseAndContinue = true;
            service.CanShutdown = true;

            ServiceBase.Run(service);
        }
    }
}
