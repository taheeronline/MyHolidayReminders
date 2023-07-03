using System.ServiceProcess;
using System.Timers;
using System.IO;
using System;

namespace MyHolidayReminders
{
    public partial class Service1 : ServiceBase
    {
        Timer myTimer=new Timer();
        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service is started at " + DateTime.Now);
            myTimer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            myTimer.Interval = 5000; //number in miliseconds
            myTimer.Enabled = true;

        }

        protected override void OnStop()
        {
            WriteToFile("Service is stopped at " + DateTime.Now);
        }

        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Service was ran at " + DateTime.Now);

        }

        public void WriteToFile(string myMessage)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\logs";

            if(!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\logs\\ServiceLog_" + DateTime.Now.Date.ToShortDateString().Replace('/','_') +".txt" ;
            if (!File.Exists(filePath))
            {
                //create a file to write to.
                using(StreamWriter sw=File.CreateText(filePath))
                {
                    sw.WriteLine(myMessage);
                }
            }
            else
            {
                using (StreamWriter sw =File.AppendText(filePath))
                {
                    sw.WriteLine(myMessage);
                }
            }
        }
    }
}
