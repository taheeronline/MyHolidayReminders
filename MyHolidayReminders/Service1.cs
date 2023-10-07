using System.ServiceProcess;
using System.Timers;
using System.IO;
using System;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Net;
using System.Configuration;

namespace MyHolidayReminders
{
    public partial class Service1 : ServiceBase
    {
        private System.Timers.Timer myTimer = new System.Timers.Timer();
        private string connectionString;
        private string email;
        private string password;

        public Service1()
        {
            InitializeComponent();

        }

        protected override void OnStart(string[] args)
        {

            WriteToFile("Service is started at " + DateTime.Now);

            // Calculate the time until the next desired execution time (e.g., 8:00 AM)
            DateTime now = DateTime.Now;
            DateTime nextExecutionTime = new DateTime(now.Year, now.Month, now.Day, 17, 00, 0); // Set to 05:00 PM
            if (now >= nextExecutionTime)
            {
                nextExecutionTime = nextExecutionTime.AddDays(1);
            }
            TimeSpan timeUntilNextExecution = nextExecutionTime - now;

            // Set the timer interval to fire daily at the desired execution time
            myTimer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            myTimer.Interval = timeUntilNextExecution.TotalMilliseconds;//60000;
            myTimer.AutoReset = false;
            myTimer.Start();

        }

        protected override void OnStop()
        {

            WriteToFile("Service is stopped at " + DateTime.Now);

        }

        private DataTable GetHolidayReminders()
        {
            WriteToFile("Service is reading reminders " + DateTime.Now);

            ExeConfigurationFileMap configFileMap = new ExeConfigurationFileMap();
            configFileMap.ExeConfigFilename = "C:\\Users\\Novatek\\source\\repos\\MyHolidayReminders\\MyHolidayReminders\\App.config";
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configFileMap, ConfigurationUserLevel.None);

            //ConfigurationManager.RefreshSection("appSettings");

            connectionString = config.AppSettings.Settings["ConnectionString"].Value;


            email = config.AppSettings.Settings["Email"].Value;
            password = config.AppSettings.Settings["Password"].Value;


            DataTable reminders = new DataTable();

            try
            {
                using (SqlConnection connection = new SqlConnection(connectionString))
                {
                    connection.Open();

                    string query = "SELECT Date, HolidayName FROM HolidayReminders";
                    SqlCommand command = new SqlCommand(query, connection);

                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(reminders);
                    WriteToFile("Service completed reading reminders " + DateTime.Now);
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur while retrieving holiday reminders
                WriteToFile("Error retrieving holiday reminders: " + ex.Message);
            }

            return reminders;
        }
        private void OnElapsedTime(object source, ElapsedEventArgs e)
        {
            WriteToFile("Service was ran at " + DateTime.Now);

            try
            {

                // Retrieve holiday reminders from the database
                DataTable reminders = GetHolidayReminders();
                foreach (DataRow row in reminders.Rows)
                {
                    DateTime holidayDate = Convert.ToDateTime(row["Date"]);
                    string holidayName = row["HolidayName"].ToString();

                    // Check if the holiday reminder date matches the current date

                    if (holidayDate.Date.AddDays(-3) == DateTime.Today)
                    {
                        // Send email alert
                        WriteToFile("Service sending email " + DateTime.Now);
                        SendEmailNotification(holidayDate, holidayName);
                        WriteToFile("Service finished sending email for " + holidayName + " " + DateTime.Now);
                    }
                }
                // Calculate the time until the next desired execution time for the next day
                DateTime now = DateTime.Now;
                DateTime nextExecutionTime = new DateTime(now.Year, now.Month, now.Day, 17, 00, 0).AddDays(1); // Set to 05:00 PM
                TimeSpan timeUntilNextExecution = nextExecutionTime - now;

                // Reset the timer for the next day
                myTimer.Interval = timeUntilNextExecution.TotalMilliseconds;
                myTimer.Start();
            }
            catch (Exception ex)
            {

                WriteToFile("Error occured While generating reminders. " + ex.Message);
            }
        }

        private void SendEmailNotification(DateTime holidayDate, string holidayName)
        {
            try
            {

                using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtpClient.EnableSsl = true;
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new NetworkCredential(email, password);

                    using (MailMessage mailMessage = new MailMessage())
                    {
                        mailMessage.From = new MailAddress(email);
                        mailMessage.To.Add("amir.keynia@ntint.com");
                        mailMessage.To.Add("barry.prabhu@ntint.com");
                        mailMessage.Bcc.Add("shamshuddin.wantmu@ntint.com");
                        mailMessage.Bcc.Add("anish.kuriakose@ntint.com");
                        mailMessage.Subject = "NTINT India Holiday Reminder (Generated from SHAMS Windows Service)";
                        mailMessage.Body = $"{holidayName} is on {holidayDate}. India team will be off on this date.";

                        smtpClient.Send(mailMessage);
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that occur during email sending
                WriteToFile("Failed to send email notification: " + ex.Message);
            }
        }
        public void WriteToFile(string myMessage)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\logs";

            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }

            string filePath = AppDomain.CurrentDomain.BaseDirectory + "\\logs\\ReminderLog_" + DateTime.Now.Date.ToShortDateString().Replace('/', '_') + ".txt";
            if (!File.Exists(filePath))
            {
                //create a file to write to.
                using (StreamWriter sw = File.CreateText(filePath))
                {
                    sw.WriteLine(myMessage);
                }
            }
            else
            {
                using (StreamWriter sw = File.AppendText(filePath))
                {
                    sw.WriteLine(myMessage);
                }
            }
        }
    }
}
