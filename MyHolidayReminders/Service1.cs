using System.ServiceProcess;
using System.Timers;
using System.IO;
using System;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.Net;

namespace MyHolidayReminders
{
    public partial class Service1 : ServiceBase
    {
        private System.Timers.Timer myTimer =new System.Timers.Timer();
        private string connectionString = "Data Source=IE-0024\\SQLEXPRESS;Initial Catalog=NTINTReminders;User ID=sa;Password=Ntint123;"; // Update with your database connection details

        // Gmail account details for sending email
        private string gmailUsername = "taheeronline@gmail.com"; // Replace with your Gmail email address
        private string gmailPassword = "uxqjzzimesgwzpay"; // Replace with your Gmail password or app password



        public Service1()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {

            WriteToFile("Service is started at " + DateTime.Now);
            myTimer.Elapsed += new ElapsedEventHandler(OnElapsedTime);
            myTimer.Interval = TimeSpan.FromHours(24).TotalMilliseconds; //number in miliseconds
            myTimer.Enabled = true;

        }

        protected override void OnStop()
        {

            WriteToFile("Service is stopped at " + DateTime.Now);

        }

        private DataTable GetHolidayReminders()
        {
            WriteToFile("Service is reading reminders " + DateTime.Now);
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
                WriteToFile("Service sending email " + DateTime.Now);
                foreach (DataRow row in reminders.Rows)
                {
                    DateTime holidayDate = Convert.ToDateTime(row["Date"]);
                    string holidayName = row["HolidayName"].ToString();

                    // Check if the holiday reminder date matches the current date
                    
                    if (holidayDate.Date.AddDays(-2) == DateTime.Today)
                    {
                        // Send email alert
                        SendEmailNotification(holidayDate, holidayName);
                    }
                }
                WriteToFile("Service finished sending email " + DateTime.Now);
            }
            catch (Exception ex)
            {
                
                WriteToFile("Error occured While generating reminders. " + ex.Message);
            }
        }

        private void SendEmailNotification(DateTime holidayDate,string holidayName)
        {
            try
            {
                
                using (SmtpClient smtpClient = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtpClient.EnableSsl = true;
                    smtpClient.UseDefaultCredentials = false;
                    smtpClient.Credentials = new NetworkCredential(gmailUsername, gmailPassword);

                    using (MailMessage mailMessage = new MailMessage())
                    {
                        mailMessage.From = new MailAddress(gmailUsername);
                        mailMessage.To.Add("amir.keynia@ntint.com");
                        mailMessage.To.Add("barry.prabhu@ntint.com");
                        mailMessage.CC.Add("shamshuddin.wantmu@ntint.com");
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
