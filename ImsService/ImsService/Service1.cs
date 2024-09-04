using System;
using System.Data.SqlClient;
using System.IO;
using System.Net.Mail;
using System.ServiceProcess;
using System.Timers;

namespace ImsService
{
    public partial class Service1 : ServiceBase
    {
        private Timer timer;
        private const string ConnectionString = "your_connection_string_here"; // Replace with your connection string
        private const string SmtpServer = "smtp.office365.com";
        private const int SmtpPort = 587;
        private const string SenderEmail = "your_email@domain.com"; // Replace with your email
        private const string SenderPassword = "your_password"; // Replace with your password

        public Service1()
        {
            InitializeComponent();
            timer = new Timer();
            timer.Interval = GetNextInterval();
            timer.Elapsed += OnTimer;
        }

        protected override void OnStart(string[] args)
        {
            WriteToFile("Service started at " + DateTime.Now + " - HTVVMS Balance Confirmation email service");
            timer.AutoReset = true;
            timer.Start();
        }

        protected override void OnStop()
        {
            WriteToFile("Service stopped at " + DateTime.Now + " - HTVVMS Balance Confirmation email service");
            timer.AutoReset = false;
            timer.Stop();
        }
                
        private double GetNextInterval()
        {
            string timeString = "3:00 PM";
            DateTime targetTime = DateTime.Parse(timeString, System.Globalization.CultureInfo.InvariantCulture);
            TimeSpan timeDifference = targetTime - DateTime.Now;

            if (timeDifference.TotalMilliseconds < 0)
            {
                timeDifference = targetTime.AddDays(1) - DateTime.Now;
            }

            return timeDifference.TotalMilliseconds;
        }


        private void OnTimer(object sender, ElapsedEventArgs e)
        {
            CheckForPendingReturnsAndSendEmail();
            timer.Stop();
            System.Threading.Thread.Sleep(1000000); // Adjust sleep time as necessary
            timer.Interval = GetNextInterval();
            timer.Start();
        }

        private void CheckForPendingReturnsAndSendEmail()
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(ConnectionString))
                {
                    connection.Open();
                    WriteToFile("Database connection opened at " + DateTime.Now);

                    string tomorrowDate = DateTime.Now.AddDays(1).ToString("yyyy-MM-dd");
                    string query = "SELECT VendorEmail, ExpectedReturnDate FROM RepairOrders WHERE ExpectedReturnDate = @tomorrowDate";

                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        command.Parameters.AddWithValue("@tomorrowDate", tomorrowDate);

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                string vendorEmail = reader["VendorEmail"].ToString();
                                DateTime expectedReturnDate = Convert.ToDateTime(reader["ExpectedReturnDate"]);
                                SendEmail(vendorEmail, expectedReturnDate);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteToFile("Error in CheckForPendingReturnsAndSendEmail: " + ex.Message);
            }
        }

        private void SendEmail(string vendorEmail, DateTime expectedReturnDate)
        {
            try
            {
                MailMessage mail = new MailMessage
                {
                    From = new MailAddress(SenderEmail),
                    Subject = "Reminder: Material Return Due Tomorrow",
                    Body = $"Dear Vendor,\n\nThis is a reminder that you promised to return the materials on {expectedReturnDate.ToShortDateString()}.\n\nRegards,\nYour Company",
                    IsBodyHtml = false
                };
                mail.To.Add(vendorEmail);

                SmtpClient smtpClient = new SmtpClient(SmtpServer)
                {
                    Port = SmtpPort,
                    Credentials = new System.Net.NetworkCredential(SenderEmail, SenderPassword),
                    EnableSsl = true
                };

                smtpClient.Send(mail);
                WriteToFile("Email sent to " + vendorEmail + " at " + DateTime.Now);
            }
            catch (Exception ex)
            {
                WriteToFile("Error while sending email: " + ex.Message);
            }
        }

        private void WriteToFile(string message)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory + "\\Logs";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string filepath = Path.Combine(path, "ServiceLog_" + DateTime.Now.ToString("yyyy-MM-dd") + ".txt");
            using (StreamWriter sw = File.AppendText(filepath))
            {
                sw.WriteLine(message);
            }
        }
    }
}
