using System;
using System.Configuration;
using System.IO;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;

namespace ReportManager
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Report report = new Report();
                report.SendEmail();
            }
            catch (Exception ex)
            {
                Console.Write($" Mail sending falied, exception : {ex.Message}");
            }
        }

    }

    public class Report
    {

        public void SendEmail()
        {
            string s = ConfigurationManager.AppSettings["To"];
            string[] toName = s.Split('@');
            string username = toName[0].ToUpper() ?? "Dude";
            string subject = "A new report has been published on " + DateTime.Now.ToShortDateString();
            string title = "Sales Report for Year 2016";
            string url = "Sales Report for Year 2016";
            string description = "Please find Sales Report  " + " for year 2016-17, " + "hope you will come up with  proper actions as soon as possible";
            string reportName = "Sales Report for Year 2016";
            string body = this.PopulateBody(username, subject, title, url, description, reportName);

            this.SendHtmlFormattedEmail(ConfigurationManager.AppSettings["To"], subject, body);
        }

        private string PopulateBody(string userName, string subject, string title, string url, string description, string reportName)
        {
            string body = string.Empty;
            using (StreamReader reader = new StreamReader("C:\\Project2017\\OfflineReport\\ReportManager\\content\\content.html"))
            {
                body = reader.ReadToEnd();
            }
            body = body.Replace("{UserName}", userName);
            body = body.Replace("{Subject}", subject);
            body = body.Replace("{Title}", title);
            body = body.Replace("{Url}", url);
            body = body.Replace("{Description}", description);
            body = body.Replace("{ReportName}", reportName);

            return body;
        }
       
        private void SendHtmlFormattedEmail(string recepientEmail, string subject, string body)
        {
            using (MailMessage mailMessage = new MailMessage())
            {
               
                PopulateTemplate("Sales Report 2016", "Sales Report 2016");
                string file = "C:\\Project2017\\OfflineReport\\ReportManager\\temp\\template.html";
                mailMessage.From = new MailAddress(ConfigurationManager.AppSettings["UserName"]);
                mailMessage.Subject = subject;
                mailMessage.Body = body;
                mailMessage.IsBodyHtml = true;
                // Create  the file attachment for this e-mail message.
                Attachment data = new Attachment(file, MediaTypeNames.Application.Octet);
                mailMessage.Attachments.Add(data);
                mailMessage.To.Add(new MailAddress(recepientEmail));
                SmtpClient smtp = new SmtpClient();
                smtp.Host = ConfigurationManager.AppSettings["Host"];
                smtp.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings["EnableSsl"]);
                System.Net.NetworkCredential NetworkCred = new System.Net.NetworkCredential();
                NetworkCred.UserName = ConfigurationManager.AppSettings["UserName"];
                NetworkCred.Password = ConfigurationManager.AppSettings["Password"];
                smtp.UseDefaultCredentials = true;
                smtp.Credentials = NetworkCred;
                smtp.Port = int.Parse(ConfigurationManager.AppSettings["Port"]);

                smtp.Send(mailMessage);

            }
        }

        void PopulateTemplate(string title, string reportName)
        {
            string template = string.Empty;

            const string fileName = "C:\\Project2017\\OfflineReport\\ReportManager\\templates\\template.html";
            const string outFileName = "C:\\Project2017\\OfflineReport\\ReportManager\\temp\\template.html";

            //Read HTML from file
            template = File.ReadAllText(fileName);
            File.ReadAllText(outFileName);

            //Replace all values in the HTML
            template = template.Replace("{Title}", title ?? string.Empty);
            template = template.Replace("{ReportName}", reportName ?? string.Empty);

            //Write new HTML string to file
            File.WriteAllText(outFileName, template);

        }

    }
}


