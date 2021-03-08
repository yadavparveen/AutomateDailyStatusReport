using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using OpenXmlPowerTools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Xml.Linq;

namespace AutomateDailyStatusReport
{
    class Program
    {
        static void Main(string[] args)
        {
            var todayDate = DateTime.Now.ToString("dddd, dd MMMM yyyy");

            var statusReports = LoadStatusReportConfig();

            foreach(var statusReport in statusReports)
            {
                var htmlBody = ConverWordToHtml(statusReport.StatusReportFilePath);
                SendEmail(htmlBody, $"{statusReport.EmailSubject} {todayDate}");
            }
        }

        private static void SendEmail(string htmlString, string subject)
        {
            try
            {
                MailMessage msg = new MailMessage();

                msg.From = new MailAddress("from@email.com");
                msg.To.Add("to@email.com");
                msg.Subject = subject;
                msg.Body = htmlString;
                msg.IsBodyHtml = true;

                using (SmtpClient client = new SmtpClient())
                {
                    client.EnableSsl = true;
                    client.UseDefaultCredentials = false;
                    client.Credentials = new NetworkCredential("userId@gmail.com", "your-password-here");
                    client.Host = "smtp.gmail.com";
                    client.Port = 587;
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;

                    client.Send(msg);
                }
            }
            catch (Exception ex) 
            {
                throw ex;
            }
        }

        private static string ConverWordToHtml(string wordFilePath)
        {
            byte[] byteArray = File.ReadAllBytes(wordFilePath);

            using (MemoryStream memoryStream = new MemoryStream())
            {
                memoryStream.Write(byteArray, 0, byteArray.Length);
                using (WordprocessingDocument doc = WordprocessingDocument.Open(memoryStream, true))
                {
                    HtmlConverterSettings settings = new HtmlConverterSettings()
                    {
                        PageTitle = "Daily Status"
                    };

                    XElement html = HtmlConverter.ConvertToHtml(doc, settings);

                    return html.ToStringNewLineOnAttributes();
                }
            }
        }

        private static List<StatusReportConfig> LoadStatusReportConfig()
        {
            try
            {
                var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
                var userInfoFilePath = Path.Combine(baseDirectory, "StatusReportConfig.json");

                using (StreamReader r = new StreamReader(userInfoFilePath))
                {
                    return JsonConvert.DeserializeObject<List<StatusReportConfig>>(r.ReadToEnd());
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }
        }
    }
}
