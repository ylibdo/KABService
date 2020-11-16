using Microsoft.Extensions.Configuration;
using System;
using System.Net.Mail;
using System.Text;
using UtilityLibrary.Log;
using static UtilityLibrary.Log.LogObject;

namespace KABService.Helper
{
    class SMTPHelper
    {
        private readonly IConfiguration _configuration;

        public SMTPHelper(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public bool SendEmailAsync(string _emailBody)
        {
            LogHelper logHelper = new LogHelper(_configuration, "SMTP");
            bool sendEmailSuccess;
            try
            {
                string host = _configuration.GetValue<string>("SMTP:Host");
                SmtpClient client = new SmtpClient(host);
                string fromEmail = _configuration.GetValue<string>("SMTP:From");
                string displayName = _configuration.GetValue<string>("SMTP:DisplayName");
                string toEmail = _configuration.GetValue<string>("SMTP:To");
                MailAddress from = new MailAddress(fromEmail, displayName, Encoding.GetEncoding("iso-8859-1"));
                MailAddress to = new MailAddress(toEmail);
                MailMessage message = new MailMessage(from, to)
                {
                    Body = _emailBody
                };
                try
                {
                    string ccListString = _configuration.GetValue<string>("SMTP:CC");
                    string[] ccListArray = ccListString.Split(";");
                    foreach(string cc in ccListArray)
                    {
                        message.CC.Add(new MailAddress(cc.Trim()));
                    }
                }
                catch (Exception)
                {
                    // do nothing just ignore cc
                }
                message.BodyEncoding = Encoding.GetEncoding("iso-8859-1");
                message.Subject = _configuration.GetValue<string>("SMTP:Subject");
                message.SubjectEncoding = Encoding.GetEncoding("iso-8859-1");

                client.SendAsync(message, "KAB service notification");

                message.Dispose();
                sendEmailSuccess = true;
            }
            catch (Exception ex)
            {
                sendEmailSuccess = false;
                logHelper.InsertLog(new LogObject(LogType.Error, "Sending SMTP email failed. " + ex.Message));
            }

            return sendEmailSuccess;
        }
    }
}
