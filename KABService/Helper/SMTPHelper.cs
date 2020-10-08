using Microsoft.Extensions.Configuration;
using System;
using System.Net.Mail;
using System.Text;

namespace KABService.Helper
{
    class SMTPHelper
    {
        private readonly IConfiguration _configuration;

        public SMTPHelper(IConfiguration configuration)
        {
            _configuration = configuration;
        }
        public bool SendEmailAsync(string _emailBody) // argument must be the SMTP host.
        {
            bool sendEmailSuccess;
            try
            {
                string host = _configuration.GetValue<string>("SMTP:Host");
                SmtpClient client = new SmtpClient(host);
                string fromEmail = _configuration.GetValue<string>("SMTP:From");
                string displayName = _configuration.GetValue<string>("SMTP:DisplayName");
                string toEmail = _configuration.GetValue<string>("SMTP:Tos");
                MailAddress from = new MailAddress(fromEmail, displayName, Encoding.UTF8);
                MailAddress to = new MailAddress(toEmail);
                MailMessage message = new MailMessage(from, to)
                {
                    Body = _emailBody
                };
                message.BodyEncoding = Encoding.UTF8;
                message.Subject = _configuration.GetValue<string>("SMTP:Subject");
                message.SubjectEncoding = Encoding.UTF8;

                client.SendAsync(message, "KAB service notification");

                message.Dispose();
                sendEmailSuccess = true;
            }
            catch (Exception)
            {
                sendEmailSuccess = false;
            }

            return sendEmailSuccess;
        }
    }
}
