using System;
using System.Configuration;
using System.Net.Mail;

namespace TarantusSAPService
{
    public static class Email
    {
        /// <summary>
        /// Send e-mail
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="message"></param>
        public static void Send(String subject, String message)
        {
            Boolean enable = Convert.ToBoolean(ConfigurationManager.AppSettings["Email.Enable"]);
            String host = ConfigurationManager.AppSettings["Email.SMTP.Host"].ToString();
            int port = Convert.ToInt32(ConfigurationManager.AppSettings["Email.SMTP.Port"]);
            Boolean ssl = Convert.ToBoolean(ConfigurationManager.AppSettings["Email.SMTP.SSL"]);
            String user = ConfigurationManager.AppSettings["Email.SMTP.User"].ToString();
            String password = ConfigurationManager.AppSettings["Email.SMTP.Password"].ToString();
            String to = ConfigurationManager.AppSettings["Email.To"].ToString();

            if (!enable)
            {
                return;
            }

            try
            {
                SmtpClient client = new SmtpClient();
                client.Host = host;
                client.Port = port;
                client.EnableSsl = ssl;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential(user, password);

                MailMessage mail = new MailMessage();
                mail.To.Add(new MailAddress(to));
                mail.From = new MailAddress(user);
                mail.Subject = "Tarantus SAP Service: " + subject;
                mail.Body = message;

                client.Send(mail);
            }
            catch (System.Exception ex)
            {
                throw new System.Exception("SMTP error: " + ex.ToString());
            }
        }
    }
}
