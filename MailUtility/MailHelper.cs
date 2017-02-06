using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary.MailUtility
{
    public class MailHelper : IMailHelper
    {
        protected string smtpSection;
        protected SmtpClient smtp;
        protected MailAddress formAddress;

        private string _baseSection = "BASE";

        protected MailHelper()
        {
            smtpSection = "SMTP";
        }

        public void Send(string subject, string content, List<string> toList, List<string> ccList)
        {
            try
            {
                using (var mail = new MailMessage())
                {
                    mail.From = formAddress;
                    mail.To.Add(string.Join(",", toList.ToArray()));
                    mail.Subject = (IniHelper.GetValue(_baseSection, "ISTEST") != null && string.Compare(IniHelper.GetValue(_baseSection, "ISTEST"), "true", true) == 0
                        ? IniHelper.GetValue(_baseSection, "MAIL_PREFIX") : "") + subject;
                    mail.Body = content;
                    mail.IsBodyHtml = true;

                    smtp.Send(mail);
                }
            }
            catch (Exception e)
            {
                LogHelper.Write(e.Message);
                throw e;
            }
        }

        public void SetSmtpClient(SmtpClient smtp)
        {
            this.smtp = smtp;
        }

        public void SetFormAddress(string address,string alias)
        {
            this.formAddress = new MailAddress(address, alias);
        }

    }
}
