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

        public void SetSmtpClient(SmtpClient smtp)
        {
            this.smtp = smtp;
        }

        public void SetFormAddress(string address, string alias)
        {
            this.formAddress = new MailAddress(address, alias);
        }

        public void Send(string subject, string content, List<string> toList, List<string> ccList)
        {
            Send(subject, content, toList, ccList, null);
        }

        public void Send(string subject, string content, List<string> toList, List<string> ccList, List<string> bccList)
        {
            try
            {
                using (var mail = new MailMessage())
                {
                    mail.From = formAddress;
                    if (toList == null)
                    {
                        throw new Exception("Send To is Empty!!");
                    }
                    mail.To.Add(string.Join(",", toList.Where(e => !string.IsNullOrEmpty(e)).ToArray()));
                    if (ccList != null)
                        mail.CC.Add(string.Join(",", ccList.Where(e => !string.IsNullOrEmpty(e)).ToArray()));
                    if (bccList != null)
                        mail.Bcc.Add(string.Join(",", bccList.Where(e => !string.IsNullOrEmpty(e)).ToArray()));

                    var defaultBcc = IniHelper.GetValue(_baseSection, "MAIL_DEFAULT_BCC");
                    if (defaultBcc != null && !string.IsNullOrWhiteSpace(defaultBcc))
                        mail.Bcc.Add(defaultBcc);

                    //處理測試環境，預設為測試
                    if (IniHelper.GetValue(_baseSection, "ISTEST") == null || string.Compare(IniHelper.GetValue(_baseSection, "ISTEST"), "false", true) != 0)
                    {
                        subject = IniHelper.GetValue(_baseSection, "MAIL_PREFIX") + subject;
                        content = "<p>To : " + mail.To + "</p><p>CC : " + mail.CC + "</p><p>Bcc : " + mail.Bcc + "</p><hr/>" + content;
                        mail.To.Clear();
                        mail.To.Add(IniHelper.GetValue(_baseSection, "MAIL_TEST_TO"));
                        mail.CC.Clear();
                        mail.Bcc.Clear();
                    }

                    mail.Subject = subject;
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

    }
}
