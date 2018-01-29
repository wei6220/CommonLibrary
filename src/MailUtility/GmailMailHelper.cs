using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary.MailUtility
{
    public class GmailMailHelper : MailHelper
    {
        public GmailMailHelper(string account, string password)
        {
            this.smtp = new SmtpClient
            {
                Host = "smtp.gmail.com",
                Port = 587,
                EnableSsl = true,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Credentials = new NetworkCredential(account, password)
            };

            this.formAddress = new MailAddress(account);
        }

    }
}
