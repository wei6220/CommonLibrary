using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary.MailUtility
{
    public class ExchangeMailHelper : MailHelper
    {
        public ExchangeMailHelper()
        {
            try
            {
                this.smtp = new SmtpClient
                {
                    Host = IniHelper.GetValue(smtpSection, "HOST"),
                    Port = string.IsNullOrEmpty(IniHelper.GetValue(smtpSection, "PORT")) ? 25 : Convert.ToInt32(IniHelper.GetValue(smtpSection, "PORT")),
                    EnableSsl = IniHelper.GetValue(smtpSection, "ENABLESSL") != null && string.Compare(IniHelper.GetValue(smtpSection, "ENABLESSL"), "true", true) == 0 ? true : false,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                };

                this.formAddress = new MailAddress(IniHelper.GetValue(smtpSection, "FormAddress"), IniHelper.GetValue(smtpSection, "FormAlias"));
            }
            catch (Exception e)
            {
                LogHelper.Write(e.Message);
                throw e;
            }
        }
    
    }
}
