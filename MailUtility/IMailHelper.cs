using System.Collections.Generic;

namespace CommonLibrary.MailUtility
{
    public interface IMailHelper
    {
        void Send(string subject, string content, List<string> toList, List<string> ccList);
        void Send(string subject, string content, List<string> toList, List<string> ccList, List<string> bccList);
    }
}