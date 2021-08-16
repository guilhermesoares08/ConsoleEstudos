using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTeste
{
    public class MailOptions
    {
        public List<MailAddress> To { get; set; }
        public List<MailAddress> Bcc { get; set; }

        public string XslPath { get; set; }
        public MailOptions(List<MailAddress> to, List<MailAddress> bcc, string xslPath)
        {
            To = to;
            Bcc = bcc;
            XslPath = xslPath;
        }
    }
}
