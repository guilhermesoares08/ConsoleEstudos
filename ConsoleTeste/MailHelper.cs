using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleTeste
{
    public class MailHelper
    {
        static bool mailSent = false;
        private static void SendCompletedCallback(object sender, AsyncCompletedEventArgs e)
        {
            
            // Get the unique identifier for this asynchronous operation.
            String token = (string)e.UserState;

            if (e.Cancelled)
            {
                Console.WriteLine("[{0}] Send canceled.", token);
            }
            if (e.Error != null)
            {
                Console.WriteLine("[{0}] {1}", token, e.Error.ToString());
            }
            else
            {
                Console.WriteLine("Message sent.");
            }
            mailSent = true;
        }
        public static void Send(string fromAddress, string password, string toAddress)
        {
            SmtpClient client = new SmtpClient
            { 
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                EnableSsl = true,
                Host = "smtp.gmail.com",
                Port = 587,
                Credentials = new NetworkCredential(fromAddress, password)
            };

            string subject = "teste emails";
            string body = "xsl talvez";

            try
            {
                Console.WriteLine("Enviando email...");
                client.Send(fromAddress, toAddress, subject, body);
                Console.WriteLine("Email enviado!");
            }
            catch(SmtpException ex)
            {
                Console.BackgroundColor = ConsoleColor.Red;
                Console.WriteLine(ex.Message.ToString());
                Console.ResetColor();
            }
            //StringBuilder sbTo = new StringBuilder();
            //foreach (var item in to)
            //{
            //    sbTo.Append(to);
            //    sbTo.Append(",");
            //}
            //MailOptions mailOptions = new MailOptions(to, bcc, xslPath);
            //MailMessage message = new MailMessage(from.Address, sbTo.ToString());
            //message.Body = "This is a test email message sent by an application. ";
            //message.BodyEncoding = System.Text.Encoding.UTF8;
            //client.SendCompleted += new
            //SendCompletedEventHandler(SendCompletedCallback);
            
            //var attachment = new Attachment(File.Open("fileFullPath", FileMode.Open), "xslPath");
            //attachment.ContentType = new ContentType("application/vnd.ms-excel");
            ////attachmentCollection.Add(attachment);
        }
    }
}
