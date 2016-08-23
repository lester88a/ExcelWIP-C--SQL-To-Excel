using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelWIP.VersionTwo.EmailService
{
    public class EmailService
    {
        //default setting of the email environment
        private string _recipient;
        private string _sender;
        private string _smtpServer;
        private int _smtpPort;
        private string _senderUserName;
        private string _senderUserPass;
        private Attachment attachment;
        public static string emailEnFile = System.AppDomain.CurrentDomain.BaseDirectory + @"Email.xml";

        MailMessage message = new MailMessage();


        public EmailService()
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(emailEnFile);
            
            this._recipient = doc.DocumentElement.SelectSingleNode("/email/recipient").InnerText;
            this._sender = doc.DocumentElement.SelectSingleNode("/email/sender").InnerText;
            this._smtpServer = doc.DocumentElement.SelectSingleNode("/email/smtpServer").InnerText;
            this._smtpPort = Convert.ToInt32(doc.DocumentElement.SelectSingleNode("/email/smtpPort").InnerText);
            this._senderUserName = doc.DocumentElement.SelectSingleNode("/email/senderUserName").InnerText;
            this._senderUserPass = doc.DocumentElement.SelectSingleNode("/email/senderUserPass").InnerText;

        }
        
        public void SendEmailMethod(string fileName, string errorSubject, string errorBody)
        {
            message.To.Add(_recipient);
            message.Subject = errorSubject;
            message.From = new MailAddress(_sender);
            message.Body = errorBody;
            SmtpClient smtp = new SmtpClient(_smtpServer, _smtpPort);

            //configure the client 
            smtp.EnableSsl = false;
            smtp.UseDefaultCredentials = false;
            smtp.Credentials = new System.Net.NetworkCredential(_senderUserName, _senderUserPass);


            //attachment
            attachment = new Attachment(fileName);
            message.Attachments.Add(attachment);

            smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtp.Send(message);

        }
    }
}
