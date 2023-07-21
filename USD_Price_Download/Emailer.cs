using System.Configuration;
using System.Data;
using System.Net;
using System.Net.Mail;

namespace USD_Price_Download
{
    class Emailer
    {
        public static void SentMail(string subject, string msg, int hasAttachment, string path)
        {
            MailMessage message = new MailMessage();
            string fileName = ConfigurationSettings.AppSettings["EMailIds"].ToString();
            DataSet dataSet = new DataSet();
            int num = (int)dataSet.ReadXml(fileName);
            foreach (DataRow row in (InternalDataCollectionBase)dataSet.Tables["Emails"].Rows)
                message.To.Add(new MailAddress(row["MailId"].ToString(), row["Alias"].ToString()));
            SmtpClient smtpClient = new SmtpClient("smtpout.secureserver.net");
            message.From = new MailAddress("admin6@wealtherp.com", "WealthERP");
            message.Subject = subject;
            message.IsBodyHtml = true;
            message.Body = "Dear Sidharth,<br /> " + msg + "<br /> <br />Thanks & Regards <br />  WealthERP Team.";
            smtpClient.Port = 3535;
            if (hasAttachment == 1)
            {
                Attachment attachment = new Attachment(path);
                message.Attachments.Add(attachment);
            }
            smtpClient.Credentials = (ICredentialsByHost)new NetworkCredential("admin6@wealtherp.com", "Ampsys123#");
            smtpClient.EnableSsl = false;
            smtpClient.Send(message);
        }
    }
}
