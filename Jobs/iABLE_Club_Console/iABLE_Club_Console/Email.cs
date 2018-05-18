using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace iABLE_Club_Console
{
    public class Email
    {
        private string _smtpMasterHost = ConfigurationManager.AppSettings.Get("MasterSMTP");
        private string _smtpSlaveHost = ConfigurationManager.AppSettings.Get("SlaveSMTP");

        private string _mailToAddress = ConfigurationManager.AppSettings.Get("MailTo");
        private string _cc = "";
        public string MailToAddress
        {
            get { return _mailToAddress; }
            set { this._mailToAddress = value; }
        }

        public string CC
        {
            get { return _cc; }
            set { this._cc = value; }
        }

        private string _mailFrom = ConfigurationManager.AppSettings.Get("MailFrom");
        public string MailFrom
        {
            get { return _mailFrom; }
        }

        private string _subject = string.Empty;
        public string Subject
        {
            get { return _subject; }
            set { _subject = value; }
        }

        private string _mailBody = string.Empty;
        public string MailBody
        {
            get { return _mailBody; }
            set { _mailBody = value; }
        }

        public void SendEmail()
        {
            MailMessage mail = new MailMessage();
            mail.SubjectEncoding = System.Text.Encoding.UTF8;
            mail.BodyEncoding = System.Text.Encoding.UTF8;

            mail.Priority = MailPriority.High;
            mail.Subject = _subject;

            MailAddress from = new MailAddress(_mailFrom);
            mail.From = from;

            string[] mailTo = _mailToAddress.Split(',');
            foreach (string m in mailTo)
            {
                MailAddress addr = new MailAddress(m);
                mail.To.Add(addr);
            }

            string[] cc = _cc.Split(',');
            foreach (string c in cc)
            {
                if (!string.IsNullOrEmpty(c))
                {
                    MailAddress ccAddr = new MailAddress(c);
                    mail.CC.Add(ccAddr);
                }
            }

            mail.IsBodyHtml = true;
            mail.Body = _mailBody;

            SmtpClient masterSMTP = new SmtpClient(_smtpMasterHost);
            try
            {
                masterSMTP.Send(mail);
            }
            catch (SmtpFailedRecipientsException smtpEx)
            {
                for (int i = 0; i < smtpEx.InnerExceptions.Length; i++)
                {
                    SmtpStatusCode status = smtpEx.InnerExceptions[i].StatusCode;
                    // If mail server is busy or server is unavailable, mailer will resend mail in 5 minutes.
                    if (status == SmtpStatusCode.MailboxBusy || status == SmtpStatusCode.MailboxUnavailable)
                    {
                        System.Threading.Thread.Sleep(100000);
                        masterSMTP.Send(mail);
                    }

                }

            }
            catch
            {
                SmtpClient slaveSMTP = new SmtpClient(_smtpSlaveHost);
                try
                {
                    slaveSMTP.Send(mail);
                }
                catch (SmtpFailedRecipientsException smtpEx)
                {
                    for (int i = 0; i < smtpEx.InnerExceptions.Length; i++)
                    {
                        SmtpStatusCode status = smtpEx.InnerExceptions[i].StatusCode;
                        // If mail server is busy or server is unavailable, mailer will resend mail in 5 minutes.
                        if (status == SmtpStatusCode.MailboxBusy || status == SmtpStatusCode.MailboxUnavailable)
                        {
                            System.Threading.Thread.Sleep(100000);
                            slaveSMTP.Send(mail);
                        }

                    }

                }
            }
        }
    }
}
