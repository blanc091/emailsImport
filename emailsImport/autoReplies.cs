using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using static emailsImport.emImport;

namespace emailsImport
{
    internal class autoReplies
    {
        public const string SMTP_SERVER = "*********";
        public const string SMTP_USERNAME = "*********";
        public const string SMTP_PASSWORD = "*********";
        public const int SMTP_PORT = 26;

        public void SendAlertEmail(string ToAddress, string EmailSubject, int errFlag)
        {
            try
            {
                System.Net.Mail.MailMessage message = new System.Net.Mail.MailMessage
                {
                    From = new System.Net.Mail.MailAddress("autoreply@DocVault.com"), 
                    IsBodyHtml = true
                };
                message.To.Add(ToAddress.Trim());

                Logger.Log(GlobalConfig.ClientID, errFlag.ToString());

                switch (errFlag)
                {
                    case 0:
                        message.Subject = "We have received your invoice(s)";
                        message.Body = @"<font face=""Arial"" Size=2 Color=#000000>Dear Vendor,</font><br /><font face=""Arial"" Size=2 Color=#000000>This message is to confirm that we have successfully received your invoice(s) with email subject: " + EmailSubject + @".<br /> Provided no issues are found with your invoice(s), it will be paid according to your contracted payment terms.</font><br /><font face=""Arial"" Size=2 Color=#000000>Please do not reply to this message as it is automatically generated.</font><br /><font face=""Arial"" Size=2 Color=#000000>In case of any queries regarding your invoice(s), please contact [helpdesk link / address / phone No.].</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Best Regards,</font><br /><font face=""Arial"" Size=2 Color=#000000>Accounts Payable</font>";
                        break;
                    case 2:
                        message.Subject = "We could not find a valid attachment added to your email";
                        message.Body = @"<font face=""Arial"" Size=2 Color=#000000>Dear Vendor,</font><br /><font face=""Arial"" Size=2 Color=#000000>This message is to inform you that we could not find a valid attachment(s) in your email with subject: " + EmailSubject + @".<br /> Please resend your email and ensure your attachments adhere to the below requirements: </font><br /><font face=""Arial"" Size=2 Color=#000000>- Valid attachments are .PDF or .TIFF file formats. Any other file format will be rejected.</font><br /><font face=""Arial"" Size=2 Color=#000000>- Each attachment must contain only one invoice and its supporting documentation (if applicable). A single invoice split across multiple attachments will result in a processing rejection. A single attachment containing multiple invoices will also result in a processing rejection.</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Please do not reply to this message as it is automatically generated.</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>In case of any queries regarding your invoice(s), please contact [helpdesk link / address / phone No.].</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Best Regards,</font><br /><font face=""Arial"" Size=2 Color=#000000>Accounts Payable</font>";
                        break;
                    case 4:
                        message.Subject = "Invalid attachments have been identified in your email";
                        message.Body = @"<font face=""Arial"" Size=2 Color=#000000>Dear Vendor,</font><br /><font face=""Arial"" Size=2 Color=#000000>This message is to inform you that not all of the attachments in your email are valid." + @"<br /><br />Valid attachments in .pdf / .tiff format will be processed and provided that no issues are found with your invoice(s), it will be paid according to your contracted payment terms.</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Please change the format of the invalid attachments and resend them. Always ensure your attachments adhere to the below requirements:</font><br /><font face=""Arial"" Size=2 Color=#000000>- Valid attachments are .PDF or .TIFF file formats. Any other file format will be rejected.</font><br /><font face=""Arial"" Size=2 Color=#000000>- Each attachment must contain only one invoice and its supporting documentation (if applicable). A single invoice split across multiple attachments will result in a processing rejection. A single attachment containing multiple invoices will also result in a processing rejection.</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Please do not reply to this message as it is automatically generated.</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>In case of any queries regarding your invoice(s), please contact [helpdesk link / address / phone No.].</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Best Regards,</font><br /><font face=""Arial"" Size=2 Color=#000000>Accounts Payable</font>";
                        break;
                    default:
                        message.Subject = "We have received your invoice(s)";
                        message.Body = @"<font face=""Arial"" Size=2 Color=#000000>Dear Vendor,</font><br /><font face=""Arial"" Size=2 Color=#000000>This message is to confirm that we have successfully received your invoice(s) with email subject: " + EmailSubject + @". <br />Provided no issues are found with your invoice(s), it will be paid according to your contracted payment terms.</font><br /><font face=""Arial"" Size=2 Color=#000000>Please do not reply to this message as it is automatically generated.</font><br /><font face=""Arial"" Size=2 Color=#000000>In case of any queries regarding your invoice(s), please contact [helpdesk link / address / phone No.].</font><br /><br /><font face=""Arial"" Size=2 Color=#000000>Best Regards,</font><br /><font face=""Arial"" Size=2 Color=#000000>Accounts Payable</font>";
                        break;
                }

                using (System.Net.Mail.SmtpClient smtp = new System.Net.Mail.SmtpClient(SMTP_SERVER, SMTP_PORT))
                {
                    //smtp.EnableSsl = true;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(SMTP_USERNAME, SMTP_PASSWORD);
                    smtp.Timeout = 10000;
                    smtp.Send(message);
                }

                message.Dispose();
            }
            catch (Exception ex)
            {
                Logger.Log(GlobalConfig.ClientID, "Error sending email: " + ex.Message);
            }
        }
    }

}
