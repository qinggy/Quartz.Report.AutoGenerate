using Esd.Report.AutoGenerate.Service;
using log4net;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;

namespace Esd.Report.AutoGenerate.Application
{
    public class CommHelper
    {
        public readonly static ILog AppLogger = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public static string MD5Encrypt(string pToEncrypt, string sKey)
        {
            DESCryptoServiceProvider des = new DESCryptoServiceProvider();
            byte[] inputByteArray = Encoding.Default.GetBytes(pToEncrypt);
            des.Key = ASCIIEncoding.ASCII.GetBytes(sKey);
            des.IV = ASCIIEncoding.ASCII.GetBytes(sKey);
            MemoryStream ms = new MemoryStream();
            CryptoStream cs = new CryptoStream(ms, des.CreateEncryptor(), CryptoStreamMode.Write);
            cs.Write(inputByteArray, 0, inputByteArray.Length);
            cs.FlushFinalBlock();
            StringBuilder ret = new StringBuilder();
            foreach (byte b in ms.ToArray())
            {
                ret.AppendFormat("{0:X2}", b);
            }
            ret.ToString();
            return ret.ToString();
        }

        public static bool SendMailUseGmail(EmailDetail content)
        {
            bool hasSend = true;
            MailMessage msg = new MailMessage();
            GenerateMailMsg(content, msg);
            SmtpClient client = new SmtpClient();
            client.Host = content.SendServer;
            client.Port = int.Parse(content.Port);
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(content.SendAddress, MD5Encrypt(content.Password, "angusesd"));
            client.DeliveryMethod = SmtpDeliveryMethod.Network;

            try
            {
                client.Send(msg);
            }
            catch (SmtpException ex)
            {
                AppLogger.Error(ex.Message, ex.InnerException);
                hasSend = false;
            }

            return hasSend;
        }

        private static void GenerateMailMsg(EmailDetail content, MailMessage msg)
        {
            AddRecipients(content.Recipients, msg);
            AddCopies(content.Cc, msg);
            msg.From = new MailAddress(content.SendAddress, content.Theme, Encoding.UTF8);
            msg.Subject = content.Theme;
            msg.SubjectEncoding = Encoding.UTF8;
            msg.Body = content.Body;
            msg.Attachments.Add(new Attachment(content.Attachment));
            msg.BodyEncoding = Encoding.UTF8;
            msg.IsBodyHtml = true;
            msg.Priority = MailPriority.High;
        }

        private static void AddCopies(string contentCC, MailMessage msg)
        {
            if (!string.IsNullOrEmpty(contentCC) && contentCC.IndexOf(";") > 0)
            {
                var copies = contentCC.Split(';');
                if (copies != null)
                {
                    foreach (var item in copies)
                    {
                        if (!string.IsNullOrEmpty(item))
                        {
                            msg.CC.Add(item);
                        }
                    }
                }
            }
        }

        private static void AddRecipients(string contentTo, MailMessage msg)
        {
            var recipients = contentTo.Split(';');
            if (recipients != null)
            {
                foreach (var item in recipients)
                {
                    if (!string.IsNullOrEmpty(item))
                    {
                        msg.To.Add(item);
                    }
                }
            }
        }

        public static string ConvertXlsToPdf(Stream xlsStream, string companyId, string fileName)
        {
            Workbook workbook = new Workbook();
            workbook.LoadFromStream(xlsStream);
            var pdfDir = Path.Combine(AppSettings.GetValue("pdfSaveDirectory"), companyId);
            if (!Directory.Exists(pdfDir)) Directory.CreateDirectory(pdfDir);
            var filePath = Path.Combine(pdfDir, fileName + DateTime.Now.ToString("yyyyMMddHHmmssffff") + ".pdf");
            workbook.SaveToFile(filePath);
            return filePath;
        }
    }
}
