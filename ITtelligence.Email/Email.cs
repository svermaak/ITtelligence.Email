using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;

namespace ITtelligence
{
    public class EMail
    {
        public static bool Send(List<string> MailAddresses, string Host, int Port, bool SSL, int Timeout, string Domain, string UserName, string Password, string From, string Subject, string Title, string Body, string LogoImage, string BackgroundImage, out string ReturnMessage)
        {
            try
            {
                SmtpClient client = new SmtpClient();
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.Host = Host;
                client.Port = Port;
                client.EnableSsl = SSL;
                client.Timeout = Timeout;


                if ((!string.IsNullOrEmpty(Domain)) && (!string.IsNullOrEmpty(UserName)) && (!string.IsNullOrEmpty(Password)))
                {
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(UserName, Password, Domain);
                }
                else if ((!string.IsNullOrEmpty(UserName)) && (!string.IsNullOrEmpty(Password)))
                {
                    client.UseDefaultCredentials = false;
                    client.Credentials = new System.Net.NetworkCredential(UserName, Password);
                }
                else
                {
                    client.UseDefaultCredentials = false;
                }

                MailMessage mail = new MailMessage();

                mail.From = new MailAddress(From);
                foreach (string mailAddress in MailAddresses.Distinct().ToArray())
                {
                    if (!string.IsNullOrEmpty(mailAddress))
                    {
                        mail.To.Add(new MailAddress(mailAddress)); ;
                    }
                }

                // If file doesn't exist, perhaps it is base64 image?
                if (!File.Exists(LogoImage))
                {
                    LogoImage = LoadBase64ImageIntoTEMP(LogoImage);
                }
                if (!File.Exists(BackgroundImage))
                {
                    BackgroundImage = LoadBase64ImageIntoTEMP(BackgroundImage);
                }

                // Alternate View HTML version
                LinkedResource logoLinkedResource = new LinkedResource(LogoImage);
                LinkedResource backgroundLinkedResource = new LinkedResource(BackgroundImage);
                logoLinkedResource.ContentId = Guid.NewGuid().ToString();
                backgroundLinkedResource.ContentId = Guid.NewGuid().ToString();
                string avHtmlBody = $"<html><head></head><body style='background-image: url(\"cid:{backgroundLinkedResource.ContentId}\");background-repeat: repeat;'><center><img src=\"cid:{logoLinkedResource.ContentId}\"></center><h1>{Title}</h1>{Body}</body></html>";
                AlternateView avHtml = AlternateView.CreateAlternateViewFromString(avHtmlBody, null, MediaTypeNames.Text.Html);
                avHtml.LinkedResources.Add(logoLinkedResource);
                avHtml.LinkedResources.Add(backgroundLinkedResource);
                mail.AlternateViews.Add(avHtml);

                // HTML version
                Attachment backgroundAttachment = new Attachment(BackgroundImage);
                Attachment logoAttachment = new Attachment(LogoImage);
                string htmlBody = $"<html><head></head><body style='background-image: url(\"{backgroundAttachment.Name}\");background-repeat: repeat;'><center><img src=\"cid:{logoAttachment.Name}\"></center><h1>{Title}</h1>{Body}</body></html>";
                mail.Body = htmlBody;
                mail.IsBodyHtml = true;
                mail.Attachments.Add(logoAttachment);
                mail.Attachments.Add(backgroundAttachment);

                mail.Subject = Subject;
                client.Send(mail);
                ReturnMessage = "Send";

                return true;
            }
            catch (Exception ex)
            {
                ReturnMessage = ex.Message;
                Console.WriteLine(ex.Message);
                return false;
            }
        }
        private static string LoadBase64ImageIntoTEMP(string Base64String)
        {
            try
            {
                byte[] bytes = Convert.FromBase64String(Base64String);

                System.Drawing.Image image;
                using (MemoryStream ms = new MemoryStream(bytes))
                {
                    string temporaryFileName = "";

                    temporaryFileName = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid().ToString() + ".jpg");
                    image = System.Drawing.Image.FromStream(ms);
                    image.Save(temporaryFileName, System.Drawing.Imaging.ImageFormat.Jpeg);

                    return temporaryFileName;
                } 
            }
            catch
            {
                
            }
            return null;
        }
    }
}
