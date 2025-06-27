using System;
using System.Net;
using System.Net.Mail;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace pdf2images
{
    public class EmailService
    {
        private readonly IConfiguration _configuration;
        private readonly string _host;
        private readonly int _port;
        private readonly string _username;
        private readonly string _recipients;

        public EmailService(IConfiguration configuration)
        {
            _configuration = configuration;
    
            
            _host = _configuration["SMTP:Host"] ?? "smtp.example.com";
            _recipients = _configuration["SMTP:Recipients"] ?? "hualin1.zhu@voith.com";
            if (int.TryParse(_configuration["SMTP:Port"], out int port))
            {
                _port = port;
            }
            else
            {
                _port = 25;
            }
            _username = _configuration["SMTP:Username"] ?? "";
        }

        public async Task SendEmailAsync(string subject, string body, bool isHtml = false)
        {
            try
            {
                using var client = new SmtpClient(_host, _port);
                
                // If username is provided, we assume authentication is required
                if (!string.IsNullOrEmpty(_username))
                {
                    client.Credentials = new NetworkCredential(_username, _configuration["SMTP:Password"]);
                }
                
                // You might want to set these properties based on your needs
                client.EnableSsl = _port == 587 || _port == 465;
                
                var message = new MailMessage
                {
                    From = new MailAddress(_username.Length > 0 ? _username : "noreply@voith.com"),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = isHtml
                };
                
                message.To.Add(_recipients);

                await client.SendMailAsync(message);
     
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        public async Task SendEmailWithAttachmentAsync(string to, string subject, string body, string attachmentPath, bool isHtml = false)
        {
            try
            {
                using var client = new SmtpClient(_host, _port);
                
                if (!string.IsNullOrEmpty(_username))
                {
                    client.Credentials = new NetworkCredential(_username, _configuration["SMTP:Password"]);
                }
                
                client.EnableSsl = _port == 587 || _port == 465;
                
                var message = new MailMessage
                {
                    From = new MailAddress(_username.Length > 0 ? _username : "noreply@voith.com"),
                    Subject = subject,
                    Body = body,
                    IsBodyHtml = isHtml
                };
                
                message.To.Add(to);
                
                if (File.Exists(attachmentPath))
                {
                    message.Attachments.Add(new Attachment(attachmentPath));
                }
               

                await client.SendMailAsync(message);
               
            }
            catch (Exception ex)
            {
               
                throw;
            }
        }
    }
}