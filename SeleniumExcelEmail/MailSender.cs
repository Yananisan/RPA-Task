using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace SeleniumExcelEmail
{
    class MailSender
    {
        LoadToExcel data = new LoadToExcel();        

        public void ExcelSender(string email)
        {
            data.LoadDataToExcel();

            using (SmtpClient client = new SmtpClient("smtp.gmail.com"))
            {
                client.Port = 587;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential("yanaivanovskaya99@gmail.com", "*****");
                client.EnableSsl = true;
                MailMessage mail = new MailMessage();

                Attachment attach = new Attachment(@"D:\MicrowavesBook.xlsx");
                mail.Attachments.Add(attach);
                mail.IsBodyHtml = true;
                mail.Subject = "ExcelSender";
                mail.Body = "ExcelBook";
                mail.To.Add(email);
                mail.From = new MailAddress("yanaivanovskaya99@gmail.com");
                client.Send(mail);
            }
        }        
    }
}
