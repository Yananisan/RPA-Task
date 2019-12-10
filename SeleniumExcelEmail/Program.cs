using System;
using System.Net.Mail;

namespace SeleniumExcelEmail
{
    class Program
    {
        static void Main()
        {       
            Console.Write("Please, write email address to send data to: ");
            string email = Console.ReadLine();
            MailSender sender = new MailSender();
            sender.ExcelSender(email);
        }



           

    }
}
