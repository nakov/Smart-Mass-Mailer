using System;
using System.Net;
using System.Net.Mail;
using Microsoft.Extensions.Configuration;
using System.IO;
using System.Text;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Threading;

class SmartMassMailer
{
    static IConfigurationRoot Configuration { get; set; }
    static SmtpClient SmtpClient { get; set; }
    
    static SmartMassMailer()
    {
        var builder = new ConfigurationBuilder()
            .AddJsonFile("appsettings.json");
        Configuration = builder.Build();

        SmtpClient = new SmtpClient(Configuration["smtpSettings:host"])
        {
            UseDefaultCredentials = false,
            Credentials = new NetworkCredential(
                Configuration["smtpSettings:username"],
                Configuration["smtpSettings:password"])
        };

        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    static void SendEmail(string recipientEmail, string subject, string bodyHtml)
    {
        var msg = new MailMessage
        {
            From = new MailAddress(
                Configuration["fromEmail"],
                Configuration["fromName"]),
            IsBodyHtml = true,
            Body = bodyHtml,
            BodyEncoding = Encoding.UTF8,
            SubjectEncoding = Encoding.UTF8,
            HeadersEncoding = Encoding.UTF8,
            Subject = subject
        };
        msg.To.Add(recipientEmail);
        SmtpClient.Send(msg);
    }

    static void Main()
    {
        int delayBetweenEmails = int.Parse(
            Configuration["delayBetweenEmailsMilliseconds"]);
        int startRow = int.Parse(Configuration["startRow"]);
        if (startRow < 2)
            startRow = 2;

        using (var package = new ExcelPackage(
            new FileInfo(Configuration["recipientsExcelFile"])))
        {
            ExcelWorksheet workSheet = package.Workbook.Worksheets[1];
            
            // Read Excel file column titles (the first worhsheet row)
            var columnIndexByName = new Dictionary<string, int>();
            for (int col = 1; col <= workSheet.Dimension.End.Column; col++)
            {
                string columnName = workSheet.Cells[1, col].Value
                    .ToString().ToLower();
                columnIndexByName[columnName] = col;
            }

            // Process all rows from the Excel file --> send email for each row
            for (int row = startRow; row <= workSheet.Dimension.End.Row; row++)
            {
                string email = (string) workSheet.Cells[row, columnIndexByName["email"]].Value;
                string personName = (string) workSheet.Cells[row, columnIndexByName["name"]].Value;
                string subject = Configuration["emailSubject"];
                string bodyHtml = File.ReadAllText(
                    Configuration["emailHtmlTemplate"], Encoding.UTF8);
                bodyHtml = bodyHtml.Replace("[name]", personName);
                Console.Write($"[{row} of {workSheet.Dimension.End.Row}] Sending email to: {email} (Name = {personName}) ... ");
                SendEmail(email, subject, bodyHtml);
                Console.WriteLine("Done.");
                Thread.Sleep(delayBetweenEmails);
            }
        }            
    }
}
