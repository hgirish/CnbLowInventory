using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using Dapper;
using Microsoft.Azure.WebJobs;
using Microsoft.Extensions.Logging;

namespace CnbLowInventory
{
    class Inventory
    {
        public string Sku { get; set; }
        public string Product { get; set; }
        public int AvailableUnits { get; set; }
        public int MinimumRequired { get; set; }
        public int ProductId { get; set; }
    }
    public static class LowInventory
    {
        // 0 */5 * * * *	every 5 minutes
        // 0 0 * * * *	every hour (hourly)
        //  0 0 0 * * *	every day(daily)
        // 0 0 0 * * 0	every sunday (weekly)
        // 0 0 9 * * MON	every monday at 09:00:00
        [FunctionName("LowInventory")]
        public static void Run([TimerTrigger("0 0 9 * * MON", RunOnStartup =true)]TimerInfo myTimer, ILogger log) // Every monday at 9
        {
            var connStr = GetSqlConnectionString("sqldb_connection");
     
            using (SqlConnection conn = new SqlConnection(connStr))
            {        
                
               
                var sql = @"select sku, LEFT(title, 90) as Product, unitsinstock as AvailableUnits, reorderlevel as MinimumRequired, ProductId 
from hsi_products
where clientid = 18 and departmentid = 742 and reorderlevel > 0 and unitsinstock<reorderlevel and (discontinued is null or discontinued = 0)
";
                var details = conn.Query<Inventory>(sql).ToList();
                if (details.Count > 0)
                {
                  var table =  GenerateMessage(details);
                      SendEmailAsync(table, log).Wait();
                }
               
                log.LogInformation($"Rows returned: {details.Count}");
            }
            log.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
        }

        public static string GetSqlConnectionString(string name)
        {
            string conStr = Environment.GetEnvironmentVariable($"ConnectionStrings:{name}", EnvironmentVariableTarget.Process);
            if (string.IsNullOrEmpty(conStr)) // Azure Functions App Service naming convention
                conStr = Environment.GetEnvironmentVariable($"SQLCONNSTR_{name}", EnvironmentVariableTarget.Process);
            return conStr;
        }
        public static string GetEnvironmentVariable(string name)
        {
            return Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);
        }
        private static async Task SendEmailAsync(string message, ILogger log)
        {           
            var username = GetEnvironmentVariable("Smtp:UserName");
            var password = GetEnvironmentVariable("Smtp:Password");
            int.TryParse(GetEnvironmentVariable("Smtp:Port"), out int port);
            if (port == 0)
            {
                port = 587;
            }
            int.TryParse(GetEnvironmentVariable("Smtp:DeliveryMethod"), out int deliveryMethod);
            var host = GetEnvironmentVariable("Smtp:Host");
            var to = GetEnvironmentVariable("Smtp:To");
            var cc = GetEnvironmentVariable("Smtp:Cc");
            var pickupDirectory = GetEnvironmentVariable("Smtp:PickupDirectory");
           
            var emailMessage = new MailMessage();
            emailMessage.From = new MailAddress(username, "gohands");
            var tos = to.Split(",");
            foreach (var item in tos)
            {
                emailMessage.To.Add(item);
            }
            if (!string.IsNullOrEmpty(cc))
            {
                var ccs = cc.Split(",");
                foreach (var email in ccs)
                {
                    emailMessage.CC.Add(email);
                }
            }
            emailMessage.Subject = "Low Inventory Alert for Event Items" ;
            emailMessage.Body = message;
            emailMessage.IsBodyHtml = true;
            using (var client = new SmtpClient())
            {
                if (deliveryMethod == (int) SmtpDeliveryMethod.SpecifiedPickupDirectory)
                {
                    client.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                    client.PickupDirectoryLocation = pickupDirectory;

                }
                else
                {
                    client.DeliveryMethod = SmtpDeliveryMethod.Network;
                    client.Host = host;
                    client.Port = port;
                    client.EnableSsl = true;
                    client.Credentials = new NetworkCredential(username, password);
                }
                await client.SendMailAsync(emailMessage);
                log.LogInformation($"Email sent to: {to}, cc: {cc} at {DateTime.UtcNow} UTC");
            }

        }

        private static string  GenerateMessage(List<Inventory> details)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("<h2>Following items are low in stock</h2>");
            sb.AppendLine("<table border='1' cellpadding='10' style='border-collapse:collapse;'><thead><tr><th>Sku</th><th>Product</th><th>Available Units</th><th>Minimum Required</th></tr></thead><tbody>");
            foreach (var item in details)
            {
                sb.AppendLine($"<tr><td>{item.Sku}</td><td>{item.Product}</td><td style='color:red;'>{item.AvailableUnits}</td><td style='color:orange;'>{item.MinimumRequired}</td></tr>");
            }
            sb.AppendLine("</tbody></table>");
            return sb.ToString();
        }

      
    }
}
