using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web.UI.WebControls;


namespace Email
{
    public class EmailGeneration
    {
        public void Gen([In, MarshalAs(UnmanagedType.LPArray, SizeConst = 7)]string[] args)
        {

            string[] CustomerKey = new string[] { "DV02", "Grunenthal", "OPEL", "OPUK", "Cleanbrite", "Nordea" };
            string CKeyResult = CustomerKey.FirstOrDefault(s => args[4].Contains(s));
            string CustomerEmail = string.Empty;

            switch (CKeyResult)
            {
                case "DV02":
                    CustomerEmail = "hdcalls@dv02.co.uk";
                    break;
                case "Gruenthal":
                    CustomerEmail = "uk.ithelpdesk@grunenthal.com";
                    break;
                case "OPEL":
                    CustomerEmail = "itsupport@otsuka-europe.com";
                    break;
                case "OPUK":
                    CustomerEmail = "itsupport@otsuka-europe.com";
                    break;
                case "Cleanbrite":
                    CustomerEmail = "customercare@cleanbrite.co.uk";
                    break;
                case "Nordea":
                    CustomerEmail = "dlldlondonit@nordea.com";
                    break;
            }
            
            string emailBody = File.ReadAllText(args[6]);
            MailDefinition md = new MailDefinition()
            {
                From = "noreply@dv02.co.uk",
                IsBodyHtml = true,
                Subject = args[5]
            };
            ListDictionary replacements = new ListDictionary
            {
                { "{name}", args[0] },
                { "{mobile}", args[1] },
                { "{date}", args[2] },
                { "{Amount_Exceeded}", args[3] }
            };
            MailMessage msg = md.CreateMailMessage(CustomerEmail, replacements, emailBody, new System.Web.UI.Control());
            //               msg.To.Add(new MailAddress(CustomerEmail));
            //               msg.Body = emailBody;

            SmtpClient client = new SmtpClient()
            {
                Port = 25,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Host = "dv02-co-uk.mail.protection.outlook.com"
            };
            client.Send(msg);



            // Args[] expected array composition
            // args[0] Username
            // args[1] User Number
            // args[2] Date
            // args[3] Amount Exceeded (Data or £s)
            // args[4] Customer
            // args[5] Subject
            // args[6] File Path

        }
    }
}