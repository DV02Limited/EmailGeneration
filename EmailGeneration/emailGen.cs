using System;
using System.Collections.Specialized;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Web.UI.WebControls;

// Args[] expected array composition
// args[0] Username
// args[1] User Number
// args[2] Date
// args[3] Amount Exceeded (Data or £s)
// args[4] Customer
// args[5] Subject
// args[6] File Path

namespace Email
{
    public enum Customers { DV02, Grunenthal, OPEL, OPUK, Cleanbrite, Nordea};

    public class EmailGeneration
    {
        public MailMessage EmailMessage { get; private set; }

        public void Gen([In, MarshalAs(UnmanagedType.LPArray, SizeConst = 7)]string[] args)
        {
            string CustomerEmail;

            //Set customer enail address defined on contents of CKeyResult
            switch (Enum.Parse(typeof(Customers), args[4]))
            {
                case Customers.DV02:
                    CustomerEmail = "hdcalls@dv02.co.uk";
                    break;
                case Customers.Grunenthal:
                    CustomerEmail = "uk.ithelpdesk@grunenthal.com";
                    break;
                case Customers.OPEL:
                    CustomerEmail = "itsupport@otsuka-europe.com";
                    break;
                case Customers.OPUK:
                    CustomerEmail = "itsupport@otsuka-europe.com";
                    break;
                case Customers.Cleanbrite:
                    CustomerEmail = "customercare@cleanbrite.co.uk";
                    break;
                case Customers.Nordea:
                    CustomerEmail = "dlldlondonit@nordea.com";
                    break;
                default:
                    CustomerEmail = null;
                    break;
            }

            string emailBody = File.ReadAllText(args[6]); //Read email body text and place in variable emailBody

            MailDefinition md = new MailDefinition()
            {
                From = "noreply@dv02.co.uk",
                IsBodyHtml = true,
                Subject = args[5]
            };

            //Defines variables in email body and what arguments they will be replaced with
            ListDictionary replacements = new ListDictionary
            {
                { "{name}", args[0] },
                { "{mobile}", args[1] },
                { "{date}", args[2] },
                { "{Amount_Exceeded}", args[3] }
            };

            try
            {
                // Creates email message in variable msg
                EmailMessage = CreateEmail(CustomerEmail, emailBody, md, replacements);
            }
            catch (NullReferenceException e) // Exception thrown if variables not defined for email message creation
            {
                Console.WriteLine("Error: Email variables not defined");
                Console.WriteLine(e.ToString());
            }
            catch (Exception e) //Exception thrown for undefined error
            {
                Console.WriteLine("An unknown error has occurred:");
                Console.WriteLine(e.ToString());
            }
            
            //Define SMTP Connection
            SmtpClient client = new SmtpClient()
            {
                Port = 25,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Host = "dv02-co-uk.mail.protection.outlook.com"
            };

            // Send email
            client.Send(EmailMessage);
        }

        private static MailMessage CreateEmail(string CustomerEmail, string emailBody, MailDefinition md, ListDictionary replacements)
        {
            MailMessage msg = md.CreateMailMessage(CustomerEmail, replacements, emailBody, new System.Web.UI.Control());
            return msg;
        }
    }
}