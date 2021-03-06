﻿using System;
using System.Collections.Specialized;
using System.IO;
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
        public static string CustomerEmail { get; set; }
        

        public int Gen([In, MarshalAs(UnmanagedType.LPArray, SizeConst = 7)]string[] args)
        {

            try
            {
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
            }
            catch (ArgumentException e)
            {
                Console.WriteLine($"Customer name {args[4]} does not match");
                Console.WriteLine(e.ToString());
                return 1;
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

            // Creates email message in variable msg
            CreateEmail(CustomerEmail, emailBody, md, replacements);
            return 0;
        }

        private static void CreateEmail(string CustomerEmail, string emailBody, MailDefinition md, ListDictionary replacements)
        {
            MailMessage msg = md.CreateMailMessage(CustomerEmail, replacements, emailBody, new System.Web.UI.Control());

            //Define SMTP Connection
            SmtpClient client = new SmtpClient()
            {
                Port = 25,
                DeliveryMethod = SmtpDeliveryMethod.Network,
                UseDefaultCredentials = false,
                Host = "dv02-co-uk.mail.protection.outlook.com"
            };

            // Send email
            client.Send(msg);
        }

        public int CreateErrorEmail(string emailBody)
        {
            try
            {
                MailMessage msg = new MailMessage("noreply@dv02.co.uk", "hdcalls@dv02.co.uk")
                {
                    Subject = "Mobile Usage Alert Discrepancy",
                    Body = emailBody,
                    IsBodyHtml = true
                };

                //Define SMTP Connection
                SmtpClient client = new SmtpClient()
                {
                    Port = 25,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    UseDefaultCredentials = false,
                    Host = "dv02-co-uk.mail.protection.outlook.com"
                };

                // Send email
                client.Send(msg);
                return 0;
            }
            catch (ArgumentException e)
            {
                Console.WriteLine($"Error sending email");
                Console.WriteLine(e.ToString());
                return 1;
            }

        }
    }
}