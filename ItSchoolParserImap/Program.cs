using ImapX;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace ItSchoolParserImap
{
    class Program
    {
        private static List<Message> messages = new List<Message>();

        private static EmailAuth emailAuth;
        private static GoogleAuth googleAuth;
        private static XmlSerializer emailAuthFormatter = new XmlSerializer(typeof(EmailAuth));
        private static XmlSerializer googleAuthFormatter = new XmlSerializer(typeof(GoogleAuth));


        static void Main(string[] args)
        {
            try
            {
                using (FileStream fs = new FileStream("EmailAuth.xml", FileMode.Open))
                {
                    emailAuth = (EmailAuth)emailAuthFormatter.Deserialize(fs);
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine($"Creating a new EmailAuth.xml ...");
                using (FileStream fs = new FileStream("EmailAuth.xml", FileMode.Create))
                {
                    emailAuth = new EmailAuth();
                    emailAuthFormatter.Serialize(fs, emailAuth);
                }
            }

            EmailHandler emailHandler = new EmailHandler(emailAuth);
            messages = emailHandler.ObtainMessages();

            try
            {
                using (FileStream fs = new FileStream("GoogleAuth.xml", FileMode.Open))
                {
                    googleAuth = (GoogleAuth)googleAuthFormatter.Deserialize(fs);
                }
            }
            catch (FileNotFoundException)
            {
                Console.WriteLine($"Creating a new GoogleAuth.xml ...");
                using (FileStream fs = new FileStream("GoogleAuth.xml", FileMode.Create))
                {
                    googleAuth = new GoogleAuth();
                    googleAuthFormatter.Serialize(fs, googleAuth);
                }
            }

            GoogleSpreadsheetHandler googleSpreadsheetHandler = new GoogleSpreadsheetHandler(messages, googleAuth);
            googleSpreadsheetHandler.SpreadsheetsConnect();
        }
    }
}
