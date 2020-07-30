using System;
using System.Collections.Generic;
using System.Text;

namespace ItSchoolParserImap
{
    [Serializable]
    public class Message__c
    {
        public DateTime? DateSent { get; set; }
        public string MessageAsText { get; set; }


        public Message__c()
        {
            DateSent = new DateTime(2000, 7, 22);
            MessageAsText = "Leer";
        }

        public Message__c(DateTime? dateSent, string messageAsText)
        {
            DateSent = dateSent;
            MessageAsText = messageAsText;
        }
    }
}
