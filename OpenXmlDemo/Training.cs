using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Office2010.ExcelAc;

namespace OpenXmlDemo
{
    public class Training
    {

        public Training()
        {
            Contents = new List<string>();
            Attendees = new List<Person>();
        }
        public string Title { get; set; }
        public DateTime From { get; set; }

        public DateTime To { get; set; }

        public List<string> Contents { get; set; }

        public List<Person> Attendees { get; set; }
    }
}