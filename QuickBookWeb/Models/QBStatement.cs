using System;
using System.Collections.Generic;
using System.Runtime.Serialization;


namespace QuickBookWeb.Models
{
    [Serializable]
    public class QBStatement
    {
        public string Title { get; set; }
        public DateTime Date { get; set; }
        public List<string> LineItems { get; set; }
    }
}
