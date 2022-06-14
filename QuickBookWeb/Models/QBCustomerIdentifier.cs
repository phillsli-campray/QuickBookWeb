using System;
using System.Runtime.Serialization;

namespace QuickBookWeb.Models
{
    [Serializable]
    public class QBCustomerIdentifiers
    {
        public string QBId { get; set; }
        public int CHId { get; set; }
    }
}
