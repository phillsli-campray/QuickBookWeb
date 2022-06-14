using System;
using System.Runtime.Serialization;


namespace QuickBookWeb.Models
{
    [Serializable]
    public class QBInvoice
    {
        public string LineItemReferenceId { get; set; }
        public string CustomerReferenceId { get; set; }
        public string AccountReferenceId { get; set; }
        public string InvoiceNumber { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public string Terms { get; set; }
        public double Rate { get; set; }
        public double Amount { get; set; }
        public DateTime Date { get; set; }
        public DateTime ShipDate { get; set; }
    }
}
