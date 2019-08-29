using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DownPaymentLogic.Classes
{
    public  class DataForCalculationRate
    {
        public string BplName { get; set; }
        public string BusinesPartnerName { get; set; }
        public string DocNum { get; set; }
        public string CardCode { get; set; }
        public string DownPaymentAmount { get; set; }
        public string FormTypex { get; set; }
        public DateTime PostingDate { get; set; }
        public decimal TotalInv { get; set; }
        public decimal RateInv { get; set; }
        public string DocCurrency { get; set; }
        public string FormUIdInv { get; set; }
        public string FormUIdDps { get; set; }
        public string GlobalRate { get; set; }
        public bool IsCalculated { get; set; }
        public List<Dictionary<string, string>> NetAmountsForDownPayment { get; set; }
        public List<Dictionary<string, string>> GrossAmountsForDownPayment { get; set; }

    }
}
