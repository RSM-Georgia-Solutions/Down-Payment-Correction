using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DownPaymentLogic.Classes;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;

namespace DownPaymentLogic
{
    public static class DownPaymentLogic
    {
        /// <summary>
        ///   
        /// </summary>
        /// <param name="downPaymentToDrow"> down payment - is forma gaxsnili invoisidan (A/R ; A/P) </param>
        /// <param name="data"></param>
        /// <param name="_comp"> SAPbobsCOM company </param>
        /// <param name="formType"> invoisis(mshobeli) formis tipi (A/R ; A/P) </param>
        /// <param name="docCurrency">invoisis(mshobeli) formis valuta </param>
        /// <param name="totalInv"> Total Befor Discounts damatebuli Tax-i invoisis pormidan </param>
        /// <param name="isRateCalculated"> tu isRateCalculated true daabruna eseigi invoisis BP Currency velshi  unda cahvsvat  globalRate </param>
        /// <param name="globalRate"> tu isRateCalculated true daabruna eseigi invoisis BP Currency velshi  unda cahvsvat  globalRate </param>
        /// <param name="ratInv">invoisis(mshobeli) formis valuta </param>
        public static void ExchangeRateCorrectionUi(DataForCalculationRate data, SAPbobsCOM.Company _comp)
        {
            data.GlobalRate = "1,0000";
            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net AmountFC To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net AmountFC To Drow 

            if (data.GrossAmountsForDownPayment == null)
            {
                return;
            }

            foreach (var downpayment in data.GrossAmountsForDownPayment)
            {
                if (data.FormTypex == "133")
                {
                    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet2.DoQuery("SELECT DocEntry FROM ODPI WHERE DocNum = '" + downpayment.First().Key + "'");
                    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

                    string orctDocEntrys =
                        "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                        dpDocEntry +
                        "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები 

                    recSet2.DoQuery(
                        "select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum',  avg(ORCT.CashSum) as 'CashSum', SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                        "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + orctDocEntrys +
                        ") group by ORCT.DocEntry");
                    // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                    if (recSet2.EoF)
                    {
                        recSet2.DoQuery($"SELECT BaseRef FROM DPI1 WHERE DocEntry = '{dpDocEntry}'");
                        dpDocEntry = recSet2.Fields.Item("BaseRef").Value.ToString();

                        orctDocEntrys =
                            $"select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '{dpDocEntry}' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები
                        recSet2.DoQuery(
                            "select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum', avg(ORCT.CashSum) as 'CashSum', SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                            "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + orctDocEntrys +
                            ") group by ORCT.DocEntry");
                    }

                    List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
                    while (!recSet2.EoF)
                    {
                        int ocrtDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                        decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) +
                                                  decimal.Parse(recSet2.Fields.Item("CashSum").Value.ToString());
                        decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                        sumPayments.Add(
                            new Tuple<int, decimal, decimal>(ocrtDocEntry, appliedAmountLc, appliedAmountFc));
                        recSet2.MoveNext();
                    }
                    //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                    Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet4.DoQuery(
                        "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  ORCT.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then RCT2.SumApplied else 0 end ) as 'LcPrices' from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" +
                        orctDocEntrys +
                        ") group by  RCT2.SumApplied , ORCT.DocEntry ) LcPricesTable group by DocEntry");

                    Dictionary<string, decimal> documentLcPriceSums = new Dictionary<string, decimal>();
                    while (!recSet4.EoF)
                    {
                        string ocrtDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                        decimal sumLcPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                        documentLcPriceSums.Add(ocrtDocEntry, sumLcPayments);
                        recSet4.MoveNext();

                        // აქ არის ლოკალურ ვალუტაში გატარებულ დოკუმენტებზე გადახდილი სტრული თანხა  
                    }

                    Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

                    List<XContainer> docsWithRateAndValue = new List<XContainer>();

                    foreach (var tuple in sumPayments)
                    {
                        var rate = (tuple.Item2 - documentLcPriceSums[tuple.Item1.ToString()]) / tuple.Item3;
                        var paymentDocEntry = tuple.Item1.ToString();
                        rateByDocuments.Add(paymentDocEntry, rate);
                        docsWithRateAndValue.Add(new XContainer
                        {
                            CurrRate = rate,
                            OrctDocEntry = paymentDocEntry
                        });
                    }
                    // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში 
                    Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet3.DoQuery(
                        "select ORCT.DocEntry, RCT2.DocEntry as 'DpDocEntry',    RCT2.AppliedFC from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry  in ( " +
                        orctDocEntrys + ") and RCT2.DocEntry = '" + dpDocEntry + "' and InvType = 203");
                    Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                    while (!recSet3.EoF)
                    {
                        decimal appliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                        string paymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                        dPIncomingPaymentShareAmountFc.Add(paymentDocEntry, appliedFcbyDp);
                        docsWithRateAndValue.Where(x => x.OrctDocEntry == paymentDocEntry).ToList()
                            .ForEach(s => s.AmountFC = appliedFcbyDp);
                        recSet3.MoveNext();
                    }

                    decimal lcSum = docsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                    decimal fcSum = docsWithRateAndValue.Sum(doc => doc.AmountFC);
                    decimal weightedRate = lcSum / fcSum;
                    paidAmountDpLc += decimal.Parse(downpayment.First().Value) * weightedRate;
                }
                else
                {
                    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + downpayment.First().Key + "'");
                    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

                    string orctDocEntrys =
                        $"select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = ' { dpDocEntry}' and InvType = 204 and OVPM.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები 

                    recSet2.DoQuery($"select  OVPM.DocEntry, avg(OVPM.TrsfrSum) as 'TrsfrSum',  avg(OVPM.CashSum) as 'CashSum', SUM(VPM2.AppliedFC) as 'AppliedFC' from OVPM inner join VPM2 on " +
                                    $"OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (  {orctDocEntrys} ) group by OVPM.DocEntry");
                    // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                    if (recSet2.EoF)
                    {
                        recSet2.DoQuery($"SELECT BaseRef FROM DPO1 WHERE DocEntry = '{dpDocEntry}'");
                        dpDocEntry = recSet2.Fields.Item("BaseRef").Value.ToString();

                        orctDocEntrys =
                            $"select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '{dpDocEntry}' and InvType = 204 and OVPM.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები
                        recSet2.DoQuery("select OVPM.DocEntry, avg(OVPM.TrsfrSum) as 'TrsfrSum',  avg(OVPM.CashSum) as 'CashSum', SUM(VPM2.AppliedFC) as 'AppliedFC' from OVPM inner join VPM2 on " +
                                        "OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" + orctDocEntrys + ") group by OVPM.DocEntry");
                    }

                    List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
                    while (!recSet2.EoF)
                    {
                        int ocrtDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                        decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) + decimal.Parse(recSet2.Fields.Item("CashSum").Value.ToString());
                        decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                        sumPayments.Add(new Tuple<int, decimal, decimal>(ocrtDocEntry, appliedAmountLc, appliedAmountFc));
                        recSet2.MoveNext();
                    }
                    //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                    Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet4.DoQuery(
                        "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  OVPM.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then VPM2.SumApplied else 0 end ) as 'LcPrices' from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" +
                        orctDocEntrys +
                        ") group by  VPM2.SumApplied , OVPM.DocEntry ) LcPricesTable group by DocEntry");

                    Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
                    while (!recSet4.EoF)
                    {
                        string ocrtDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                        decimal sumLcPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                        DocumentLcPriceSums.Add(ocrtDocEntry, sumLcPayments);
                        recSet4.MoveNext();

                        // აქ არის ლოკალურ ვალუტაში გატარებულ დოკუმენტებზე გადახდილი სტრული თანხა  
                    }

                    Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

                    List<XContainer> DocsWithRateAndValue = new List<XContainer>();

                    foreach (var tuple in sumPayments)
                    {
                        var rate = (tuple.Item2 - DocumentLcPriceSums[tuple.Item1.ToString()]) / tuple.Item3;
                        var paymentDocEntry = tuple.Item1.ToString();
                        rateByDocuments.Add(paymentDocEntry, rate);
                        DocsWithRateAndValue.Add(new XContainer
                        {
                            CurrRate = rate,
                            OrctDocEntry = paymentDocEntry
                        });
                    }
                    // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში 
                    Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet3.DoQuery("select OVPM.DocEntry, VPM2.DocEntry as 'DpDocEntry',    VPM2.AppliedFC from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry  in ( " + orctDocEntrys + ") and VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204");
                    Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                    while (!recSet3.EoF)
                    {
                        decimal appliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                        string paymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                        dPIncomingPaymentShareAmountFc.Add(paymentDocEntry, appliedFcbyDp);
                        DocsWithRateAndValue.Where(x => x.OrctDocEntry == paymentDocEntry).ToList().ForEach(s => s.AmountFC = appliedFcbyDp);
                        recSet3.MoveNext();
                    }

                    decimal lcSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                    decimal fcSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
                    decimal weightedRate = lcSum / fcSum;
                    paidAmountDpLc += decimal.Parse(downpayment.First().Value) * weightedRate;
                }
                try
                {
                    paidAmountDpFc += decimal.Parse(downpayment.First().Value);
                }
                catch (Exception e)
                {
                    Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            CalculateWaightedRate(data, paidAmountDpLc, paidAmountDpFc);
        }

        private static void CalculateWaightedRate(decimal totalInvFc, /*ref bool isRateCalculated,*/ ref string globalRate,
            decimal ratInv, ref decimal paidAmountDpLc, decimal paidAmountDpFc)
        {
            if (totalInvFc == paidAmountDpFc)
            {
                var rate = paidAmountDpLc / totalInvFc;
                globalRate = rate.ToString(CultureInfo.InvariantCulture);
                //isRateCalculated = true;
            }
            else if (totalInvFc > paidAmountDpFc)
            {
                var dif = (totalInvFc - paidAmountDpFc) * ratInv; //invocie Open AmountFC
                paidAmountDpLc += dif;
                var rate = paidAmountDpLc / totalInvFc;
                //isRateCalculated = true;
                globalRate = Math.Round(rate, 6).ToString();
            }
        }
        private static void CalculateWaightedRate(DataForCalculationRate data,
             decimal paidAmountDpLc, decimal paidAmountDpFc)
        {
            if (data.TotalInv == paidAmountDpFc)
            {
                var rate = paidAmountDpLc / data.TotalInv;
                data.GlobalRate = rate.ToString(CultureInfo.InvariantCulture);
                data.IsCalculated = true;
            }
            else if (data.TotalInv > paidAmountDpFc)
            {
                var dif = (data.TotalInv - paidAmountDpFc) * data.RateInv; //invocie Open AmountFC
                paidAmountDpLc += dif;
                var rate = paidAmountDpLc / data.TotalInv;
                data.IsCalculated = true;
                //isRateCalculated = true;
                data.GlobalRate = Math.Round(rate, 6).ToString(CultureInfo.InvariantCulture);
            }
        }

        public static decimal ExchangeRateCorrectionDi(decimal netAmountToDrow, decimal totalInv, decimal ratInv,
            int downPaymentDocEntry, string docCurrency, SAPbobsCOM.Company _comp)
        {
            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net AmountFC To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net AmountFC To Drow 

            var recSetTransferDocEntry =
                (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            var recSerTranferRate =
                (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);

            recSetTransferDocEntry.DoQuery(
                "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum " +
                "where RCT2.DocEntry = '" + downPaymentDocEntry + "' and InvType = 203 and ORCT.Canceled = 'N'");

            decimal sendRate = 0;
            if (recSetTransferDocEntry.RecordCount == 0)
            {
                return 0;
            }
            if (recSetTransferDocEntry.RecordCount == 1)
            {
                var orctDocEntry = recSetTransferDocEntry.Fields.Item("DocEntry").Value.ToString();

                recSerTranferRate.DoQuery(
                    "select ORCT.TrsfrSum , RCT2.AppliedFC, RCT2.DocRate from ORCT inner join RCT2 on " +
                    "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry = '" + orctDocEntry + "'");

                decimal transferSumLc = decimal.Parse(recSerTranferRate.Fields.Item("TrsfrSum").Value.ToString());

                decimal appliedAmountFcSum = 0;

                while (!recSerTranferRate.EoF)
                {
                    appliedAmountFcSum += decimal.Parse(recSerTranferRate.Fields.Item("AppliedFC").Value.ToString());
                    recSerTranferRate.MoveNext();
                }

                if (appliedAmountFcSum == 0)
                {
                    return 0;
                }

                sendRate = transferSumLc / appliedAmountFcSum;
            }

            else
            {
                Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                string orctDocEntrys = "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                                       downPaymentDocEntry +
                                       "' and InvType = 203 and ORCT.Canceled = 'N'";

                recSet2.DoQuery(
                    "select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum' , SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                    "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + orctDocEntrys +
                    ") group by ORCT.DocEntry");
                // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
                while (!recSet2.EoF)
                {
                    int ocrtDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                    decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
                    decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                    sumPayments.Add(new Tuple<int, decimal, decimal>(ocrtDocEntry, appliedAmountLc, appliedAmountFc));
                    recSet2.MoveNext();
                }

                //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet4.DoQuery(
                    "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  ORCT.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then RCT2.SumApplied else 0 end ) as 'LcPrices' from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" +
                    orctDocEntrys +
                    ") group by  RCT2.SumApplied , ORCT.DocEntry ) LcPricesTable group by DocEntry");

                Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
                while (!recSet4.EoF)
                {
                    string ocrtDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                    decimal sumLcPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                    DocumentLcPriceSums.Add(ocrtDocEntry, sumLcPayments);
                    recSet4.MoveNext();
                    // აქ არის ლოკალურ ვალუტაში გატარებულ დოკუმენტებზე გადახდილი სტრული თანხა  
                }

                Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

                List<XContainer> docsWithRateAndValue = new List<XContainer>();

                foreach (var tuple in sumPayments)
                {
                    var rate = (tuple.Item2 - DocumentLcPriceSums[tuple.Item1.ToString()]) / tuple.Item3;
                    var paymentDocEntry = tuple.Item1.ToString();

                    rateByDocuments.Add(paymentDocEntry, rate);

                    docsWithRateAndValue.Add(new XContainer()
                    {
                        CurrRate = rate,
                        OrctDocEntry = paymentDocEntry
                    });
                }


                Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet3.DoQuery(
                    "select ORCT.DocEntry, RCT2.DocEntry as 'DpDocEntry',    RCT2.AppliedFC from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry  in ( " +
                    orctDocEntrys + ") and RCT2.DocEntry = '" + downPaymentDocEntry + "' and InvType = 203");


                Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                while (!recSet3.EoF)
                {
                    decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                    string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                    dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

                    docsWithRateAndValue.Where(z => z.OrctDocEntry == PaymentDocEntry).ToList()
                        .ForEach(s => s.AmountFC = AppliedFcbyDp);
                    recSet3.MoveNext();
                }

                decimal lcSum = docsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                decimal fcSum = docsWithRateAndValue.Sum(doc => doc.AmountFC);
                sendRate = lcSum / fcSum;
            }

            paidAmountDpLc += sendRate * netAmountToDrow;
            paidAmountDpFc += netAmountToDrow;
            string globalRate = string.Empty;
            CalculateWaightedRate(totalInv, /*ref x,*/ ref globalRate, ratInv, ref paidAmountDpLc, paidAmountDpFc);
            return decimal.Parse(globalRate);
        }

        public static void AddJournalEntryCreditAp(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235, string vatAccount = "", double vatAmount = 0, string vatGroup = "")
        {

            JournalEntries vJE =
                (JournalEntries)_comp.GetBusinessObject(BoObjectTypes.oJournalEntries);

            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;

            vJE.Memo = "Income Correction -   Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;
            vJE.TransactionCode = "1";
            vJE.Series = series;

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Debit = amount;
            vJE.Lines.Credit = 0;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.ControlAccount = creditCode;
            vJE.Lines.ShortName = code;
            vJE.Lines.Debit = -amount;
            vJE.Lines.Credit = 0;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            if (!string.IsNullOrWhiteSpace(vatGroup))
            {
                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = debitCode;
                vJE.Lines.Debit = 0;
                vJE.Lines.Credit = vatAmount;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();

                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = vatAccount;
                vJE.Lines.TaxGroup = vatGroup;
                vJE.Lines.Debit = vatAmount;
                vJE.Lines.Credit = 0;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();
            }

            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            var x = vJE.GetAsXML();
            string des = _comp.GetLastErrorDescription();
            Application.SBO_Application.MessageBox(des + " Invoice " + reference);
        }


        public static void AddJournalEntryCreditAr(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235, string vatAccount = "", double vatAmount = 0, string vatGroup = "")
        {

            JournalEntries vJE =
                (JournalEntries)_comp.GetBusinessObject(BoObjectTypes.oJournalEntries);

            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;

            vJE.Memo = "Income Correction -   Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;
            vJE.TransactionCode = "1";
            vJE.Series = series;

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Debit = amount;
            vJE.Lines.Credit = 0;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.ControlAccount = creditCode;
            vJE.Lines.ShortName = code;
            vJE.Lines.Debit = -amount;
            vJE.Lines.Credit = 0;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            if (!string.IsNullOrWhiteSpace(vatGroup))
            {
                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = debitCode;
                vJE.Lines.Debit = vatAmount;
                vJE.Lines.Credit = 0;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();

                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = vatAccount;
                vJE.Lines.TaxGroup = vatGroup;
                vJE.Lines.Debit = 0;
                vJE.Lines.Credit = vatAmount;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();
            }

            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            var x = vJE.GetAsXML();
            string des = _comp.GetLastErrorDescription();
            Application.SBO_Application.MessageBox(des + " Invoice " + reference);
        }




        public static void AddJournalEntryDebitAp(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string shortName, DateTime DocDate, int BPLID = 235, string vatAccount = "", double vatAmount = 0, string vatGroup = "")
        {
            JournalEntries vJE =
                (JournalEntries)_comp.GetBusinessObject(BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = "Income Correction -   Invoice " + reference;
            vJE.Reference = reference;
            vJE.Series = series;
            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Credit = -amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.ShortName = shortName;
            vJE.Lines.Credit = amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            if (!string.IsNullOrWhiteSpace(vatGroup))
            {
                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = creditCode;
                vJE.Lines.Debit = -vatAmount;
                vJE.Lines.Credit = 0;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();

                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = vatAccount;
                vJE.Lines.TaxGroup = vatGroup;
                vJE.Lines.Debit = 0;
                vJE.Lines.Credit = -vatAmount;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();
            }

            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            string des = _comp.GetLastErrorDescription();
            Application.SBO_Application.MessageBox(des + " Invoice " + reference);
            var x = vJE.GetAsXML();
        }


        public static void AddJournalEntryDebitAr(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string shortName, DateTime DocDate, int BPLID = 235, string vatAccount = "", double vatAmount = 0, string vatGroup = "")
        {
            JournalEntries vJE =
                (JournalEntries)_comp.GetBusinessObject(BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = "Income Correction -   Invoice " + reference;
            vJE.Reference = reference;
            vJE.Series = series;
            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Credit = -amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.ShortName = shortName;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.Credit = amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            if (!string.IsNullOrWhiteSpace(vatGroup))
            {
                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = debitCode;
                vJE.Lines.Debit = 0;
                vJE.Lines.Credit = -vatAmount;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();

                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = vatAccount;
                vJE.Lines.TaxGroup = vatGroup;
                vJE.Lines.Debit = -vatAmount;
                vJE.Lines.Credit = 0;
                vJE.Lines.FCCredit = 0;
                vJE.Lines.FCDebit = 0;
                vJE.Lines.Add();
            }

            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            string des = _comp.GetLastErrorDescription();
            Application.SBO_Application.MessageBox(des + " Invoice " + reference);
            var x = vJE.GetAsXML();
        }



        private static Dictionary<int, string> FormToTransId = new Dictionary<int, string>()
        {
            {133, "13"},
            {141, "18"}
        };

        public static void CorrectionJournalEntryUI(SAPbobsCOM.Company _comp, int FormType, string businesPartnerCardCode, string applied, string docNumber, string bplName, string ExchangeGain, string ExchangeLoss, DateTime docDate)
        {

            string vatAccountDownPayment = string.Empty;
            string vatGroup = string.Empty;
            decimal vatDifferenceBetweenDpmInv = 0;
            decimal vatAmountInvDpm = 0;
            decimal exchangeRateAmount = 0;
            decimal invTransitAmount = 0;
            decimal correctionAmount = 0;
            Recordset recSet = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetNewAp = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetNewAr = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetJdtAp = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetJdtAr = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetDeter = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);

            recSetDeter.DoQuery($"select * from OACP where Year(OACP.FinancYear) =  {DateTime.Now.Year}");
            string transitAccountaP = recSetDeter.Fields.Item("PurcVatOff").Value.ToString();
            string transitAccountaR = recSetDeter.Fields.Item("SaleVatOff").Value.ToString();
            string dpmControlAcc = recSetDeter.Fields.Item("VDownPymnt").Value.ToString();
            string dpmControlAccaR = recSetDeter.Fields.Item("CDownPymnt").Value.ToString();


            recSetNewAp.DoQuery($"select * from PCH9 where DocEntry = {docNumber}");
            recSetNewAr.DoQuery($"select * from INV9 where DocEntry = {docNumber}");

            List<KeyValuePair<string, string>> vatGroupAndAccount = new List<KeyValuePair<string, string>>();
            List<string> transides = new List<string>();

            while (!recSetNewAr.EoF)
            {
                var dpm = recSetNewAr.Fields.Item("BaseAbs").Value.ToString();

                recSetNewAr.DoQuery($"SELECT transid FROM ODPI WHERE DocEntry = {dpm}");
                var transid = recSetNewAr.Fields.Item("transid").Value;
                if (transid != null && transid.ToString() != "0")
                {
                    transides.Add(transid.ToString());
                    recSetJdtAp.DoQuery($"SELECT * FROM JDT1 WHERE TransId = {transid}");
                    while (!recSetJdtAp.EoF)
                    {
                        if (recSetJdtAp.Fields.Item("Account").Value.ToString() == transitAccountaP)
                        {
                            recSetJdtAp.MoveNext();
                        }
                        else if (recSetJdtAp.Fields.Item("Account").Value.ToString() == transitAccountaR)
                        {
                            recSetJdtAp.MoveNext();
                        }
                        else
                        {
                            vatAccountDownPayment = recSetJdtAp.Fields.Item("Account").Value.ToString();
                            vatGroup = recSetJdtAp.Fields.Item("VatGroup").Value.ToString();
                            vatGroupAndAccount.Add(new KeyValuePair<string, string>(vatGroup, vatAccountDownPayment));
                            break;
                        }
                    }
                }
                recSetNewAr.MoveNext();
            }

            while (!recSetNewAp.EoF)
            {
                var dpm = recSetNewAp.Fields.Item("BaseAbs").Value.ToString();

                recSetNewAp.DoQuery($"SELECT transid FROM ODPO WHERE DocEntry = {dpm}");
                var transid = recSetNewAp.Fields.Item("transid").Value;
                if (transid != null && transid.ToString() != "0")
                {
                    transides.Add(transid.ToString());
                    recSetJdtAr.DoQuery($"SELECT * FROM JDT1 WHERE TransId = {transid}");
                    while (!recSetJdtAr.EoF)
                    {
                        if (recSetJdtAr.Fields.Item("Account").Value.ToString() == transitAccountaP)
                        {
                            recSetJdtAr.MoveNext();
                        }
                        else if (recSetJdtAr.Fields.Item("Account").Value.ToString() == transitAccountaR)
                        {
                            recSetJdtAr.MoveNext();
                        }
                        else
                        {
                            vatAccountDownPayment = recSetJdtAr.Fields.Item("Account").Value.ToString();
                            vatGroup = recSetJdtAr.Fields.Item("VatGroup").Value.ToString();
                            vatGroupAndAccount.Add(new KeyValuePair<string, string>(vatGroup, vatAccountDownPayment));
                            break;
                        }
                    }
                }
                recSetNewAp.MoveNext();
            }

            recSet.DoQuery(@"SELECT DebPayAcct, 
            CRD3.AcctType,
            CRD3.AcctCode
                FROM OCRD
                LEFT JOIN CRD3 ON OCRD.CardCode = CRD3.CardCode where OCRD.CardCode = '" + businesPartnerCardCode + "'");

            string bpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();


            while (!recSet.EoF)
            {
                var acctType = recSet.Fields.Item("AcctType").Value.ToString();
                var acctCode = recSet.Fields.Item("AcctCode").Value.ToString();
                if (acctType == "D")
                {
                    dpmControlAcc = acctCode;
                }

                recSet.MoveNext();
            }


            if (!string.IsNullOrWhiteSpace(applied))
            {
                var objRS = (Recordset)(_comp).GetBusinessObject(BoObjectTypes.BoRecordset);
                objRS.DoQuery(@"select * from OJDT where baseRef = " + docNumber + " and TransType = " +
                              FormToTransId[FormType] + "");
                objRS.MoveFirst();
                var x = objRS.Fields.Item("TransType").Value.ToString();
                if (objRS.Fields.Item("TransType").Value.ToString() != "13" &&
                    objRS.Fields.Item("TransType").Value.ToString() != "18")
                {
                    objRS.MoveNext();
                }
                var transID = objRS.Fields.Item("TransId").Value.ToString();
                objRS.DoQuery(@"select * from JDT1 where TransId = " + transID);
                var objRS234 = (Recordset)(_comp).GetBusinessObject(BoObjectTypes.BoRecordset);

                if (objRS.Fields.Item("TransType").Value.ToString() != "13")
                {
                    objRS234.DoQuery("select  BPLId from OPCH where BPLName = N'" + bplName + "'");
                }
                else if (objRS.Fields.Item("TransType").Value.ToString() != "18")
                {
                    objRS234.DoQuery("select  BPLId from OINV where BPLName = N'" + bplName + "'");
                }

                int bplID = Convert.ToInt32(objRS234.Fields.Item("BPLId").Value);


                decimal dpmTotalWeightedRateInvAp = 0;
                decimal dpmTotalOriginalRateAp = 0;
                decimal dpmVatTotalOriginalRateAp = 0;

                decimal dpmTotalWeightedRateInvAr = 0;
                decimal dpmTotalOriginalRateAr = 0;
                decimal dpmVatTotalOriginalRateAr = 0;

                while (!objRS.EoF)
                {
                    bool difCalculated = false;
                    bool exchangeRateCalculated = false;
                    bool transitCalculated = false;
                    string invAcc = objRS.Fields.Item("Account").Value.ToString();
                    string invvatGroup = objRS.Fields.Item("VatGroup").Value.ToString();
                    decimal invCreditFc = decimal.Parse(objRS.Fields.Item("FCCredit").Value.ToString());
                    decimal invCredit = decimal.Parse(objRS.Fields.Item("Credit").Value.ToString());
                    decimal invDebit = decimal.Parse(objRS.Fields.Item("Debit").Value.ToString());
                    var defauKeyValuePair = default(KeyValuePair<string, string>);

                    if (invAcc == bpControlAcc && invDebit != 0)
                    {
                        dpmTotalWeightedRateInvAp = invDebit;
                    }

                    if (invAcc == dpmControlAcc && invCredit != 0)
                    {
                        dpmTotalOriginalRateAp += invCredit;
                    }

                    if (vatGroupAndAccount.FirstOrDefault(acc => acc.Value == invAcc).Equals(defauKeyValuePair) && vatGroupAndAccount.FirstOrDefault(vatGr => vatGr.Key == invvatGroup).Equals(defauKeyValuePair))
                    {
                        dpmVatTotalOriginalRateAp += invCredit;
                    }

                    if (invAcc == ExchangeGain || invAcc == ExchangeLoss)
                    {
                        exchangeRateAmount = invCredit + invDebit;
                    }
                    ///////////////////////////////////////////////////////
                    if (invAcc == bpControlAcc && invCredit != 0)
                    {
                        dpmTotalWeightedRateInvAr = invCredit;
                    }

                    if (invAcc == dpmControlAccaR && invDebit != 0)
                    {
                        dpmTotalOriginalRateAr += invDebit;
                    }

                    if (vatGroupAndAccount.FirstOrDefault(acc => acc.Value == invAcc).Equals(defauKeyValuePair) && vatGroupAndAccount.FirstOrDefault(vatGr => vatGr.Key == invvatGroup).Equals(defauKeyValuePair))
                    {
                        dpmVatTotalOriginalRateAr += invDebit;
                    }

                    if (invAcc == ExchangeGain || invAcc == ExchangeLoss)
                    {
                        exchangeRateAmount = invCredit + invDebit;
                    }

                    objRS.MoveNext();
                }


                decimal receivibleCancellationAmountAp = dpmTotalWeightedRateInvAp - dpmTotalOriginalRateAp;
                decimal vatCencellationAmountGainAp = receivibleCancellationAmountAp - exchangeRateAmount;
                decimal vatCencellationAmountLossAp = receivibleCancellationAmountAp + exchangeRateAmount;

                decimal receivibleCancellationAmountAr = dpmTotalWeightedRateInvAr - dpmTotalOriginalRateAr;
                decimal vatCencellationAmountGainAr = receivibleCancellationAmountAr + exchangeRateAmount;
                decimal vatCencellationAmountLossAr = receivibleCancellationAmountAr - exchangeRateAmount;

                bool validation = dpmVatTotalOriginalRateAp == vatCencellationAmountGainAp;
                if (validation)
                {
                    Application.SBO_Application.MessageBox("გადაამოწმეთ შექმნილი გატარება");
                }

                Recordset recSet12 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet12.DoQuery("select Series from NNM1 where ObjectCode = 30 and Locked = 'N' and BPLId is  null");
                int series = int.Parse(recSet12.Fields.Item("Series").Value.ToString());

                if (bplID == 0)
                {
                    bplID = 235;
                }

                objRS.MoveFirst();
                while (!objRS.EoF)
                {
                    var account = objRS.Fields.Item("Account").Value.ToString();

                    if (FormType.ToString() == "133")
                    {
                        if (account == ExchangeGain)
                        {
                            AddJournalEntryDebitAr(_comp, bpControlAcc, ExchangeGain,
                                Convert.ToDouble(receivibleCancellationAmountAr), series, docNumber, businesPartnerCardCode, docDate,
                                bplID, vatGroupAndAccount.FirstOrDefault().Value, Convert.ToDouble(vatCencellationAmountGainAr), vatGroupAndAccount.FirstOrDefault().Key);

                        }
                        else if (account == ExchangeLoss)
                        {
                            AddJournalEntryCreditAr(_comp, bpControlAcc, ExchangeLoss,
                                Convert.ToDouble(receivibleCancellationAmountAr), series, docNumber, businesPartnerCardCode, docDate,
                                bplID, vatGroupAndAccount.FirstOrDefault().Value, Convert.ToDouble(vatCencellationAmountLossAr), vatGroupAndAccount.FirstOrDefault().Key);
                        }

                    }
                    else if (FormType.ToString() == "141")
                    {
                        if (account == ExchangeGain)
                        {
                            AddJournalEntryCreditAp(_comp, bpControlAcc, ExchangeGain,
                                Convert.ToDouble(receivibleCancellationAmountAp), series, docNumber, businesPartnerCardCode, docDate,
                                bplID, vatGroupAndAccount.FirstOrDefault().Value, Convert.ToDouble(vatCencellationAmountGainAp), vatGroupAndAccount.FirstOrDefault().Key);
                        }
                        else if (account == ExchangeLoss)
                        {
                            AddJournalEntryDebitAp(_comp, ExchangeLoss, bpControlAcc,
                                Convert.ToDouble(receivibleCancellationAmountAp), series, docNumber, businesPartnerCardCode, docDate, bplID,
                                 vatGroupAndAccount.FirstOrDefault().Value, Convert.ToDouble(vatCencellationAmountLossAp), vatGroupAndAccount.FirstOrDefault().Key);
                        }
                    }

                    objRS.MoveNext();
                }
            }

        }

        public static void CorrectionJournalEntryDi(SAPbobsCOM.Company _comp, int formType, string businesPartnerCardCode, string docNumber, string bplName, string exchangeGain, string ExchangeLoss, DateTime docDate)
        {
            Recordset recSet = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery("select DebPayAcct from OCRD where CardCode = '" + businesPartnerCardCode + "'");
            string bpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();

            var objRs = (Recordset)(_comp).GetBusinessObject(BoObjectTypes.BoRecordset);
            objRs.DoQuery(@"select * from OJDT where baseRef = " + docNumber + " and TransType = " +
                          FormToTransId[formType] + "");
            objRs.MoveFirst();
            if (objRs.Fields.Item("TransType").Value.ToString() != "13" &&
                objRs.Fields.Item("TransType").Value.ToString() != "18")
            {
                objRs.MoveNext();
            }
            var transId = objRs.Fields.Item("TransId").Value.ToString();
            objRs.DoQuery(@"select * from JDT1 where TransId = " + transId);
            var objRs234 = (Recordset)(_comp).GetBusinessObject(BoObjectTypes.BoRecordset);

            if (objRs.Fields.Item("TransType").Value.ToString() != "13")
            {
                objRs234.DoQuery("select  BPLId from OPCH where BPLName = '" + bplName + "'");
            }
            else if (objRs.Fields.Item("TransType").Value.ToString() != "18")
            {
                objRs234.DoQuery("select  BPLId from OINV where BPLName = '" + bplName + "'");
            }

            int bplId = Convert.ToInt32(objRs234.Fields.Item("BPLId").Value);
            Recordset recSet12 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet12.DoQuery("select Series from NNM1 where ObjectCode = 30 and Locked = 'N' and BPLId is  null");
            int series = int.Parse(recSet12.Fields.Item("Series").Value.ToString());

            if (bplId == 0)
            {
                bplId = 235;
            }

            while (!objRs.EoF)
            {
                var account = objRs.Fields.Item("Account").Value.ToString();

                if (formType.ToString() == "133")
                {
                    if (account == exchangeGain)
                    {
                        AddJournalEntryCreditAp(_comp, bpControlAcc, exchangeGain,
                            Convert.ToDouble(objRs.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                            bplId);
                    }
                    else if (account == ExchangeLoss)
                    {
                        AddJournalEntryDebitAp(_comp, ExchangeLoss, bpControlAcc,
                            Convert.ToDouble(objRs.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplId);
                    }
                }
                else if (formType.ToString() == "141")
                {
                    if (account == exchangeGain)
                    {
                        AddJournalEntryCreditAp(_comp, bpControlAcc, exchangeGain,
                            Convert.ToDouble(objRs.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                            bplId);
                    }
                    else if (account == ExchangeLoss)
                    {
                        AddJournalEntryDebitAp(_comp, ExchangeLoss, bpControlAcc,
                            Convert.ToDouble(objRs.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplId);
                    }
                }
                objRs.MoveNext();
            }
        }
    }
}
