using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using DownPaymentLogic.Classes;
using SAPbobsCOM;
using SAPbouiCOM;

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

        //Form downPaymentToDrow, SAPbobsCOM.Company _comp, string formType,
        //string docCurrency, decimal totalInv,    out string globalRate, decimal ratInv)
        {
            //isRateCalculated = false;
            //Form downPaymentToDrowForm = downPaymentToDrow;
            //Item downPaymentFormMatrix = downPaymentToDrowForm.Items.Item("6"); //Down Payment to drow
            //Matrix matrix = (SAPbouiCOM.Matrix)downPaymentFormMatrix.Specific;
            data.GlobalRate = "1,0000";

            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net AmountFC To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net AmountFC To Drow 

            foreach (var downpayment in data.GrossAmountsForDownPayment)
            {
                if (data.FormTypex == "133")
                {
                    //string ORCTDocEntrys = string.Empty;

                    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet2.DoQuery("SELECT DocEntry FROM ODPI WHERE DocNum = '" + downpayment.First().Key + "'");
                    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

                    string ORCTDocEntrys =
                        "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                        dpDocEntry + "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



                    recSet2.DoQuery("select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum' , SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                        "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + ORCTDocEntrys + ") group by ORCT.DocEntry");
                    // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                    List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();


                    if (recSet2.EoF)
                    {
                        recSet2.DoQuery($"SELECT BaseRef FROM DPI1 WHERE DocEntry = '{dpDocEntry}'");
                        dpDocEntry = recSet2.Fields.Item("BaseRef").Value.ToString();

                        ORCTDocEntrys =
                            "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                            dpDocEntry + "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები
                        recSet2.DoQuery("select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum' , SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                                        "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + ORCTDocEntrys + ") group by ORCT.DocEntry");
                    }

                    while (!recSet2.EoF)
                    {
                        int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                        decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
                        decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                        sumPayments.Add(new Tuple<int, decimal, decimal>(OCRTDocEntry, appliedAmountLc, appliedAmountFc));
                        recSet2.MoveNext();
                    }


                    string docEntrysJounralEntrysQuery =
                        "select  RCT2.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum inner join OJDT on RCT2.DocTransId = OJDT.TransId where ORCT.DocEntry in (" +
                        ORCTDocEntrys + ") and OJDT.TransCode = 2 and RCT2.InvType = 30";// ისეტი საჯურნალო გატარებები რომლებსაც აქვს ტრანზაქციის კოდი 2 და არ უდნა დაიტვალოს კურსის დაანგარიშებისას

                    Recordset recSetLast = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSetLast.DoQuery(docEntrysJounralEntrysQuery);

                    string docEntrysJounralEntrys = string.Empty;
                    while (!recSetLast.EoF)
                    {
                        docEntrysJounralEntrys += recSetLast.Fields.Item("DocEntry").Value + ",";
                        recSetLast.MoveNext();
                    }
                    docEntrysJounralEntrys = docEntrysJounralEntrys.Remove(docEntrysJounralEntrys.Length - 1, 1);


                    //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                    Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet4.DoQuery(
                        "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  ORCT.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 and RCT2.DocEntry not in (" + docEntrysJounralEntrys + ") then RCT2.SumApplied else 0 end ) as 'LcPrices' from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" +
                        ORCTDocEntrys +
                        ") group by  RCT2.SumApplied , ORCT.DocEntry ) LcPricesTable group by DocEntry");

                    Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
                    while (!recSet4.EoF)
                    {
                        string OCRTDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                        decimal SumLCPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                        DocumentLcPriceSums.Add(OCRTDocEntry, SumLCPayments);
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

                        DocsWithRateAndValue.Add(new XContainer()
                        {
                            CurrRate = rate,
                            OrctDocEntry = paymentDocEntry
                        });
                    }
                    // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში


                    //while (!recSet2.EoF)
                    //{
                    //    decimal rate = (decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) - DocumentLcPriceSums[recSet2.Fields.Item("DocEntry").Value.ToString()]) /
                    //            decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                    //    string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                    //    rateByDocuments.Add(paymentDocEntry, rate);
                    //    recSet2.MoveNext();

                    //}

                    Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet3.DoQuery("select ORCT.DocEntry, RCT2.DocEntry as 'DpDocEntry',    RCT2.AppliedFC from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry  in ( " + ORCTDocEntrys + ") and RCT2.DocEntry = '" + dpDocEntry + "' and InvType = 203");


                    Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                    while (!recSet3.EoF)
                    {
                        decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                        string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                        dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

                        DocsWithRateAndValue.Where(x => x.OrctDocEntry == PaymentDocEntry).ToList().ForEach(s => s.AmountFC = AppliedFcbyDp);


                        recSet3.MoveNext();
                    }

                    //var rata = DocsWithRateAndValue.

                    decimal LCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                    decimal FCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
                    decimal WeightedRate = LCSum / FCSum;

                    paidAmountDpLc += decimal.Parse(downpayment.First().Value) * WeightedRate;
                    Console.WriteLine();

                    //#region ჩემი ლოგიკა

                    //Dictionary<string, decimal> amountRatio = new Dictionary<string, decimal>(); // თანხების ფარდობა გადასხდისთვის


                    //foreach (KeyValuePair<string, decimal> variable in dPIncomingPaymentShareAmountFc)
                    //{
                    //    decimal division = variable.Value / dPIncomingPaymentShareAmountFc.Sum(x => x.Value);
                    //    amountRatio.Add(variable.Key, division);
                    //}


                    //Dictionary<string, decimal> docEntryPaymentFcRatioFromNetAmount = new Dictionary<string, decimal>();

                    //foreach (var VARIABLE in amountRatio)
                    //{
                    //    decimal fcRatioFromNetAmount = VARIABLE.Value * decimal.Parse(netAmountToDrow);
                    //    string paymentDocEntry = VARIABLE.Key;
                    //    docEntryPaymentFcRatioFromNetAmount.Add(paymentDocEntry, fcRatioFromNetAmount);
                    //}

                    //foreach (var VARIABLE in docEntryPaymentFcRatioFromNetAmount)
                    //{
                    //    paidAmountDpLc += VARIABLE.Value * rateByDocuments[VARIABLE.Key];
                    //}

                    ////აქ ვამრავლევ ავანსის ტანხის წილს თავის შესაბამისს კურსზე 
                    //#endregion

                    #region იმათი ლოგიკა



                    #endregion

                }
                else
                {


                    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + downpayment.First().Key + "'");
                    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

                    string ORCTDocEntrys =
                        "select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '" +
                        dpDocEntry + "' and InvType = 204 and OVPM.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



                    recSet2.DoQuery("select OVPM.DocEntry, avg(OVPM.TrsfrSum) as 'TrsfrSum' , SUM(VPM2.AppliedFC) as 'AppliedFC' from OVPM inner join VPM2 on " +
                                    "OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" + ORCTDocEntrys + ") group by OVPM.DocEntry");
                    // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                    if (recSet2.EoF)
                    {
                        recSet2.DoQuery($"SELECT BaseRef FROM DPO1 WHERE DocEntry = '{dpDocEntry}'");
                        dpDocEntry = recSet2.Fields.Item("BaseRef").Value.ToString();

                        ORCTDocEntrys =
                            $"select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '{dpDocEntry}' and InvType = 204 and OVPM.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები
                        recSet2.DoQuery("select OVPM.DocEntry, avg(OVPM.TrsfrSum) as 'TrsfrSum' , SUM(VPM2.AppliedFC) as 'AppliedFC' from OVPM inner join VPM2 on " +
                                        "OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" + ORCTDocEntrys + ") group by OVPM.DocEntry");
                    }




                    List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
                    while (!recSet2.EoF)
                    {
                        int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                        decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
                        decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                        sumPayments.Add(new Tuple<int, decimal, decimal>(OCRTDocEntry, appliedAmountLc, appliedAmountFc));
                        recSet2.MoveNext();
                    }








                    //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                    Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet4.DoQuery(
                        "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  OVPM.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then VPM2.SumApplied else 0 end ) as 'LcPrices' from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" +
                        ORCTDocEntrys +
                        ") group by  VPM2.SumApplied , OVPM.DocEntry ) LcPricesTable group by DocEntry");

                    Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
                    while (!recSet4.EoF)
                    {
                        string OCRTDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                        decimal SumLCPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                        DocumentLcPriceSums.Add(OCRTDocEntry, SumLCPayments);
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

                        DocsWithRateAndValue.Add(new XContainer()
                        {
                            CurrRate = rate,
                            OrctDocEntry = paymentDocEntry
                        });
                    }
                    // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში


                    //while (!recSet2.EoF)
                    //{
                    //    decimal rate = (decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) - DocumentLcPriceSums[recSet2.Fields.Item("DocEntry").Value.ToString()]) /
                    //            decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                    //    string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                    //    rateByDocuments.Add(paymentDocEntry, rate);
                    //    recSet2.MoveNext();

                    //}

                    Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet3.DoQuery("select OVPM.DocEntry, VPM2.DocEntry as 'DpDocEntry',    VPM2.AppliedFC from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry  in ( " + ORCTDocEntrys + ") and VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204");


                    Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                    while (!recSet3.EoF)
                    {
                        decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                        string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                        dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

                        DocsWithRateAndValue.Where(x => x.OrctDocEntry == PaymentDocEntry).ToList().ForEach(s => s.AmountFC = AppliedFcbyDp);


                        recSet3.MoveNext();
                    }

                    //var rata = DocsWithRateAndValue.

                    decimal LCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                    decimal FCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
                    decimal WeightedRate = LCSum / FCSum;

                    paidAmountDpLc += decimal.Parse(downpayment.First().Value) * WeightedRate;
                    Console.WriteLine();

                    //#region ჩემი ლოგიკა

                    //Dictionary<string, decimal> amountRatio = new Dictionary<string, decimal>(); // თანხების ფარდობა გადასხდისთვის


                    //foreach (KeyValuePair<string, decimal> variable in dPIncomingPaymentShareAmountFc)
                    //{
                    //    decimal division = variable.Value / dPIncomingPaymentShareAmountFc.Sum(x => x.Value);
                    //    amountRatio.Add(variable.Key, division);
                    //}


                    //Dictionary<string, decimal> docEntryPaymentFcRatioFromNetAmount = new Dictionary<string, decimal>();

                    //foreach (var VARIABLE in amountRatio)
                    //{
                    //    decimal fcRatioFromNetAmount = VARIABLE.Value * decimal.Parse(netAmountToDrow);
                    //    string paymentDocEntry = VARIABLE.Key;
                    //    docEntryPaymentFcRatioFromNetAmount.Add(paymentDocEntry, fcRatioFromNetAmount);
                    //}

                    //foreach (var VARIABLE in docEntryPaymentFcRatioFromNetAmount)
                    //{
                    //    paidAmountDpLc += VARIABLE.Value * rateByDocuments[VARIABLE.Key];
                    //}

                    ////აქ ვამრავლევ ავანსის ტანხის წილს თავის შესაბამისს კურსზე 
                    //#endregion

                    #region იმათი ლოგიკა



                    #endregion

                }
                //{
                //    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                //    recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + txtID.Value + "'");
                //    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                //    objMD.DoQuery(
                //        "select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204 and OVPM.Canceled = 'N'");

                //    string OVPMDocEntrys = string.Empty;
                //    while (!recSet2.EoF)
                //    {
                //        OVPMDocEntrys += " OVPM.DocEntry = '" + recSet2.Fields.Item("DocEntry").Value + "'" + " OR ";
                //        recSet2.MoveNext();
                //    }

                //    OVPMDocEntrys = OVPMDocEntrys.Remove(OVPMDocEntrys.Length - 3, 3);

                //    recSet2.DoQuery("select OVPM.TrsfrSum , VPM2.AppliedFC from OVPM inner join VPM2 on " +
                //                    "OVPM.DocEntry = VPM2.DocNum where  '" + OVPMDocEntrys + "'");

                //    Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

                //    while (!recSet2.EoF)
                //    {
                //        decimal rate = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) /
                //                       decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                //        string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                //        rateByDocuments.Add(paymentDocEntry, rate);
                //        recSet2.MoveNext();
                //    }

                //    Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                //    recSet3.DoQuery("select OVPM.DocEntry, VPM2.DocEntry as 'DpDocEntry',    VPM2.AppliedFC from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where ( ORCTDocEntrys) and VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204");

                //    while (!recSet3.EoF)
                //    {
                //        paidAmountDpLc += decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString()) *
                //                          rateByDocuments[recSet3.Fields.Item("AppliedFC").Value.ToString()];
                //        recSet3.MoveNext();
                //    }

                //}

                try
                {
                    paidAmountDpFc += decimal.Parse(downpayment.First().Value);
                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            //}


            //}

            //for (int i = 1; i <= matrix.RowCount; i++)
            //{
            //    var checkbox = (SAPbouiCOM.CheckBox)matrix.Columns.Item("380000138").Cells.Item(i).Specific;
            //    if (checkbox.Checked)
            //    {
            //        EditText txtMoney =
            //            (SAPbouiCOM.EditText)matrix.Columns.Item("29").Cells.Item(i)
            //                .Specific; //net amount to drow//TODO
            //        EditText txtID = (SAPbouiCOM.EditText)matrix.Columns.Item("68").Cells.Item(i).Specific; //docNumber
            //        string netAmountToDrow = txtMoney.Value.Split(' ')[0]; //net amount to drow

            //        var objMD = (SAPbobsCOM.Recordset)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            //        if (formType == "133")
            //        {
            //            //string ORCTDocEntrys = string.Empty;

            //            Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //            recSet2.DoQuery("SELECT DocEntry FROM ODPI WHERE DocNum = '" + txtID.Value + "'");
            //            var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

            //            string ORCTDocEntrys =
            //                "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
            //                dpDocEntry + "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



            //            recSet2.DoQuery("select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum' , SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
            //                "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + ORCTDocEntrys + ") group by ORCT.DocEntry");
            //            // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

            //            List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
            //            while (!recSet2.EoF)
            //            {
            //                int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
            //                decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
            //                decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
            //                sumPayments.Add(new Tuple<int, decimal, decimal>(OCRTDocEntry, appliedAmountLc, appliedAmountFc));
            //                recSet2.MoveNext();
            //            }








            //            //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

            //            Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //            recSet4.DoQuery(
            //                "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  ORCT.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then RCT2.SumApplied else 0 end ) as 'LcPrices' from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" +
            //                ORCTDocEntrys +
            //                ") group by  RCT2.SumApplied , ORCT.DocEntry ) LcPricesTable group by DocEntry");

            //            Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
            //            while (!recSet4.EoF)
            //            {
            //                string OCRTDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
            //                decimal SumLCPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
            //                DocumentLcPriceSums.Add(OCRTDocEntry, SumLCPayments);
            //                recSet4.MoveNext();

            //                // აქ არის ლოკალურ ვალუტაში გატარებულ დოკუმენტებზე გადახდილი სტრული თანხა  
            //            }


            //            Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

            //            List<XContainer> DocsWithRateAndValue = new List<XContainer>();

            //            foreach (var tuple in sumPayments)
            //            {
            //                var rate = (tuple.Item2 - DocumentLcPriceSums[tuple.Item1.ToString()]) / tuple.Item3;
            //                var paymentDocEntry = tuple.Item1.ToString();

            //                rateByDocuments.Add(paymentDocEntry, rate);

            //                DocsWithRateAndValue.Add(new XContainer()
            //                {
            //                    CurrRate = rate,
            //                    OrctDocEntry = paymentDocEntry
            //                });
            //            }
            //            // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში


            //            //while (!recSet2.EoF)
            //            //{
            //            //    decimal rate = (decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) - DocumentLcPriceSums[recSet2.Fields.Item("DocEntry").Value.ToString()]) /
            //            //            decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
            //            //    string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
            //            //    rateByDocuments.Add(paymentDocEntry, rate);
            //            //    recSet2.MoveNext();

            //            //}

            //            Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //            recSet3.DoQuery("select ORCT.DocEntry, RCT2.DocEntry as 'DpDocEntry',    RCT2.AppliedFC from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry  in ( " + ORCTDocEntrys + ") and RCT2.DocEntry = '" + dpDocEntry + "' and InvType = 203");


            //            Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
            //            while (!recSet3.EoF)
            //            {
            //                decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
            //                string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
            //                dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

            //                DocsWithRateAndValue.Where(x => x.OrctDocEntry == PaymentDocEntry).ToList().ForEach(s => s.AmountFC = AppliedFcbyDp);


            //                recSet3.MoveNext();
            //            }

            //            //var rata = DocsWithRateAndValue.

            //            decimal LCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
            //            decimal FCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
            //            decimal WeightedRate = LCSum / FCSum;

            //            paidAmountDpLc += decimal.Parse(netAmountToDrow) * WeightedRate;
            //            Console.WriteLine();

            //            //#region ჩემი ლოგიკა

            //            //Dictionary<string, decimal> amountRatio = new Dictionary<string, decimal>(); // თანხების ფარდობა გადასხდისთვის


            //            //foreach (KeyValuePair<string, decimal> variable in dPIncomingPaymentShareAmountFc)
            //            //{
            //            //    decimal division = variable.Value / dPIncomingPaymentShareAmountFc.Sum(x => x.Value);
            //            //    amountRatio.Add(variable.Key, division);
            //            //}


            //            //Dictionary<string, decimal> docEntryPaymentFcRatioFromNetAmount = new Dictionary<string, decimal>();

            //            //foreach (var VARIABLE in amountRatio)
            //            //{
            //            //    decimal fcRatioFromNetAmount = VARIABLE.Value * decimal.Parse(netAmountToDrow);
            //            //    string paymentDocEntry = VARIABLE.Key;
            //            //    docEntryPaymentFcRatioFromNetAmount.Add(paymentDocEntry, fcRatioFromNetAmount);
            //            //}

            //            //foreach (var VARIABLE in docEntryPaymentFcRatioFromNetAmount)
            //            //{
            //            //    paidAmountDpLc += VARIABLE.Value * rateByDocuments[VARIABLE.Key];
            //            //}

            //            ////აქ ვამრავლევ ავანსის ტანხის წილს თავის შესაბამისს კურსზე 
            //            //#endregion

            //            #region იმათი ლოგიკა



            //            #endregion

            //        }


            //        else
            //        {
            //            //string ORCTDocEntrys = string.Empty;

            //            Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //            recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + txtID.Value + "'");
            //            var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();

            //            string ORCTDocEntrys =
            //                "select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '" +
            //                dpDocEntry + "' and InvType = 204 and OVPM.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



            //            recSet2.DoQuery("select OVPM.DocEntry, avg(OVPM.TrsfrSum) as 'TrsfrSum' , SUM(VPM2.AppliedFC) as 'AppliedFC' from OVPM inner join VPM2 on " +
            //                            "OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" + ORCTDocEntrys + ") group by OVPM.DocEntry");
            //            // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

            //            List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
            //            while (!recSet2.EoF)
            //            {
            //                int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
            //                decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
            //                decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
            //                sumPayments.Add(new Tuple<int, decimal, decimal>(OCRTDocEntry, appliedAmountLc, appliedAmountFc));
            //                recSet2.MoveNext();
            //            }








            //            //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

            //            Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //            recSet4.DoQuery(
            //                "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  OVPM.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then VPM2.SumApplied else 0 end ) as 'LcPrices' from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry in (" +
            //                ORCTDocEntrys +
            //                ") group by  VPM2.SumApplied , OVPM.DocEntry ) LcPricesTable group by DocEntry");

            //            Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
            //            while (!recSet4.EoF)
            //            {
            //                string OCRTDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
            //                decimal SumLCPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
            //                DocumentLcPriceSums.Add(OCRTDocEntry, SumLCPayments);
            //                recSet4.MoveNext();

            //                // აქ არის ლოკალურ ვალუტაში გატარებულ დოკუმენტებზე გადახდილი სტრული თანხა  
            //            }


            //            Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

            //            List<XContainer> DocsWithRateAndValue = new List<XContainer>();

            //            foreach (var tuple in sumPayments)
            //            {
            //                var rate = (tuple.Item2 - DocumentLcPriceSums[tuple.Item1.ToString()]) / tuple.Item3;
            //                var paymentDocEntry = tuple.Item1.ToString();

            //                rateByDocuments.Add(paymentDocEntry, rate);

            //                DocsWithRateAndValue.Add(new XContainer()
            //                {
            //                    CurrRate = rate,
            //                    OrctDocEntry = paymentDocEntry
            //                });
            //            }
            //            // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში


            //            //while (!recSet2.EoF)
            //            //{
            //            //    decimal rate = (decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) - DocumentLcPriceSums[recSet2.Fields.Item("DocEntry").Value.ToString()]) /
            //            //            decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
            //            //    string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
            //            //    rateByDocuments.Add(paymentDocEntry, rate);
            //            //    recSet2.MoveNext();

            //            //}

            //            Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //            recSet3.DoQuery("select OVPM.DocEntry, VPM2.DocEntry as 'DpDocEntry',    VPM2.AppliedFC from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where OVPM.DocEntry  in ( " + ORCTDocEntrys + ") and VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204");


            //            Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
            //            while (!recSet3.EoF)
            //            {
            //                decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
            //                string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
            //                dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

            //                DocsWithRateAndValue.Where(x => x.OrctDocEntry == PaymentDocEntry).ToList().ForEach(s => s.AmountFC = AppliedFcbyDp);


            //                recSet3.MoveNext();
            //            }

            //            //var rata = DocsWithRateAndValue.

            //            decimal LCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
            //            decimal FCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
            //            decimal WeightedRate = LCSum / FCSum;

            //            paidAmountDpLc += decimal.Parse(netAmountToDrow) * WeightedRate;
            //            Console.WriteLine();

            //            //#region ჩემი ლოგიკა

            //            //Dictionary<string, decimal> amountRatio = new Dictionary<string, decimal>(); // თანხების ფარდობა გადასხდისთვის


            //            //foreach (KeyValuePair<string, decimal> variable in dPIncomingPaymentShareAmountFc)
            //            //{
            //            //    decimal division = variable.Value / dPIncomingPaymentShareAmountFc.Sum(x => x.Value);
            //            //    amountRatio.Add(variable.Key, division);
            //            //}


            //            //Dictionary<string, decimal> docEntryPaymentFcRatioFromNetAmount = new Dictionary<string, decimal>();

            //            //foreach (var VARIABLE in amountRatio)
            //            //{
            //            //    decimal fcRatioFromNetAmount = VARIABLE.Value * decimal.Parse(netAmountToDrow);
            //            //    string paymentDocEntry = VARIABLE.Key;
            //            //    docEntryPaymentFcRatioFromNetAmount.Add(paymentDocEntry, fcRatioFromNetAmount);
            //            //}

            //            //foreach (var VARIABLE in docEntryPaymentFcRatioFromNetAmount)
            //            //{
            //            //    paidAmountDpLc += VARIABLE.Value * rateByDocuments[VARIABLE.Key];
            //            //}

            //            ////აქ ვამრავლევ ავანსის ტანხის წილს თავის შესაბამისს კურსზე 
            //            //#endregion

            //            #region იმათი ლოგიკა



            //            #endregion

            //        }
            //        //{
            //        //    Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //        //    recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + txtID.Value + "'");
            //        //    var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
            //        //    objMD.DoQuery(
            //        //        "select OVPM.DocEntry from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204 and OVPM.Canceled = 'N'");

            //        //    string OVPMDocEntrys = string.Empty;
            //        //    while (!recSet2.EoF)
            //        //    {
            //        //        OVPMDocEntrys += " OVPM.DocEntry = '" + recSet2.Fields.Item("DocEntry").Value + "'" + " OR ";
            //        //        recSet2.MoveNext();
            //        //    }

            //        //    OVPMDocEntrys = OVPMDocEntrys.Remove(OVPMDocEntrys.Length - 3, 3);

            //        //    recSet2.DoQuery("select OVPM.TrsfrSum , VPM2.AppliedFC from OVPM inner join VPM2 on " +
            //        //                    "OVPM.DocEntry = VPM2.DocNum where  '" + OVPMDocEntrys + "'");

            //        //    Dictionary<string, decimal> rateByDocuments = new Dictionary<string, decimal>();

            //        //    while (!recSet2.EoF)
            //        //    {
            //        //        decimal rate = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) /
            //        //                       decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
            //        //        string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
            //        //        rateByDocuments.Add(paymentDocEntry, rate);
            //        //        recSet2.MoveNext();
            //        //    }

            //        //    Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            //        //    recSet3.DoQuery("select OVPM.DocEntry, VPM2.DocEntry as 'DpDocEntry',    VPM2.AppliedFC from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where ( ORCTDocEntrys) and VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204");

            //        //    while (!recSet3.EoF)
            //        //    {
            //        //        paidAmountDpLc += decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString()) *
            //        //                          rateByDocuments[recSet3.Fields.Item("AppliedFC").Value.ToString()];
            //        //        recSet3.MoveNext();
            //        //    }

            //        //}

            //        try
            //        {
            //            paidAmountDpFc += decimal.Parse(netAmountToDrow);
            //        }
            //        catch (Exception e)
            //        {
            //            SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
            //                BoMessageTime.bmt_Short, true);
            //            SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
            //        }
            //    }

            //}

            CalculateWaightedRate(data, paidAmountDpLc, paidAmountDpFc);



        }

        private static void CalculateWaightedRate(decimal totalInvFc, /*ref bool isRateCalculated,*/ ref string globalRate,
            decimal ratInv, ref decimal paidAmountDpLc, decimal paidAmountDpFc)
        {
            if (totalInvFc == paidAmountDpFc)
            {
                var rate = paidAmountDpLc / totalInvFc;
                globalRate = rate.ToString();
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
                data.GlobalRate = rate.ToString();
                data.IsCalculated = true;
            }
            else if (data.TotalInv > paidAmountDpFc)
            {
                var dif = (data.TotalInv - paidAmountDpFc) * data.RateInv; //invocie Open AmountFC
                paidAmountDpLc += dif;
                var rate = paidAmountDpLc / data.TotalInv;
                data.IsCalculated = true;
                //isRateCalculated = true;
                data.GlobalRate = Math.Round(rate, 6).ToString();
            }
        }

        public static decimal ExchangeRateCorrectionDi(decimal netAmountToDrow, decimal totalInv, decimal ratInv,
            int downPaymentDocEntry, string docCurrency, SAPbobsCOM.Company _comp)
        {
            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net AmountFC To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net AmountFC To Drow 

            var recSetTransferDocEntry =
                (SAPbobsCOM.Recordset)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            var recSerTranferRate =
                (SAPbobsCOM.Recordset)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

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

                var ORCTDocEntry = recSetTransferDocEntry.Fields.Item("DocEntry").Value.ToString();

                recSerTranferRate.DoQuery(
                    "select ORCT.TrsfrSum , RCT2.AppliedFC, RCT2.DocRate from ORCT inner join RCT2 on " +
                    "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry = '" + ORCTDocEntry + "'");

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
                string ORCTDocEntrys = string.Empty;
                Recordset recSet2 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                ORCTDocEntrys =
                    "select ORCT.DocEntry from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" +
                    downPaymentDocEntry +
                    "' and InvType = 203 and ORCT.Canceled = 'N'"; // ეს არის Incoming Paymentebis docentry -ები



                recSet2.DoQuery(
                    "select ORCT.DocEntry, avg(ORCT.TrsfrSum) as 'TrsfrSum' , SUM(RCT2.AppliedFC) as 'AppliedFC' from ORCT inner join RCT2 on " +
                    "ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" + ORCTDocEntrys +
                    ") group by ORCT.DocEntry");
                // აქ მოგვაქვს ინფორმაცია გადახდების მიხედვით სრუტლი თანხა LC - ში დოკუმენტის ნომერი და გადახდილი თანხა უცხოურ ვალუტაში

                List<Tuple<int, decimal, decimal>> sumPayments = new List<Tuple<int, decimal, decimal>>();
                while (!recSet2.EoF)
                {
                    int OCRTDocEntry = int.Parse(recSet2.Fields.Item("DocEntry").Value.ToString());
                    decimal appliedAmountLc = decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString());
                    decimal appliedAmountFc = decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                    sumPayments.Add(new Tuple<int, decimal, decimal>(OCRTDocEntry, appliedAmountLc, appliedAmountFc));
                    recSet2.MoveNext();
                }








                //აქ გვაქვს  ლოკალულ ვალუტაში გატარებილი დოკუმენტების ჯამი და ნომერი რომელიც უნდა გამოაკლდეს გადახდის სრულ თანხას LC- ში

                Recordset recSet4 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet4.DoQuery(
                    "select DocEntry, SUM(LcPrices) as SumLCPayments from ( select  ORCT.DocEntry as 'DocEntry',  SUM(case when AppliedFC = 0 then RCT2.SumApplied else 0 end ) as 'LcPrices' from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry in (" +
                    ORCTDocEntrys +
                    ") group by  RCT2.SumApplied , ORCT.DocEntry ) LcPricesTable group by DocEntry");

                Dictionary<string, decimal> DocumentLcPriceSums = new Dictionary<string, decimal>();
                while (!recSet4.EoF)
                {
                    string OCRTDocEntry = recSet4.Fields.Item("DocEntry").Value.ToString();
                    decimal SumLCPayments = decimal.Parse(recSet4.Fields.Item("SumLCPayments").Value.ToString());
                    DocumentLcPriceSums.Add(OCRTDocEntry, SumLCPayments);
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

                    DocsWithRateAndValue.Add(new XContainer()
                    {
                        CurrRate = rate,
                        OrctDocEntry = paymentDocEntry
                    });
                }
                // აქ გადახდის სრულ თანხას ვაკლებ ლოკალურ ვალუტაში გადახდილი დოკუმენტის ჯამს და ვყოფ სრულ თანხაზე უცხოურ ვალუტაში


                //while (!recSet2.EoF)
                //{
                //    decimal rate = (decimal.Parse(recSet2.Fields.Item("TrsfrSum").Value.ToString()) - DocumentLcPriceSums[recSet2.Fields.Item("DocEntry").Value.ToString()]) /
                //            decimal.Parse(recSet2.Fields.Item("AppliedFC").Value.ToString());
                //    string paymentDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                //    rateByDocuments.Add(paymentDocEntry, rate);
                //    recSet2.MoveNext();

                //}

                Recordset recSet3 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                recSet3.DoQuery(
                    "select ORCT.DocEntry, RCT2.DocEntry as 'DpDocEntry',    RCT2.AppliedFC from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where ORCT.DocEntry  in ( " +
                    ORCTDocEntrys + ") and RCT2.DocEntry = '" + downPaymentDocEntry + "' and InvType = 203");


                Dictionary<string, decimal> dPIncomingPaymentShareAmountFc = new Dictionary<string, decimal>();
                while (!recSet3.EoF)
                {
                    decimal AppliedFcbyDp = decimal.Parse(recSet3.Fields.Item("AppliedFC").Value.ToString());
                    string PaymentDocEntry = recSet3.Fields.Item("DocEntry").Value.ToString();
                    dPIncomingPaymentShareAmountFc.Add(PaymentDocEntry, AppliedFcbyDp);

                    DocsWithRateAndValue.Where(z => z.OrctDocEntry == PaymentDocEntry).ToList()
                        .ForEach(s => s.AmountFC = AppliedFcbyDp);


                    recSet3.MoveNext();
                }

                //var rata = DocsWithRateAndValue.

                decimal LCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC * doc.CurrRate);
                decimal FCSum = DocsWithRateAndValue.Sum(doc => doc.AmountFC);
                sendRate = LCSum / FCSum;



            }



            paidAmountDpLc += sendRate * netAmountToDrow;
            paidAmountDpFc += netAmountToDrow;

            string globalRate = string.Empty;
            CalculateWaightedRate(totalInv, /*ref x,*/ ref globalRate, ratInv, ref paidAmountDpLc, paidAmountDpFc);

            return decimal.Parse(globalRate);

        }

        public static void AddJournalEntryCredit(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235, string vatAccount = "", double vatAmount = 0, string vatGroup = "")
        {

            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

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
            vJE.Lines.Debit = 0;
            vJE.Lines.Credit = amount;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();

            if (vatGroup != "")
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
            else
            {
                string des = _comp.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(des + " Invoice " + reference, BoMessageTime.bmt_Short);
            }
        }

        public static void AddJournalEntryDebit(SAPbobsCOM.Company _comp, string creditCode, string debitCode,
            double amount, int series, string reference, string code, DateTime DocDate, int BPLID = 235, string vatAccount = "", double vatAmount = 0, string vatGroup = "")
        {

            SAPbobsCOM.JournalEntries vJE =
                (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DocDate;
            vJE.DueDate = DocDate;
            vJE.TaxDate = DocDate;
            vJE.Memo = "Income Correction -   Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;
            vJE.Series = series;
            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = BPLID; //branch
            vJE.Lines.Credit = amount;
            vJE.Lines.Debit = 0;
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            vJE.Lines.Add();
            vJE.Lines.BPLID = BPLID;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.ShortName = code;
            vJE.Lines.Credit = 0;
            vJE.Lines.Debit = amount;
            vJE.Lines.FCCredit = 0;
            vJE.Lines.FCDebit = 0;
            // vJE.Series = 17;
            vJE.Lines.Add();
            if (vatGroup != "")
            {
                vJE.Lines.BPLID = BPLID;
                vJE.Lines.AccountCode = creditCode;
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

            vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            else
            {
                string des = _comp.GetLastErrorDescription();
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(des + " Invoice " + reference, BoMessageTime.bmt_Short);
            }
        }

        public static Dictionary<int, string> FormToTransId = new Dictionary<int, string>()
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
            Recordset recSetNew = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetJdt = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            Recordset recSetDeter = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);

            recSetDeter.DoQuery($"select * from OACP where Year(OACP.FinancYear) =  {DateTime.Now.Year}");
            string transitAccount = recSetDeter.Fields.Item("PurcVatOff").Value.ToString();


            recSetNew.DoQuery($"select * from PCH9 where DocEntry = {docNumber}");
            var dpm = recSetNew.Fields.Item("BaseAbs").Value.ToString();
            recSetNew.DoQuery($"SELECT transid FROM ODPO WHERE DocEntry = {dpm}");
            var transid = recSetNew.Fields.Item("transid").Value;
            decimal drFc = 0;
            decimal dr = 0;
            decimal cr = 0;
            decimal crFc = 0;
            if (transid != null && transid.ToString() != "0")
            {

                recSetJdt.DoQuery($"SELECT * FROM JDT1 WHERE TransId = {transid}");
                while (!recSetJdt.EoF)
                {

                    if (recSetJdt.Fields.Item("Account").Value.ToString() == transitAccount)
                    {
                        recSetJdt.MoveNext();
                    }
                    else
                    {
                        dr = decimal.Parse(recSetJdt.Fields.Item("Debit").Value.ToString());
                        drFc = decimal.Parse(recSetJdt.Fields.Item("FCDebit").Value.ToString());
                        cr = decimal.Parse(recSetJdt.Fields.Item("Credit").Value.ToString());
                        crFc = decimal.Parse(recSetJdt.Fields.Item("FCCredit").Value.ToString());
                        vatAccountDownPayment = recSetJdt.Fields.Item("Account").Value.ToString();
                        vatGroup = recSetJdt.Fields.Item("VatGroup").Value.ToString();
                        break;
                    }
                }
            }

            decimal vatAmountDownPayment = dr + cr;
            decimal vatAmountDownPaymentFc = drFc + crFc;

            recSet.DoQuery("select DebPayAcct from OCRD where CardCode = '" + businesPartnerCardCode + "'");
            string BpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();
            if (!string.IsNullOrWhiteSpace(applied))
            {
                var objRS = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
                var objRS234 = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                if (objRS.Fields.Item("TransType").Value.ToString() != "13")
                {
                    objRS234.DoQuery("select  BPLId from OPCH where BPLName = N'" + bplName + "'");
                }
                else if (objRS.Fields.Item("TransType").Value.ToString() != "18")
                {
                    objRS234.DoQuery("select  BPLId from OINV where BPLName = N'" + bplName + "'");
                }

                int bplID = Convert.ToInt32(objRS234.Fields.Item("BPLId").Value);


                while (!objRS.EoF && transid != null && transid.ToString() != "0")
                {
                    bool difCalculated = false;
                    bool exchangeRateCalculated = false;
                    bool transitCalculated = false;
                    string invAcc = objRS.Fields.Item("Account").Value.ToString();
                    string invvatGroup = objRS.Fields.Item("VatGroup").Value.ToString();
                    decimal invCreditFc = decimal.Parse(objRS.Fields.Item("FCCredit").Value.ToString());
                    decimal invCredit = decimal.Parse(objRS.Fields.Item("Credit").Value.ToString());
                    decimal invDebit = decimal.Parse(objRS.Fields.Item("Debit").Value.ToString());


                    if (invAcc == vatAccountDownPayment && invvatGroup == vatGroup && Math.Abs(invCreditFc) == vatAmountDownPaymentFc)
                    {
                        vatAmountInvDpm = invCredit + invDebit;
                        difCalculated = true;
                    }
                    if (invAcc == ExchangeGain || invAcc == ExchangeLoss)
                    {
                        exchangeRateAmount = invCredit + invDebit;
                        exchangeRateCalculated = true;
                    }

                    if (invAcc == transitAccount)
                    {
                        invTransitAmount = invDebit + invCredit;
                        transitCalculated = true;
                    }
                    if (exchangeRateCalculated && difCalculated && transitCalculated)
                    {
                        break;
                    }
                    objRS.MoveNext();
                }

                vatDifferenceBetweenDpmInv = invTransitAmount - vatAmountInvDpm;
                correctionAmount = exchangeRateAmount - vatDifferenceBetweenDpmInv;



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
                        if (transid != null && transid.ToString() != "0")
                        {
                            if (account == ExchangeGain)
                            {
                                AddJournalEntryCredit(_comp, BpControlAcc, ExchangeGain,
                                    Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                                    bplID, vatAccountDownPayment, Convert.ToDouble(correctionAmount), vatGroup);

                            }
                            else if (account == ExchangeLoss)
                            {
                                AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                                    Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID, vatAccountDownPayment, Convert.ToDouble(correctionAmount), vatGroup);
                            }
                        }
                        else
                        {
                            if (account == ExchangeGain)
                            {
                                AddJournalEntryCredit(_comp, BpControlAcc, ExchangeGain,
                                    Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                                    bplID);

                            }
                            else if (account == ExchangeLoss)
                            {
                                AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                                    Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID);
                            }
                        }

                    }
                    else if (FormType.ToString() == "141")
                    {
                        if (transid != null && transid.ToString() != "0")
                        {
                            if (account == ExchangeGain)
                            {
                                AddJournalEntryCredit(_comp, BpControlAcc, ExchangeGain,
                                    Convert.ToDouble(correctionAmount), series, docNumber, businesPartnerCardCode, docDate,
                                    bplID, vatAccountDownPayment, Convert.ToDouble(correctionAmount - exchangeRateAmount), vatGroup);

                            }
                            else if (account == ExchangeLoss)
                            {
                                AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                                   Convert.ToDouble(vatDifferenceBetweenDpmInv + exchangeRateAmount), series, docNumber, businesPartnerCardCode, docDate, bplID, vatAccountDownPayment, Convert.ToDouble(vatDifferenceBetweenDpmInv), vatGroup);
                            }
                        }
                        else
                        {
                            if (account == ExchangeGain)
                            {
                                AddJournalEntryCredit(_comp, BpControlAcc, ExchangeGain,
                                    Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                                    bplID);
                            }
                            else if (account == ExchangeLoss)
                            {
                                AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                                    Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID);
                            }
                        }

                    }

                    objRS.MoveNext();
                }
            }

        }

        public static void CorrectionJournalEntryDI(SAPbobsCOM.Company _comp, int FormType, string businesPartnerCardCode, string docNumber, string bplName, string ExchangeGain, string ExchangeLoss, DateTime docDate)
        {
            Recordset recSet = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet.DoQuery("select DebPayAcct from OCRD where CardCode = '" + businesPartnerCardCode + "'");
            string BpControlAcc = recSet.Fields.Item("DebPayAcct").Value.ToString();

            var objRS = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
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
            var objRS234 = (SAPbobsCOM.Recordset)(_comp).GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            if (objRS.Fields.Item("TransType").Value.ToString() != "13")
            {
                objRS234.DoQuery("select  BPLId from OPCH where BPLName = '" + bplName + "'");
            }
            else if (objRS.Fields.Item("TransType").Value.ToString() != "18")
            {
                objRS234.DoQuery("select  BPLId from OINV where BPLName = '" + bplName + "'");
            }

            int bplID = Convert.ToInt32(objRS234.Fields.Item("BPLId").Value);

            Recordset recSet12 = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
            recSet12.DoQuery("select Series from NNM1 where ObjectCode = 30 and Locked = 'N' and BPLId is  null");
            int series = int.Parse(recSet12.Fields.Item("Series").Value.ToString());

            if (bplID == 0)
            {
                bplID = 235;
            }

            while (!objRS.EoF)
            {
                var account = objRS.Fields.Item("Account").Value.ToString();

                if (FormType.ToString() == "133")
                {
                    if (account == ExchangeGain)
                    {
                        AddJournalEntryCredit(_comp, BpControlAcc, ExchangeGain,
                            Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                            bplID);

                    }
                    else if (account == ExchangeLoss)
                    {
                        AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                            Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID);
                    }
                }
                else if (FormType.ToString() == "141")
                {
                    if (account == ExchangeGain)
                    {
                        AddJournalEntryCredit(_comp, BpControlAcc, ExchangeGain,
                            Convert.ToDouble(objRS.Fields.Item("Credit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate,
                            bplID);
                    }
                    else if (account == ExchangeLoss)
                    {
                        AddJournalEntryDebit(_comp, ExchangeLoss, BpControlAcc,
                            Convert.ToDouble(objRS.Fields.Item("Debit").Value.ToString()), series, docNumber, businesPartnerCardCode, docDate, bplID);
                    }
                }

                objRS.MoveNext();
            }


        }


    }
}
