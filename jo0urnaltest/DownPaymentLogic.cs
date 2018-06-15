using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;
using SAPbouiCOM;

namespace jo0urnaltest
{
    public  static class DownPaymentLogic
    {
        /// <summary>
        ///   
        /// </summary>
        /// <param name="downPaymentToDrow"> down payment - is forma gaxsnili invoisidan (A/R ; A/P) </param>
        /// <param name="_comp"> SAPbobsCOM company </param>
        /// <param name="formType"> invoisis(mshobeli) formis tipi (A/R ; A/P) </param>
        /// <param name="docCurrency">invoisis(mshobeli) formis valuta </param>
        /// <param name="totalInv"> Total Befor Discounts damatebuli Tax-i invoisis pormidan </param>
        /// <param name="isRateCalculated"> tu isRateCalculated true daabruna eseigi invoisis BP Currency velshi  unda cahvsvat  globalRate </param>
        /// <param name="globalRate"> tu isRateCalculated true daabruna eseigi invoisis BP Currency velshi  unda cahvsvat  globalRate </param>
        /// <param name="ratInv">invoisis(mshobeli) formis valuta </param>
        public static void ExchangeRateCorrectionUi(Form downPaymentToDrow, SAPbobsCOM.Company _comp, string formType, string docCurrency, decimal totalInv, out bool isRateCalculated, out string globalRate, decimal ratInv)
        {
            isRateCalculated = false;
            Form downPaymentToDrowForm = downPaymentToDrow;
            Item downPaymentFormMatrix = downPaymentToDrowForm.Items.Item("6");//Down Payment to drow
            Matrix matrix = (SAPbouiCOM.Matrix)downPaymentFormMatrix.Specific;
            globalRate = "1,0000";
            decimal paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net Amount To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net Amount To Drow 

            for (int i = 1; i <= matrix.RowCount; i++)
            {
                var checkbox = (SAPbouiCOM.CheckBox)matrix.Columns.Item("380000138").Cells.Item(i).Specific;
                if (checkbox.Checked)
                {
                    EditText txtMoney = (SAPbouiCOM.EditText)matrix.Columns.Item("29").Cells.Item(i).Specific;//net amount to drow//TODO
                    EditText txtID = (SAPbouiCOM.EditText)matrix.Columns.Item("68").Cells.Item(i).Specific;//docNumber
                    string netAmountToDrow = txtMoney.Value.Split(' ')[0]; //net amount to drow

                    var objMD = (SAPbobsCOM.Recordset)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    if (formType == "133")
                    {
                        Recordset recSet2 =
                            (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                        recSet2.DoQuery("SELECT DocEntry FROM ODPI WHERE DocNum = '" + txtID.Value + "'");
                        var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                        objMD.DoQuery("select ORCT.TrsfrDate, ORCT.DocDate from ORCT inner join RCT2 on ORCT.DocEntry = RCT2.DocNum where RCT2.DocEntry = '" + dpDocEntry + "' and InvType = 203");
                    }
                    else
                    {
                        Recordset recSet2 =
                            (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                        recSet2.DoQuery("SELECT DocEntry FROM ODPO WHERE DocNum = '" + txtID.Value + "'");
                        var dpDocEntry = recSet2.Fields.Item("DocEntry").Value.ToString();
                        objMD.DoQuery("select OVPM.TrsfrDate, OVPM.DocDate from OVPM inner join VPM2 on OVPM.DocEntry = VPM2.DocNum where VPM2.DocEntry = '" + dpDocEntry + "' and InvType = 204");
                    }
                    string transferDate = Convert.ToDateTime(objMD.Fields.Item("TrsfrDate").Value).ToString("s");

                    string date = transferDate != "1899-12-30T00:00:00" ? Convert.ToDateTime(objMD.Fields.Item("TrsfrDate").Value.ToString()).ToString("s") : Convert.ToDateTime(objMD.Fields.Item("DocDate").Value.ToString()).ToString("s");

                    Recordset recSet =
                        (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);
                    recSet.DoQuery("select Rate from ORTT where RateDate = '" + date + "' and Currency = '" + docCurrency + "'");
                    string rateDp = recSet.Fields.Item("Rate").Value.ToString();

                    try
                    {
                        paidAmountDpLc += (decimal.Parse(rateDp) * decimal.Parse(netAmountToDrow));
                        paidAmountDpFc += decimal.Parse(netAmountToDrow);
                    }
                    catch (Exception e)
                    {
                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short, true);
                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                    }
                }

            }
            if (totalInv == paidAmountDpFc)
            {
                var rate = paidAmountDpLc / totalInv;
                globalRate = rate.ToString();
                isRateCalculated = true;
            }
            else if (totalInv > paidAmountDpFc)
            {
                var dif = (totalInv - paidAmountDpFc) * ratInv;//invocie Open Amount
                paidAmountDpLc += dif;
                var rate = paidAmountDpLc / totalInv;
                isRateCalculated = true;
                globalRate = Math.Round(rate, 4).ToString();
            }

        }


        public static decimal ExchangeRateCorrectionDi(decimal netAmountToDrow, decimal totalInv, decimal ratInv,
            decimal rateDp)
        {
            decimal
                paidAmountDpLc = 0m; // A/R DownPayment - ში არჩეული თანხა ლოკალურ ვალუტაში //Net Amount To Drow * Rate
            decimal paidAmountDpFc = 0m; //  A/R DownPayment - ში არჩეული თანხა FC //Net Amount To Drow 

            paidAmountDpLc += rateDp * netAmountToDrow;
            paidAmountDpFc += netAmountToDrow;
            decimal globalRate;

            if (totalInv == paidAmountDpFc)
            {
                decimal rate = paidAmountDpLc / totalInv;
                  globalRate = rate;
                return globalRate;

            }
            else if (totalInv > paidAmountDpFc)
            {
                decimal dif = (totalInv - paidAmountDpFc) * ratInv; //invocie Open Amount
                paidAmountDpLc += dif;
                decimal rate = paidAmountDpLc / totalInv;
                  globalRate = Math.Round(rate, 4);
                return globalRate;
            }
            return globalRate = 0;
        }

    }
}

