using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Security.Cryptography.X509Certificates;
using Appocalypto;
using SAPApi;
using SAPbobsCOM;
using SAPbouiCOM;
using Application = SAPbouiCOM.Framework.Application;
using Company = SAPbobsCOM.Company;

namespace jo0urnaltest
{
    class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            try
            {
                Application oApp = null;
                oApp = args.Length < 1 ? new Application() : new Application(args[0]);
                SAPApi.DIManager _diManager = new DIManager();
                _comp = (Company)Application.SBO_Application.Company.GetDICompany();
                _diManager.AddField("OINV", "OldRate", "სისტემური კურსი", BoFieldTypes.db_Alpha, 10, false, true);
                _diManager.AddField("OPCH", "OldRate", "სისტემური კურსი", BoFieldTypes.db_Alpha, 10, false, true);
                Mob appo = new Mob();
                appo.Run(5);
                Application.SBO_Application.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
                //Application.SBO_Application.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
                Recordset recSet = (Recordset)_comp.GetBusinessObject(BoObjectTypes.BoRecordset);

                string query = "SELECT LinkAct_25, LinkAct_21 FROM OACP where PeriodCat ='" +
                               DateTime.Now.Year + "'";
                recSet.DoQuery(query);
                ExchangeGain = recSet.Fields.Item("LinkAct_25").Value.ToString();
                ExchangeLoss = recSet.Fields.Item("LinkAct_21").Value.ToString();
                oApp.Run();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

        // private static readonly SAPbobsCOM.Company _comp;
        public static Company _comp { get; set; }
        public static string ExchangeGain { get; set; }
        public static string ExchangeLoss { get; set; }
        public static string BpControlAcc { get; set; }
        public static string PeriodCat { get; set; }






        private static void AddJournalEntry(string creditCode, string debitCode, double amount, string reference)
        {

            SAPbobsCOM.JournalEntries vJE = (SAPbobsCOM.JournalEntries)_comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
            vJE.ReferenceDate = DateTime.Now;
            vJE.DueDate = DateTime.Now;
            vJE.TaxDate = DateTime.Now;
            vJE.Memo = "Income Correction - A/R Invoice " + reference;
            //vJE.TransactionCode = "13";
            vJE.Reference = reference;


            vJE.TransactionCode = "1";

            vJE.Lines.BPLID = 235;
            vJE.Lines.Debit = amount;
            vJE.Lines.Credit = 0;
            vJE.Lines.AccountCode = debitCode;
            vJE.Lines.Add();
            vJE.Lines.AccountCode = creditCode;
            vJE.Lines.ShortName = CardCode;
            vJE.Lines.Debit = 0;
            vJE.Lines.Credit = amount;
            //vJE.Lines.Add();
            int i = vJE.Add();
            if (i == 0)
            {
                return;
            }
            else
            {
                string des = _comp.GetLastErrorDescription();
            }
        }
        private static string docNum;
        private static string CardCode;
        private static string downPaymentAmount;

        static decimal totalInv = 0;
        static decimal ratInv = 0;
        static bool isRateCalculated = false;
        static string globalRate = string.Empty;
        private static string formType;

        public static Dictionary<int, string> FormToTransId = new Dictionary<int, string>()
        {
            {133, "13"},
            {141, "18"}
        };

        public static string bplName { get; set; }
        public static string BusinesPartnerName { get; set; }
        public static string DocCurrency { get; internal set; }

        //private static void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        //{
        //    BubbleEvent = true;

        //    SAPbouiCOM.BoEventTypes EventEnum = 0;
        //    EventEnum = pVal.EventType;
        //    if (pVal.ItemUID == "1" && (pVal.FormTypeEx == "133" || pVal.FormTypeEx == "141") &&
        //        pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK)
        //    {
        //        try
        //        {
        //            bplName = ((SAPbouiCOM.ComboBox)(SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
        //                .Item("2001").Specific)).Selected.Description;
        //        }
        //        catch (Exception e)
        //        {
        //            //branch araa
        //        }


        //        BusinesPartnerName = ((SAPbouiCOM.EditText)(SAPbouiCOM.Framework.Application.SBO_Application
        //           .Forms.ActiveForm.Items
        //           .Item("4").Specific)).Value;
        //    }
        //    //string bplName = ((SAPbouiCOM.ComboBox)(SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
        //    //    .Item("8").Specific)).Value.ToString().Trim(' ');
        //    if (pVal.ItemUID == "1" && (pVal.FormTypeEx == "133" || pVal.FormTypeEx == "141") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        //    {
        //        try
        //        {
        //            if (pVal.BeforeAction == false && pVal.ActionSuccess == true)
        //            {
        //                DownPaymentLogic.DownPaymentLogic.CorrectionJournalEntryUI(_comp, pVal.FormType, CardCode, downPaymentAmount, docNum, bplName, ExchangeGain, ExchangeLoss, DateTime.Now);
        //            }
        //            else
        //            {
        //                var invoiceForm = Application.SBO_Application.Forms.ActiveForm;
        //                var docNumInvItm = invoiceForm.Items.Item("8");
        //                var docNumInvEditText = (SAPbouiCOM.EditText)docNumInvItm.Specific;
        //                docNum = docNumInvEditText.Value;

        //                var cardCodeItem = invoiceForm.Items.Item("4");
        //                var cardCodeEditText = (SAPbouiCOM.EditText)cardCodeItem.Specific;
        //                CardCode = cardCodeEditText.Value;

        //                var totalDownPaymentItem = invoiceForm.Items.Item("204");
        //                var totalDownPaymentEditText = (SAPbouiCOM.EditText)totalDownPaymentItem.Specific;
        //                downPaymentAmount = totalDownPaymentEditText.Value;
        //            }

        //        }
        //        catch (Exception e)
        //        {
        //            SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(_comp.GetLastErrorDescription());
        //        }
        //    }


        //    if ((pVal.FormTypeEx == "133" || pVal.FormTypeEx == "141") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE)
        //    {
        //        try
        //        {
        //            // Console.WriteLine(pVal.EventType);
        //            SAPbouiCOM.Form invoiceForm = Application.SBO_Application.Forms.ActiveForm;
        //            if (isRateCalculated)
        //            {
        //                isRateCalculated = false;

        //                try
        //                {

        //                    var txtRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("64").Specific);// BP Currency A/R Invoice  Exchange Rate
        //                    try
        //                    {
        //                        var ee = (Math.Round(decimal.Parse(globalRate), 4)).ToString().Replace(".", ",") + ((globalRate.Contains(".") || globalRate.Contains(",")) ? "" : "");
        //                        txtRate.Value = ee.Replace(",", ".");
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short, true);
        //                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
        //                    }
        //                }
        //                catch (Exception)
        //                {
        //                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("შეიყვანეთ საქონლის ფასი", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
        //                }
        //            }
        //            else
        //            {
        //                if (pVal.FormTypeEx == "133" || pVal.FormTypeEx == "141")
        //                {
        //                    try
        //                    {
        //                        var txtOldRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("Item_0000").Specific);
        //                        var txtRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("64").Specific);
        //                        if (!string.IsNullOrWhiteSpace(txtRate.Value))
        //                        {
        //                            try
        //                            {
        //                                string val = Math.Round(decimal.Parse(txtRate.Value), 4).ToString()
        //                                    .Replace(".", ",");
        //                                if (string.IsNullOrWhiteSpace(txtOldRate.Value) || txtOldRate.Value == "1,0000")
        //                                {
        //                                    txtOldRate.Value = val;
        //                                }
        //                            }
        //                            catch (Exception e)
        //                            {
        //                                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
        //                                    BoMessageTime.bmt_Short, true);
        //                                SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
        //                            }
        //                        }
        //                    }
        //                    catch (Exception e)
        //                    {
        //                        SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
        //                            BoMessageTime.bmt_Short, true);
        //                        SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();

        //                    }
        //                }
        //            }

        //        }
        //        catch (Exception e)
        //        {
        //            SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(_comp.GetLastErrorDescription());

        //        }
        //    }



        //    if (pVal.ItemUID == "213" && (pVal.FormTypeEx == "133" || pVal.FormTypeEx == "141") && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        //    {
        //        try
        //        {

        //            formType = pVal.FormTypeEx.ToString();
        //            if (pVal.BeforeAction == true)
        //            {
        //                var invoiceForm = Application.SBO_Application.Forms.ActiveForm;
        //                var txtTotal = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("22").Specific);//totalInv before discount
        //                var vat = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("27").Specific);
        //                var Discount = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("42").Specific);
        //                if (string.IsNullOrWhiteSpace(txtTotal.Value))
        //                {
        //                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage("მიუთითეთ თანხა", SAPbouiCOM.BoMessageTime.bmt_Short, true);
        //                    BubbleEvent = false;
        //                    return;
        //                }
        //                string x = txtTotal.Value.Split(' ')[0];
        //                string x1 = vat.Value.Split(' ')[0];
        //                string x2 = Discount.Value.Split(' ')[0] == string.Empty ? "0" : Discount.Value.Split(' ')[0];
        //                try
        //                {

        //                    totalInv = decimal.Parse(x) - decimal.Parse(x2) + decimal.Parse(String.IsNullOrEmpty(x1) ? "0" : x1);

        //                    var txtRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("Item_0000").Specific);
        //                    ratInv = decimal.Parse(txtRate.Value.Replace(",", "."));
        //                }
        //                catch (Exception e)
        //                {
        //                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short, true);
        //                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
        //                }
        //            }
        //        }
        //        catch (Exception e)
        //        {
        //            SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(_comp.GetLastErrorDescription());
        //        }

        //        if (pVal.BeforeAction == true)
        //        {
        //            try
        //            {

                   
        //            var invoiceForm = Application.SBO_Application.Forms.ActiveForm;
        //            var docDate = DateTime.ParseExact(
        //                ((EditText)(invoiceForm.Items.Item("10").Specific)).Value.ToString(),
        //                "yyyyMMdd", CultureInfo.InvariantCulture);
        //            ratInv = decimal.Parse(UiManager.GetCurrencyRate(DocCurrency, docDate, _comp).ToString());
        //            }
        //            catch (Exception e)
        //            {
                         
                      
        //            }
        //        }

        //    }

        //     if (pVal.ItemUID == "1" && pVal.FormTypeEx == "60511" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
        //    {

        //        try
        //        {
        //            DownPaymentLogic.DownPaymentLogic.ExchangeRateCorrectionUi(SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm, _comp, formType, DocCurrency, totalInv,/* out isRateCalculated,*/ out globalRate, ratInv);


        //        }
        //        catch (Exception e)
        //        {
        //            Application.SBO_Application.SetStatusBarMessage(e.Message,
        //                BoMessageTime.bmt_Short, true);
        //        }
        //    }
        //}



        static void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //Exit Add-On
                    System.Windows.Forms.Application.Exit();
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_FontChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_ServerTerminition:
                    break;
                default:
                    break;
            }
        }
    }
}
