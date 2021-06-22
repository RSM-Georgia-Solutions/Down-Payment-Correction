using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using DownPaymentLogic.Classes;
using SAPApi;
using SAPbobsCOM;
using SAPbouiCOM;
using SAPbouiCOM.Framework;
using Application = SAPbouiCOM.Framework.Application;
using Company = SAPbouiCOM.Company;

namespace jo0urnaltest
{
    [FormAttribute("133", "A_R Invoice.b1f")]
    class A_R_Invoice : SystemFormBase
    {
        public A_R_Invoice()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("Item_0000").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("Item_1000").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("4").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("213").Specific));
            this.Button1.PressedAfter += new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button1_PressedAfter);
            this.Button1.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button1_ClickBefore);
            try
            {
                this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
                this.Button0.PressedAfter +=
                    new SAPbouiCOM._IButtonEvents_PressedAfterEventHandler(this.Button0_PressedAfter);
                this.Button0.PressedBefore +=
                    new SAPbouiCOM._IButtonEvents_PressedBeforeEventHandler(this.Button0_PressedBefore);
            }
            catch
            {
                this.ButtonCombo0 = ((SAPbouiCOM.ButtonCombo)(this.GetItem("1").Specific));
                this.ButtonCombo0.PressedBefore += new SAPbouiCOM._IButtonComboEvents_PressedBeforeEventHandler(this.ButtonCombo0_PressedBefore);
                this.ButtonCombo0.PressedAfter += new SAPbouiCOM._IButtonComboEvents_PressedAfterEventHandler(this.ButtonCombo0_PressedAfter);
            }
           
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
            this.CloseAfter += new SAPbouiCOM.Framework.FormBase.CloseAfterHandler(this.Form_CloseAfter);
            this.ActivateAfter += new ActivateAfterHandler(this.Form_ActivateAfter);

        }

        private SAPbouiCOM.Button Button0;
        private static string id;
        private static string down;


        private void OnCustomInitialize()
        {
            Program.BusinesPartnerName = string.Empty;
            Program.bplName = string.Empty;
            DataForCalculationRate = new DataForCalculationRate();
            DownPaymentsForInvFormIds = new List<Dictionary<string, string>>();
        }



        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.EditText EditText1;
        private Button Button1;



        public static List<Dictionary<string, string>> DownPaymentsForInvFormIds { get; set; }


        public static Dictionary<string, string> DownPaymentsForInvFormId = new Dictionary<string, string>();

        private string _formUIdInv;
        private string _formUIdDps;
        //   private string globalRate = string.Empty;

        private void Button0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {

            BubbleEvent = true;

            Form arInoviceForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (pVal.FormMode != 3)
            {
                stop = true;
                return;
            }
            stop = false;

            try
            {


                DataForCalculationRate.BusinesPartnerName = ((SAPbouiCOM.EditText)(arInoviceForm.Items
                    .Item("4").Specific)).Value;
                try
                {
                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                           .Item("2001").Specific)).Selected.Description;
                }
                catch (Exception)
                {

                    // db without branch
                }

                var docNumInvItm = arInoviceForm.Items.Item("8");
                var docNumInvEditText = (SAPbouiCOM.EditText)docNumInvItm.Specific;
                DataForCalculationRate.DocNum = docNumInvEditText.Value;

                var cardCodeItem = arInoviceForm.Items.Item("4");
                var cardCodeEditText = (SAPbouiCOM.EditText)cardCodeItem.Specific;
                DataForCalculationRate.CardCode = cardCodeEditText.Value;

                var totalDownPaymentItem = arInoviceForm.Items.Item("204");
                var totalDownPaymentEditText = (SAPbouiCOM.EditText)totalDownPaymentItem.Specific;
                DataForCalculationRate.DownPaymentAmount = totalDownPaymentEditText.Value;
                DataForCalculationRate.FormTypex = "133";

                var txtTotalWithCurr =
                    (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("22").Specific); //totalInv before discount
                var vatWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("27").Specific);
                var discountWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("42").Specific);

                string txtTotal = txtTotalWithCurr.Value.Split(' ')[0];
                string vat = vatWithCurr.Value.Split(' ')[0];
                string discount = discountWithCurr.Value.Split(' ')[0] == string.Empty
                    ? "0"
                    : discountWithCurr.Value.Split(' ')[0];


                SBObob sbObob = (SBObob)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                string currency = ((ComboBox)SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
                    .Item("63").Specific).Value;
                DataForCalculationRate.DocCurrency = currency;
                if (currency == "GEL")
                {
                    stop = true;
                    return;
                }
                string postingDateString = ((EditText)arInoviceForm.Items.Item("10").Specific).Value;
                DateTime postingDate = DateTime.ParseExact(postingDateString, "yyyyMMdd", null);
                decimal currencyValue =
                    Math.Round(
                        decimal.Parse(sbObob.GetCurrencyRate(currency, postingDate).Fields.Item(0).Value
                            .ToString()), 6);
                ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                    currencyValue.ToString();
                DateTime docDate = DateTime.ParseExact(
                    ((EditText)(arInoviceForm.Items.Item("10").Specific)).Value,
                    "yyyyMMdd", CultureInfo.InvariantCulture);
                var txtRate = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific);
                DataForCalculationRate.RateInv = decimal.Parse(txtRate.Value.Replace(",", "."));
                DataForCalculationRate.RateInv = decimal.Parse(UiManager.GetCurrencyRate(DataForCalculationRate.DocCurrency, docDate, Program._comp).ToString());
                var item_serviceCombobox = (SAPbouiCOM.ComboBox)UIAPIRawForm.Items.Item("3").Specific;

                var selected = item_serviceCombobox.Selected.Value;
                Matrix invMAtrix;
                string totalTaxAmountColumn = "";
                if (selected == "S")
                {
                    invMAtrix = (Matrix)arInoviceForm.Items.Item("39").Specific;
                    totalTaxAmountColumn = "33";
                }
                else
                {
                    invMAtrix = (Matrix)arInoviceForm.Items.Item("38").Specific;
                    totalTaxAmountColumn = "83";
                }
                try
                {
                    decimal taxTotalAmount = 0;
                    var vatFromTotalBefore = decimal.Parse(txtTotal);
                    for (int i = 1; i < invMAtrix.RowCount; i++)
                    {
                        EditText txtMoney = (SAPbouiCOM.EditText)invMAtrix.Columns.Item(totalTaxAmountColumn).Cells.Item(i).Specific; //Tax Amount Line
                        decimal taxLineAmount = decimal.Parse(txtMoney.Value.Split(' ')[0]);
                        taxTotalAmount += taxLineAmount;
                    }

                    DataForCalculationRate.TotalInv = vatFromTotalBefore + taxTotalAmount;

                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short,
                    true);
                //branch araa
            }
            try
            {
                var glAccount = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("2010").Specific);
                DataForCalculationRate.GlAccount = glAccount.Value;
                DownPaymentLogic.DownPaymentLogic.ExchangeRateCorrectionUi(DataForCalculationRate, Program._comp);

            }
            catch (Exception e)
            {
                Application.SBO_Application.SetStatusBarMessage(e.Message,
                    BoMessageTime.bmt_Short, true);
            }




        }


        private void Button0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (string.IsNullOrWhiteSpace(DataForCalculationRate.DownPaymentAmount))
            {
                stop = true;
                return;
            }
            if (pVal.ActionSuccess && !stop)
            {
                DownPaymentLogic.DownPaymentLogic.CorrectionJournalEntryUI(Program._comp, 133, DataForCalculationRate.CardCode,
                    DataForCalculationRate.DownPaymentAmount, DataForCalculationRate.DocNum, DataForCalculationRate.BplName, DataForCalculationRate.GlAccount, Program.ExchangeGain, Program.ExchangeLoss, DataForCalculationRate.PostingDate);
            }

        }

        private void Button1_ClickBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            _formUIdInv = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.UDFFormUID;
            DataForCalculationRate.FormUIdInv = _formUIdInv;
            BubbleEvent = true;

            Form arInoviceForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;

            try
            {


                DataForCalculationRate.BusinesPartnerName = ((SAPbouiCOM.EditText)(arInoviceForm.Items
                    .Item("4").Specific)).Value;
                try
                {
                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                           .Item("2001").Specific)).Selected.Description;
                }
                catch (Exception)
                {

                    // db without branch
                }

                var docNumInvItm = arInoviceForm.Items.Item("8");
                var docNumInvEditText = (SAPbouiCOM.EditText)docNumInvItm.Specific;
                DataForCalculationRate.DocNum = docNumInvEditText.Value;

                var cardCodeItem = arInoviceForm.Items.Item("4");
                var cardCodeEditText = (SAPbouiCOM.EditText)cardCodeItem.Specific;
                DataForCalculationRate.CardCode = cardCodeEditText.Value;

                var totalDownPaymentItem = arInoviceForm.Items.Item("204");
                var totalDownPaymentEditText = (SAPbouiCOM.EditText)totalDownPaymentItem.Specific;
                DataForCalculationRate.DownPaymentAmount = totalDownPaymentEditText.Value;
                DataForCalculationRate.FormTypex = "133";
                var txtTotalWithCurr =
                    (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("22").Specific); //totalInv before discount
                var vatWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("27").Specific);
                var discountWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("42").Specific);

                string txtTotal = txtTotalWithCurr.Value.Split(' ')[0] == string.Empty ? "0" : txtTotalWithCurr.Value.Split(' ')[0];
                string vat = vatWithCurr.Value.Split(' ')[0] == string.Empty ? "0" : vatWithCurr.Value.Split(' ')[0];
                string discount = discountWithCurr.Value.Split(' ')[0] == string.Empty
                    ? "0"
                    : discountWithCurr.Value.Split(' ')[0];


                SBObob sbObob = (SBObob)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                string currency = ((ComboBox)SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
                    .Item("63").Specific).Value;
                DataForCalculationRate.DocCurrency = currency;
                string postingDateString = ((EditText)arInoviceForm.Items.Item("10").Specific).Value;
                if (string.IsNullOrWhiteSpace(postingDateString))
                {
                    Application.SBO_Application.SetStatusBarMessage("მიუთითეთ თარიღი",
                        BoMessageTime.bmt_Short, true);
                    return;
                }
                DateTime postingDate = DateTime.ParseExact(postingDateString, "yyyyMMdd", null);
                DataForCalculationRate.PostingDate = postingDate;
                decimal currencyValue = Math.Round(decimal.Parse(sbObob.GetCurrencyRate(currency, postingDate).Fields.Item(0).Value.ToString()), 6);
                ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                    currencyValue.ToString();
                DateTime docDate = DateTime.ParseExact(
                    ((EditText)(arInoviceForm.Items.Item("10").Specific)).Value,
                    "yyyyMMdd", CultureInfo.InvariantCulture);
                var txtRate = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific);
                DataForCalculationRate.RateInv = decimal.Parse(txtRate.Value.Replace(",", "."));
                DataForCalculationRate.RateInv = decimal.Parse(UiManager.GetCurrencyRate(DataForCalculationRate.DocCurrency, docDate, Program._comp).ToString());
                var item_serviceCombobox = (SAPbouiCOM.ComboBox)UIAPIRawForm.Items.Item("3").Specific;

                var selected = item_serviceCombobox.Selected.Value;
                Matrix invMAtrix;
                string totalTaxAmountColumn = "";
                if (selected == "S")
                {
                    invMAtrix = (Matrix)arInoviceForm.Items.Item("39").Specific;
                    totalTaxAmountColumn = "33";
                }
                else
                {
                    invMAtrix = (Matrix)arInoviceForm.Items.Item("38").Specific;
                    totalTaxAmountColumn = "83";
                }

                try
                {
                    decimal taxTotalAmount = 0;
                    decimal vatFromTotalBefore = decimal.Parse(txtTotal);
                    for (int i = 1; i < invMAtrix.RowCount; i++)
                    {
                        EditText txtMoney = (SAPbouiCOM.EditText)invMAtrix.Columns.Item(totalTaxAmountColumn).Cells.Item(i).Specific; //Tax Amount Line
                        decimal taxLineAmount = decimal.Parse(txtMoney.Value.Split(' ')[0]);
                        taxTotalAmount += taxLineAmount;
                    }

                    DataForCalculationRate.TotalInv = vatFromTotalBefore + taxTotalAmount;

                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short,
                    true);
                //branch araa
            }




        }



        private void Form_CloseAfter(SBOItemEventArg pVal)
        {
            Program.BusinesPartnerName = string.Empty;
            Program.bplName = string.Empty;

        }

        private bool stop = false;
        public DataForCalculationRate DataForCalculationRate { get; set; }

       




    

        private void Button1_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {

            Form downPaymentToDrowForm = Application.SBO_Application.Forms.ActiveForm;

            if (downPaymentToDrowForm.TypeEx != "60511")
            {
                return;
            }
            _formUIdDps = downPaymentToDrowForm.UDFFormUID;
            DataForCalculationRate.FormUIdDps = _formUIdDps;

            if (SharedClass.ListOfDataForCalculationRates.Count(x => x.FormUIdInv == DataForCalculationRate.FormUIdInv &&
                                                                     x.FormUIdDps == DataForCalculationRate.FormUIdDps) == 0)
            {
                SharedClass.ListOfDataForCalculationRates.Add(DataForCalculationRate);
            }
            else
            {
                var x1 = SharedClass.ListOfDataForCalculationRates.First(
                    x => x.FormUIdInv == DataForCalculationRate.FormUIdInv &&
                         x.FormUIdDps == DataForCalculationRate.FormUIdDps);
                x1 = DataForCalculationRate;

            }



        }


        private void Form_ActivateAfter(SBOItemEventArg pVal)
        {

            if (DataForCalculationRate.IsCalculated && (pVal.FormMode == 2 || pVal.FormMode == 3))
            {
                Form invoiceForm = Application.SBO_Application.Forms.ActiveForm;
                var txtRate = (SAPbouiCOM.EditText)(invoiceForm.Items.Item("64").Specific);
                try
                {
                    txtRate.Value = Math.Round(decimal.Parse(DataForCalculationRate.GlobalRate), 6).ToString();
                }
                catch (Exception e)
                {

                }
                DataForCalculationRate.IsCalculated = false;
            }
        }

        private ButtonCombo ButtonCombo0;

        private void ButtonCombo0_PressedAfter(object sboObject, SBOItemEventArg pVal)
        {
            if (string.IsNullOrWhiteSpace(DataForCalculationRate.DownPaymentAmount))
            {
                stop = true;
                return;
            }
            if (pVal.ActionSuccess && !stop)
            {
                DownPaymentLogic.DownPaymentLogic.CorrectionJournalEntryUI(Program._comp, 133, DataForCalculationRate.CardCode,
                    DataForCalculationRate.DownPaymentAmount, DataForCalculationRate.DocNum, DataForCalculationRate.BplName, DataForCalculationRate.GlAccount, Program.ExchangeGain, Program.ExchangeLoss, DataForCalculationRate.PostingDate);
            }

        }

        private void ButtonCombo0_PressedBefore(object sboObject, SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            Form arInoviceForm = SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm;
            if (pVal.FormMode != 3)
            {
                stop = true;
                return;
            }
            stop = false;

            try
            {


                DataForCalculationRate.BusinesPartnerName = ((SAPbouiCOM.EditText)(arInoviceForm.Items
                    .Item("4").Specific)).Value;
                try
                {
                    DataForCalculationRate.BplName = ((SAPbouiCOM.ComboBox)(arInoviceForm.Items //es iwvevs catches
                           .Item("2001").Specific)).Selected.Description;
                }
                catch (Exception)
                {

                    // db without branch
                }

                var docNumInvItm = arInoviceForm.Items.Item("8");
                var docNumInvEditText = (SAPbouiCOM.EditText)docNumInvItm.Specific;
                DataForCalculationRate.DocNum = docNumInvEditText.Value;

                var cardCodeItem = arInoviceForm.Items.Item("4");
                var cardCodeEditText = (SAPbouiCOM.EditText)cardCodeItem.Specific;
                DataForCalculationRate.CardCode = cardCodeEditText.Value;

                var totalDownPaymentItem = arInoviceForm.Items.Item("204");
                var totalDownPaymentEditText = (SAPbouiCOM.EditText)totalDownPaymentItem.Specific;
                DataForCalculationRate.DownPaymentAmount = totalDownPaymentEditText.Value;
                DataForCalculationRate.FormTypex = "133";

                var txtTotalWithCurr =
                    (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("22").Specific); //totalInv before discount
                var vatWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("27").Specific);
                var discountWithCurr = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("42").Specific);

                string txtTotal = txtTotalWithCurr.Value.Split(' ')[0];
                string vat = vatWithCurr.Value.Split(' ')[0];
                string discount = discountWithCurr.Value.Split(' ')[0] == string.Empty
                    ? "0"
                    : discountWithCurr.Value.Split(' ')[0];


                SBObob sbObob = (SBObob)Program._comp.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                string currency = ((ComboBox)SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Items
                    .Item("63").Specific).Value;
                DataForCalculationRate.DocCurrency = currency;
                if (currency == "GEL")
                {
                    stop = true;
                    return;
                }
                string postingDateString = ((EditText)arInoviceForm.Items.Item("10").Specific).Value;
                DateTime postingDate = DateTime.ParseExact(postingDateString, "yyyyMMdd", null);
                decimal currencyValue =
                    Math.Round(
                        decimal.Parse(sbObob.GetCurrencyRate(currency, postingDate).Fields.Item(0).Value
                            .ToString()), 6);
                ((SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific)).Value =
                    currencyValue.ToString();
                DateTime docDate = DateTime.ParseExact(
                    ((EditText)(arInoviceForm.Items.Item("10").Specific)).Value,
                    "yyyyMMdd", CultureInfo.InvariantCulture);
                var txtRate = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("Item_0000").Specific);
                DataForCalculationRate.RateInv = decimal.Parse(txtRate.Value.Replace(",", "."));
                DataForCalculationRate.RateInv = decimal.Parse(UiManager.GetCurrencyRate(DataForCalculationRate.DocCurrency, docDate, Program._comp).ToString());
                Matrix invMAtrix = (Matrix)arInoviceForm.Items.Item("38").Specific;
                try
                {
                    decimal taxTotalAmount = 0;
                    var vatFromTotalBefore = decimal.Parse(txtTotal);
                    for (int i = 1; i < invMAtrix.RowCount; i++)
                    {
                        EditText txtMoney = (SAPbouiCOM.EditText)invMAtrix.Columns.Item("33").Cells.Item(i).Specific; //Tax Amount Line
                        decimal taxLineAmount = decimal.Parse(txtMoney.Value.Split(' ')[0]);
                        taxTotalAmount += taxLineAmount;
                    }

                    DataForCalculationRate.TotalInv = vatFromTotalBefore + taxTotalAmount;

                }
                catch (Exception e)
                {
                    SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message,
                        BoMessageTime.bmt_Short, true);
                    SAPbouiCOM.Framework.Application.SBO_Application.Forms.ActiveForm.Close();
                }
            }
            catch (Exception e)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.SetStatusBarMessage(e.Message, BoMessageTime.bmt_Short,
                    true);
                //branch araa
            }
            try

            {
                var glAccount = (SAPbouiCOM.EditText)(arInoviceForm.Items.Item("2010").Specific);
                DataForCalculationRate.GlAccount = glAccount.Value;
                DownPaymentLogic.DownPaymentLogic.ExchangeRateCorrectionUi(DataForCalculationRate, Program._comp);

            }
            catch (Exception e)
            {
                Application.SBO_Application.SetStatusBarMessage(e.Message,
                    BoMessageTime.bmt_Short, true);
            }

        }
    }
}
