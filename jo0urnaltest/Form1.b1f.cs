using System;
using System.Collections.Generic;
using System.Xml;
using SAPbouiCOM.Framework;

namespace jo0urnaltest
{
    [FormAttribute("jo0urnaltest.Form1", "Form1.b1f")]
    class Form1 : UserFormBase
    {
        public Form1()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("Item_0").Specific));
            this.Button0.ClickBefore += new SAPbouiCOM._IButtonEvents_ClickBeforeEventHandler(this.Button0_ClickBefore);
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {

        }

        private void Button0_ClickBefore(object sboObject, SAPbouiCOM.SBOItemEventArg pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var comp = SAPbouiCOM.Framework.Application.SBO_Application.Company;
            var company = (SAPbobsCOM.Company)comp.GetDICompany();
            SAPbobsCOM.JournalEntries vJE = ( SAPbobsCOM.JournalEntries)company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                        
            vJE.ReferenceDate = new DateTime(2016, 10, 25);

            vJE.DueDate = new DateTime(2016, 10, 25);

            vJE.TaxDate = new DateTime(2016, 10, 25);

            vJE.Memo = "Test Message";
            //vJE.TransactionCode = "13";
            //vJE.Reference = "406";


           // vJE.TransactionCode = "INR";

            vJE.Lines.BPLID = 235;

            vJE.Lines.Credit = 3000;

            vJE.Lines.Debit = 0;
            vJE.Lines.AccountCode = "1210";



            vJE.Lines.Add();

            vJE.Lines.AccountCode = "8340";

            vJE.Lines.Credit = 0;

            vJE.Lines.Debit = 3000;



            vJE.Lines.Add();



            int i = vJE.Add();

            if (i == 0)
            {

                ;

                return;

            }

            else
            {

                string des = company.GetLastErrorDescription();

              

            }

        }
    }
}