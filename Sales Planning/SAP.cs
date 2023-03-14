using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON
{
    class SAP
    {
        const bool bDebugMode = true;

        public static SAPbouiCOM.Application SBOApplication;
        public static SAPbobsCOM.Company SBOCompany;
        public static int formUID;
        public static SAPbouiCOM.Form statusForm;
        public static int currentMatrixRow;

        public static void createStatusForm()
        {
            try
            {
                SAPbouiCOM.Form oForm = null;
                SAPbouiCOM.FormCreationParams oCreationParams = null;
                SAPbouiCOM.Item oItem = null;
                oCreationParams = ((SAPbouiCOM.FormCreationParams)(SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams)));

                oCreationParams.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Floating;
                oCreationParams.UniqueID = "FT_StatusForm";
                oCreationParams.FormType = "StatusForm";
                oForm = SBOApplication.Forms.AddEx(oCreationParams);
                //statusForm.Visible = false;
                oForm.ClientWidth = 450;
                oForm.ClientHeight = 80;
                oForm.Top = SBOApplication.Desktop.Top + 50;
                oForm.Left = (SBOApplication.Desktop.Width - 450) / 2;
                oItem = oForm.Items.Add("lblStatus", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 10;
                oItem.Height = 15;
                oItem.Top = (oForm.ClientHeight - oItem.Height) / 2;
                oItem.Width = oForm.ClientWidth - 20;

                oForm = null;
                GC.WaitForPendingFinalizers();
            }
            catch { }
        }

        public static void showStatus()
        {
            try
            {
                if (statusForm != null) statusForm.Visible = true;
            }
            catch { }
        }
        public static void hideStatus()
        {
            try
            {
                if (statusForm != null)
                {
                    statusForm.Visible = false;
                    ((SAPbouiCOM.StaticText)(statusForm.Items.Item("").Specific)).Caption = "";
                }
            }
            catch { }
        }
        public static void getStatusForm()
        {
            try
            {
                statusForm = SBOApplication.Forms.Item("FT_StatusForm");
            }
            catch { statusForm = null; };
        }
        public static void setStatus(string message)
        {
            if (statusForm != null)
            {
                try
                {
                    if (!statusForm.Visible) statusForm.Visible = true;
                    ((SAPbouiCOM.StaticText)(statusForm.Items.Item("lblStatus").Specific)).Caption = message;
                }
                catch (Exception ex)
                {
                    SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                }
            }
        }

        public static int getNewformUID()
        {
            formUID++;
            do
            {
                if (getUID()) break;
                formUID++;
            } while (true);
            return formUID;
        }

        private static Boolean getUID()
        {
            SAPbouiCOM.Form newForm = null;
            try
            {
                newForm = SBOApplication.Forms.Item("FT_" + formUID.ToString());
            }
            catch { }
            if (newForm == null) return true;
            newForm = null;
            return false;
        }

        public static void showDebugMessage(string Text, int defaultBtn, string btn1Caption, string btn2Caption, string btn3Caption)
        {
            if (bDebugMode) SBOApplication.MessageBox(Text, defaultBtn, btn1Caption, btn2Caption, btn3Caption);
        }

        public static void showDebugMessage(string Text)
        {
            showDebugMessage(Text, 1, "Close", "", "");
        }


        public static void setApplication()
        {
            try
            {
                SAPbouiCOM.SboGuiApi SboGuiApi = null;
                string sConnectionString = null;
                SboGuiApi = new SAPbouiCOM.SboGuiApi();

                // Connect to running SBO Application
                //sConnectionString = System.Convert.ToString(Environment.GetCommandLineArgs().GetValue(1));
                if ((System.Environment.GetCommandLineArgs().Length == 1))
                {
                    sConnectionString = "0030002C0030002C00530041005000420044005F00440061007400650076002C0050004C006F006D0056004900490056";
                }
                else
                {
                    sConnectionString = System.Environment.GetCommandLineArgs().GetValue(1).ToString();
                }
                // Fast Track SBOi AddOn License Key
                //SboGuiApi.AddonIdentifier = "56455230354241534953303030363030303439303A4C30353436383833333837BE2D8E3EA0DBD35826EF326077F8A12A43680561";

                SboGuiApi.Connect(sConnectionString);

                // Get an instantialized application object
                SBOApplication = SboGuiApi.GetApplication(-1);
            }

            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
                System.Environment.Exit(0);
            }
        }

        public static int setConnectionContext()
        {
            int setConnectionContextReturn = 0;
            string sCookie = null;
            string sConnectionContext = null;

            // Get connection context from context cookie from DIAPI
            //SBOCompany = new SAPbobsCOM.Company();
            SBOCompany = new SAPbobsCOM.Company();
            sCookie = SBOCompany.GetContextCookie();
            sConnectionContext = SBOApplication.Company.GetConnectionContext(sCookie);


            if (SBOCompany.Connected == true) SBOCompany.Disconnect();

            // Set connection context to the DIAPI
            setConnectionContextReturn = SBOCompany.SetSboLoginContext(sConnectionContext);
            return setConnectionContextReturn;
        }

        public static int connectToCompany()
        {
            // Connect to SBO company database
            int connectToCompanyReturn = 0;

            connectToCompanyReturn = SBOCompany.Connect();
            return connectToCompanyReturn;
        }
    }
}
