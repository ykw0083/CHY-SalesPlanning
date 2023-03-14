using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class Sysform_SalesQuotation
    {
        public static void processDataEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        break;
                        /*
                        SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("OQUT");
                        SAPbouiCOM.DBDataSource ods1 = oForm.DataSources.DBDataSources.Item("QUT1");
                        string errMsg = "", limitType = "";
                        double different = 0, c_usage = 0, t_limit = 0, c_limit = 0;
                        if (ft_Functions.CheckCreditTerm(oForm, ods, ods1, ref errMsg))
                        {
                            SAP.SBOApplication.MessageBox("There is/are invoices overdue for this customer", 1, "Ok", "", "");
                        }
                        errMsg = "";
                        int cnt = 0;
                        cnt = ft_Functions.CheckCreditLimit(oForm, ods, ods1, ref errMsg, ref limitType, ref different, ref c_usage, ref t_limit, ref c_limit);
                        if (cnt == -1)
                        {
                            BubbleEvent = false;
                            break;
                        }
                        else if (cnt >= 1)
                        {
                            SAP.SBOApplication.MessageBox("Credit Limit Exceeded " + Environment.NewLine + "Limit Type - " +
                            limitType + Environment.NewLine + " Over Limit Amount - RM " + different.ToString("#,###,###,###.00"), 1, "Ok", "", "");
                        }
                        //BubbleEvent = false;
                        break;
                        */
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Data Event Before " + ex.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }

        }
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        SAPbouiCOM.Item oItem = null;
                        SAPbouiCOM.EditText oEdit = null;
                        SAPbouiCOM.ComboBox oCombo = null;

                        oItem = oForm.Items.Add("U_CUsage", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "OQUT", "U_CUsage");


                        oItem = oForm.Items.Add("U_TLimit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "OQUT", "U_TLimit");

                        oItem = oForm.Items.Add("U_CLimit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "OQUT", "U_CLimit");

                        oItem = oForm.Items.Add("U_DSAPP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                        oItem.Left = oForm.Width + 100;
                        oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                        oCombo.DataBind.SetBound(true, "OQUT", "U_DSAPP");

                        oItem = oForm.Items.Add("U_APP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                        oItem.Left = oForm.Width + 100;
                        oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                        oCombo.DataBind.SetBound(true, "OQUT", "U_APP");

                        oItem = oForm.Items.Add("U_ADDONIND", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "OQUT", "U_ADDONIND");

                        oItem = oForm.Items.Add("U_CTERM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "OQUT", "U_CTERM");

                        oItem = null;
                        oEdit = null;
                        //SAPbouiCOM.Item oItem = null;
                        //SAPbouiCOM.Button oButton = null;
                        //SAPbouiCOM.ButtonCombo oButtonCombo = null;

                        //oItem = oForm.Items.Add("btncombo", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                        //oItem.DisplayDesc = true;
                        //oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5;
                        //oItem.Width = 80;
                        //oItem.Top = oForm.Items.Item("2").Top;
                        //oItem.Height = oForm.Items.Item("2").Height;
                        //oItem.AffectsFormMode = false;
                        //oButtonCombo = (SAPbouiCOM.ButtonCombo)oItem.Specific;
                        //oButtonCombo.Caption = "Shipping Doc";
                        //oButtonCombo.ValidValues.Add("SHIPD", "Shipping Doc");
                        //oButtonCombo.ValidValues.Add("SHIPL", "Shipping List");
                        //oButtonCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);

                        ////oItem = oForm.Items.Add("SHIPD", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        ////oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5;
                        ////oItem.Width = 80;
                        ////oItem.Top = oForm.Items.Item("2").Top;
                        ////oItem.Height = oForm.Items.Item("2").Height;
                        ////oButton = (SAPbouiCOM.Button)oItem.Specific;
                        ////oButton.Caption = "Shipping Doc";

                        ////oItem = oForm.Items.Add("SHIPL", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        ////oItem.Left = oForm.Items.Item("SHIPD").Left + oForm.Items.Item("SHIPD").Width + 5;
                        ////oItem.Width = 80;
                        ////oItem.Top = oForm.Items.Item("2").Top;
                        ////oItem.Height = oForm.Items.Item("2").Height;
                        ////oButton = (SAPbouiCOM.Button)oItem.Specific;
                        ////oButton.Caption = "Shipping List";

                        //oItem = oForm.Items.Add("APSA", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        //oItem.Left = oForm.Items.Item("btncombo").Left + oForm.Items.Item("btncombo").Width + 5;
                        //oItem.Width = 80;
                        //oItem.Top = oForm.Items.Item("2").Top;
                        //oItem.Height = oForm.Items.Item("2").Height;
                        //oButton = (SAPbouiCOM.Button)oItem.Specific;
                        //oButton.Caption = "Shipping Advise";

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("U_ADDONIND").Specific;
                                oEdit.String = "Y";
                            }
                        }
                        break;

                }
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }
    }
}
