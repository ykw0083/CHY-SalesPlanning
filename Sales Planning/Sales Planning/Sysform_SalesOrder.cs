using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.CHY
{
    class Sysform_SalesOrder
    {
        public static void processRightClickEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processRightClickEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal)
        {
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
                        oEdit.DataBind.SetBound(true, "ORDR", "U_CUsage");


                        oItem = oForm.Items.Add("U_TLimit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "ORDR", "U_TLimit");

                        oItem = oForm.Items.Add("U_CLimit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "ORDR", "U_CLimit");

                        oItem = oForm.Items.Add("U_DSAPP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                        oItem.Left = oForm.Width + 100;
                        oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                        oCombo.DataBind.SetBound(true, "ORDR", "U_DSAPP");

                        oItem = oForm.Items.Add("U_APP", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                        oItem.Left = oForm.Width + 100;
                        oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                        oCombo.DataBind.SetBound(true, "ORDR", "U_APP");

                        oItem = oForm.Items.Add("U_ADDONIND", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "ORDR", "U_ADDONIND");

                        oItem = oForm.Items.Add("U_CTERM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        oItem.Left = oForm.Width + 100;
                        oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                        oEdit.DataBind.SetBound(true, "ORDR", "U_CTERM");

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
        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                //long docentry = 0;
                //string docnum = "";
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //{
                        //    if (pVal.ItemUID == "btncombo")
                        //    {
                        //        SAPbouiCOM.ButtonCombo oButtonCombo = (SAPbouiCOM.ButtonCombo)oForm.Items.Item(pVal.ItemUID).Specific;
                                
                        //        if (oButtonCombo.Selected.Value == "SHIPD")
                        //        {
                        //            docentry = long.Parse(oForm.DataSources.DBDataSources.Item(0).GetValue("docentry", 0).ToString());
                        //            docnum = oForm.DataSources.DBDataSources.Item(0).GetValue("docnum", 0).ToString();
                        //            InitForm.shipdoc(oForm.UniqueID, docentry, 0, docnum);
                        //        }
                        //        else if (oButtonCombo.Selected.Value == "SHIPL")
                        //        {
                        //            docentry = long.Parse(oForm.DataSources.DBDataSources.Item(0).GetValue("docentry", 0).ToString());
                        //            InitForm.shiplist(oForm.UniqueID, docentry);
                        //        }
                        //    }
                        //    else if (pVal.ItemUID == "APSA")
                        //    {
                        //        InitForm.shipingadvice(oForm.UniqueID, true);
                        //    }
                        //}
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        //if (pVal.ItemUID == "38")
                        //{
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //    {
                        //        if (pVal.Row >= 0)
                        //        {
                        //            docentry = int.Parse(oForm.DataSources.DBDataSources.Item("RDR1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                        //            int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("RDR1").GetValue("LINENUM", pVal.Row - 1).ToString());
                        //            string bookno = oForm.DataSources.DBDataSources.Item("RDR1").GetValue("U_BookNo", pVal.Row - 1).Trim();
                        //            if (pVal.ColUID == "U_FCL")
                        //            {
                        //                InitForm.CONM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APSOC", "U_CONNO", bookno);
                        //            }
                        //            else if (pVal.ColUID == "256")
                        //            {
                        //            }
                        //            else
                        //            {
                        //                decimal del = decimal.Parse(oForm.DataSources.DBDataSources.Item("RDR1").GetValue("DelivrdQty", pVal.Row - 1).ToString());
                        //                if (del > 0 && pVal.ColUID != "0")
                        //                    InitForm.SOM(oForm.UniqueID, docentry, linenum, pVal.Row, "RDR1", "38");
                        //                else
                        //                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Popup window disable!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                        //            }
                        //        }
                        //    }
                        //}
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processMenuEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            return;
        }
        public static void processMenuEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal)
        {
            try
            {
                if (pVal.MenuUID == "1287") // Duplicate
                {

                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Menu Event After " + ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processDataEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        break; // 20210531
                        /*
                        SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("ORDR");
                        SAPbouiCOM.DBDataSource ods1 = oForm.DataSources.DBDataSources.Item("RDR1");

                        SAPbouiCOM.Item oItem = null;
                        SAPbouiCOM.EditText oEdit = null;
                        SAPbouiCOM.ComboBox oCombo = null;
                        SAPbouiCOM.ComboBox oComboapp = null;

                        oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("81").Specific;
                        if (oCombo.Selected.Value == "6")
                        {
                            break;
                        }
                        string errMsg = "", limitType = "", whscode = "";
                        double different = 0, c_usage = 0, t_limit = 0, c_limit = 0;
                        Boolean result = false;
                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //SAPbobsCOM.ApprovalRequestsService App;
                        //SAPbobsCOM.ApprovalRequest AppReq;
                        //SAPbobsCOM.ApprovalRequestsParams AppRequests;
                        //SAPbobsCOM.ApprovalRequestParams AppRequest;

                        //App = (SAPbobsCOM.ApprovalRequestsService)SAP.SBOCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
                        //AppReq = (SAPbobsCOM.ApprovalRequest)App.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
                        //AppRequests = (SAPbobsCOM.ApprovalRequestsParams)App.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
                        //AppRequest = (SAPbobsCOM.ApprovalRequestParams)App.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);

                        //AppRequests = App.GetAllApprovalRequestsList();
                        //AppRequest = AppRequests.Item(AppRequests.Count - 1);
                        //string appcode = AppRequest.Code.ToString();

                        //AppReq = App.GetApprovalRequest(AppRequest);
                        //AppReq.ApprovalRequestDecisions.Add();
                        //AppReq.ApprovalRequestDecisions.Item(0).Remarks = "Approved";
                        //AppReq.ApprovalRequestDecisions.Item(0).Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardApproved;
                        //App.UpdateRequest(AppReq);

                        //    BubbleEvent = false;
                        for (int i = 0; i < ods1.Size; i++)
                        {
                            whscode = ods1.GetValue("whscode", i).ToString();
                            rs.DoQuery("select * from owhs where whscode='" + whscode + "' and isnull(U_Dropship,'N') = 'Y'");
                            if (rs.RecordCount > 0)
                            {
                                result = true;
                                //break;
                            }
                        }
                        //if (result)
                        {
                            oItem = oForm.Items.Item("U_DSAPP");
                            oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;

                            oItem = oForm.Items.Item("U_APP");
                            oComboapp = (SAPbouiCOM.ComboBox)oItem.Specific;
                            if (ft_Functions.CheckCreditTerm(oForm, ods, ods1, ref errMsg))
                            {

                                oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("U_CTERM").Specific;
                                oEdit.String = "Y";
                                if (result)
                                {
                                    if (ObjectFunctions.Approval("112"))
                                        oCombo.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    if (ObjectFunctions.Approval("17"))
                                        oComboapp.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    #region remove due to 20210305 update
                                    //SAP.SBOApplication.MessageBox("There is/are invoices overdue for this customer", 1, "Ok", "", "");
                                    #endregion
                                }
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
                                if (result)
                                {
                                    if (ObjectFunctions.Approval("112"))
                                        oCombo.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    if (ObjectFunctions.Approval("17"))
                                        oComboapp.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                    #region remove due to 20210305 update
                                    //SAP.SBOApplication.MessageBox("Credit Limit Exceeded " + Environment.NewLine + "Limit Type - " +
                                    //    limitType + Environment.NewLine + " Over Limit Amount - RM " + different.ToString("#,###,###,###.00"), 1, "Ok", "", "");
                                    #endregion
                                }
                            }
                            
                            //ods.SetValue("U_CUsage", 0, c_usage.ToString());
                            //ods.SetValue("U_TLimit", 0, t_limit.ToString());
                            //ods.SetValue("U_CLimit", 0, c_limit.ToString());

                            oItem = oForm.Items.Item("U_CUsage");
                            //oItem = oForm.Items.Add("U_CUsage", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //oItem.Left = oForm.Width + 100;
                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                            oEdit.DataBind.SetBound(true, "ORDR", "U_CUsage");
                            oEdit.Value = c_usage.ToString();


                            oItem = oForm.Items.Item("U_TLimit");
                            //oItem = oForm.Items.Add("U_TLimit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //oItem.Left = oForm.Width + 100;
                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                            oEdit.DataBind.SetBound(true, "ORDR", "U_TLimit");
                            oEdit.Value = t_limit.ToString();

                            oItem = oForm.Items.Item("U_CLimit");
                            //oItem = oForm.Items.Add("U_CLimit", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            //oItem.Left = oForm.Width + 100;
                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                            oEdit.DataBind.SetBound(true, "ORDR", "U_CLimit");
                            oEdit.Value = c_limit.ToString();

                            oItem = null;
                            oEdit = null;
                        }
                        //BubbleEvent = false;
                        */

                break;
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Data Event Before " + ex.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }

        }
     
    }
}
