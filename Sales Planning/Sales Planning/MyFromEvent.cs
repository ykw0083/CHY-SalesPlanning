using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.CHY
{
    class MyFromEvent
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                string formtype = oForm.TypeEx;
                string cardcode = "";
                string docnum = "";

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if ((formtype == "1250000100" && pVal.ItemUID == "1250000001" && (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)) ||
                        (pVal.ItemUID == "1" && oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE))
                    {

                        if (formtype == "FT_SPLAN" || formtype == "FT_TPPLAN")
                        {
                            cardcode = oForm.DataSources.DBDataSources.Item("@" + formtype).GetValue("U_CardCode", 0);
                            if (formtype == "FT_SPLAN") docnum = oForm.DataSources.DBDataSources.Item("@" + formtype + "1").GetValue("U_SODOCNUM", 0);
                            if (formtype == "FT_TPPLAN") docnum = oForm.DataSources.DBDataSources.Item("@" + formtype + "1").GetValue("U_SPDOCNUM", 0);
                        }
                        else if (formtype == "139" || formtype == "149")
                        {
                            cardcode = oForm.DataSources.DBDataSources.Item(0).GetValue("CardCode", 0);
                        }
                        else if (formtype == "1250000100")
                        {
                            cardcode = oForm.DataSources.DBDataSources.Item(0).GetValue("BpCode", 0);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (oForm.DataSources.DBDataSources.Item(0).GetValue("Status", 0) == "A")
                                {
                                    SAPbobsCOM.Recordset rs = SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                                    string absid = oForm.DataSources.DBDataSources.Item(0).GetValue("AbsID", 0);
                                    rs.DoQuery("select Status from OOAT where Status <> 'A' and AbsID = " + absid);
                                    if (rs.RecordCount > 0)
                                    { }
                                    else
                                        return;
                                }
                                else
                                    return;

                            }
                        }
                        else
                        {
                            return;
                        }

                        if (cardcode != null && cardcode.Trim() != "")
                        {
                            SAPbouiCOM.Item oItem = null;
                            SAPbouiCOM.ComboBox oComboapp = null;
                            SAPbouiCOM.EditText oEdit = null;

                            int cnt = 0;
                            int temp = 0;
                            string NotifyV = "";
                            string errMsg = "";

                            #region Overdue
                            string documentnum = "", documentdate = "", documentduedate = "";
                            string sql = "Select U_NotifyV from [@FT_SPODCTRL] where Code='" + formtype + "'";
                            SAPbobsCOM.Recordset query = SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            SAPbobsCOM.Recordset rs = SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            query.DoQuery(sql);
                            if (query.RecordCount > 0)
                            {
                                query.MoveFirst();
                                NotifyV = query.Fields.Item(0).Value.ToString();
                                if (NotifyV != "NA")
                                {
                                    if (formtype == "FT_SPLAN" || formtype == "FT_TPPLAN")
                                        cnt = ft_Functions.CheckCreditTerm(oForm, oForm.DataSources.DBDataSources.Item("@" + formtype), oForm.DataSources.DBDataSources.Item("@" + formtype + 1), ref documentnum, ref documentdate, ref documentduedate, ref errMsg);
                                    else
                                        cnt = ft_Functions.CheckCreditTerm(oForm, oForm.DataSources.DBDataSources.Item(0), oForm.DataSources.DBDataSources.Item(1), ref documentnum, ref documentdate, ref documentduedate, ref errMsg);
                                    if (cnt == -1)
                                    {
                                        BubbleEvent = false;
                                        SAP.SBOApplication.MessageBox("Overdue SP Error");
                                    }
                                    else if (cnt > 0)
                                    {
                                        if (formtype == "139" || formtype == "149")
                                        {
                                            oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("U_CTERM").Specific;
                                            oEdit.String = "Y";
                                        }
                                        if (NotifyV == "MSG_BLOCK")
                                        {
                                            temp = ft_Functions.CheckSPNeeded("OD", formtype, docnum);
                                            if (temp == -1)
                                            {
                                                BubbleEvent = false;
                                                return;
                                            }
                                            if (formtype == "FT_SPLAN" || formtype == "FT_TPPLAN")
                                            {
                                                if (temp > 0)
                                                {
                                                    oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APP", 0, "W");
                                                    //oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                                    //oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                                    //oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));
                                                }
                                            }
                                            else if (formtype == "139" || formtype == "149")
                                            {
                                                oItem = oForm.Items.Item("U_APP");
                                                oComboapp = (SAPbouiCOM.ComboBox)oItem.Specific;
                                                oComboapp.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                            }
                                        }
                                        SAP.SBOApplication.MessageBox($"Overdue invoices {Environment.NewLine}Oldest Invoice# {documentnum}{Environment.NewLine}Dated {documentdate}{Environment.NewLine}Due By {documentduedate}"
                                            , 1, "Ok", "", "");

                                    }

                                }
                            }
                            #endregion
                            #region credit limit
                            string limitType = "";
                            double different = 0, c_usage = 0, t_limit = 0, c_limit = 0;
                            sql = "Select U_NotifyV from [@FT_SPCLCTRL] where Code='" + formtype + "'";
                            query = SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset) as SAPbobsCOM.Recordset;
                            query.DoQuery(sql);
                            if (query.RecordCount > 0)
                            {
                                query.MoveFirst();
                                NotifyV = query.Fields.Item(0).Value.ToString();
                                if (NotifyV != "NA")
                                {
                                    if (formtype == "FT_SPLAN" || formtype == "FT_TPPLAN")
                                        cnt = ft_Functions.CheckCreditLimit(oForm, oForm.DataSources.DBDataSources.Item("@" + formtype), oForm.DataSources.DBDataSources.Item("@" + formtype + "1"), ref errMsg, ref limitType, ref different, ref c_usage, ref t_limit, ref c_limit);
                                    else
                                        cnt = ft_Functions.CheckCreditLimit(oForm, oForm.DataSources.DBDataSources.Item(0), oForm.DataSources.DBDataSources.Item(1), ref errMsg, ref limitType, ref different, ref c_usage, ref t_limit, ref c_limit);
                                    if (cnt == -1)
                                    {
                                        BubbleEvent = false;
                                        SAP.SBOApplication.MessageBox("Credit Limit SP Error");
                                    }
                                    else if (cnt > 0)
                                    {
                                        if (formtype == "139" || formtype == "149")
                                        {
                                            oItem = oForm.Items.Item("U_CUsage");
                                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                                            oEdit.Value = c_usage.ToString();

                                            oItem = oForm.Items.Item("U_TLimit");
                                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                                            oEdit.Value = t_limit.ToString();

                                            oItem = oForm.Items.Item("U_CLimit");
                                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                                            oEdit.Value = c_limit.ToString();
                                        }
                                        if (NotifyV == "MSG_BLOCK")
                                        {
                                            temp = ft_Functions.CheckSPNeeded("CL", formtype, docnum);
                                            if (temp == -1)
                                            {
                                                BubbleEvent = false;
                                                return;
                                            }
                                            if (formtype == "FT_SPLAN" || formtype == "FT_TPPLAN")
                                            {
                                                if (temp > 0)
                                                {

                                                    oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APP", 0, "W");
                                                    //oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                                    //oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                                    //oForm.DataSources.DBDataSources.Item("@" + formtype).SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));
                                                }
                                            }
                                            else if (formtype == "139" || formtype == "149")
                                            {
                                                if (temp > 0)
                                                {
                                                    oItem = oForm.Items.Item("U_APP");
                                                    oComboapp = (SAPbouiCOM.ComboBox)oItem.Specific;
                                                    oComboapp.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                                }
                                            }

                                        }
                                        SAP.SBOApplication.MessageBox($"Credit Limit Exceeded {Environment.NewLine}Limit Type - {limitType}{Environment.NewLine}Total Credit Limit Use - RM {c_usage.ToString("#,###,###,###.00")}{Environment.NewLine}Credit Limit Amount - RM {c_limit.ToString("#,###,###,###.00")}{Environment.NewLine}Over Limit Amount - RM {different.ToString("#,###,###,###.00")}", 1, "Ok", "", "");
                                    }

                                }
                            }
                            #endregion
                        }

                    }
                }
            }
            catch (Exception ex)
            {

                SAP.SBOApplication.MessageBox($"Item Event Before {ex.Message} {System.Environment.NewLine} {ex.StackTrace}", 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }
    }
}
