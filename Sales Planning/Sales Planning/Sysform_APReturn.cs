using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.CHY
{
    class Sysform_APReturn
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
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    SAPbouiCOM.UserDataSource uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    uds.Value = "";
                    //SAPbouiCOM.Item oItem;
                    //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //((SAPbouiCOM.EditText)oItem.Specific).String = "";
                    //oItem.Left = 160;
                    //oItem.Width = 10;
                    //oItem.Top = oForm.Height - 60;
                    //oItem.Height = 20;
                    ////oItem.Enabled = false;
                    ////oItem.Visible = false;
                    //oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    //oForm.Items.Item("FUID").Visible = false;
                }
                if (oForm.DataSources.UserDataSources.Count > 0)
                {
                    if (oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString() != "")
                    {
                        string FUID = oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString();
                        SAPbouiCOM.Form oSForm = SAP.SBOApplication.Forms.Item(FUID);
                        oSForm.Select();
                        BubbleEvent = false;
                    }
                }
                //if (((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).Value != "")
                //{
                //    string FUID = ((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).Value.ToString();
                //    SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FUID);
                //    oSForm.Select();
                //    BubbleEvent = false;
                //} 
                
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "38")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.Row >= 0 && pVal.ColUID == "256")
                                {
                                    int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("RPD1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                                    int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("RPD1").GetValue("LINENUM", pVal.Row - 1).ToString());

                                    InitForm.TEXT(oForm.UniqueID, docentry, linenum, pVal.Row, "RPD1", "38", oForm.DataSources.DBDataSources.Item("RPD1").GetValue("TEXT", pVal.Row - 1).ToString());
                                    BubbleEvent = false;
                                }
                            }
                        }
                        break;
                }
            }            
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "38")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.Row >= 0 && pVal.ColUID != "256")
                                {
                                    int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("RPD1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                                    int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("RPD1").GetValue("LINENUM", pVal.Row - 1).ToString());
                                    //if (pVal.ColUID == "U_BookNo")
                                    //{
                                    //    InitForm.CONM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APDOC", "U_CONNO,U_BookNo");
                                    //}
                                    //else if (pVal.ColUID == "U_ProdType")
                                    //{
                                    //    InitForm.DOPTM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APDOPT", "");
                                    //}
                                    //else
                                    //{
                                    InitForm.SDM(oForm.UniqueID, docentry, linenum, pVal.Row, "RPD1", "38");
                                    //}
                                }
                            }
                        }
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
            return;
        }
    }
}
