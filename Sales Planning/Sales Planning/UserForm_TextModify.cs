using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class UserForm_TextModify
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                SAPbouiCOM.Form oSForm;
                string oFormId;
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        oFormId = oForm.DataSources.UserDataSources.Item("FUID").Value.ToString();
                        oSForm = SAP.SBOApplication.Forms.Item(oFormId);
                        oSForm.DataSources.UserDataSources.Item("cfluid").Value = "";
                        oSForm.Select();

                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            string DocEntry = oForm.DataSources.UserDataSources.Item("DocEntry").Value.ToString();
                            string LineNum = oForm.DataSources.UserDataSources.Item("LineNo").Value.ToString();
                            string DSNAME = oForm.DataSources.UserDataSources.Item("DSNAME").Value.ToString();
                            //string DocEntry = ((SAPbouiCOM.EditText)oForm.Items.Item("DocEntry").Specific).Value.ToString();
                            //string LineNum = ((SAPbouiCOM.EditText)oForm.Items.Item("LineNo").Specific).Value.ToString();
                            string text = ((SAPbouiCOM.EditText)oForm.Items.Item("TEXT").Specific).Value.ToString();
                            text = text.Replace("'", "''");
                            //string DSNAME = ((SAPbouiCOM.EditText)oForm.Items.Item("DSNAME").Specific).Value.ToString();
                            string sql = "";
                            SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("INV1");

                            //SAPbobsCOM.Recordset rss = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //rss.DoQuery("SELECT TOP 1 LineId FROM [@" + dsname + "] WHERE DocEntry = " + docno.ToString() + " ORDER BY LineId DESC");
                            //if (rss.RecordCount > 0)
                            //{
                            //    rss.MoveFirst();
                            //    lineidnew = int.Parse(rss.Fields.Item(0).Value.ToString());
                            //}
                            //else
                            //{
                            //    lineidnew = 0;
                            //}

                            FT_ADDON.SAP.SBOCompany.StartTransaction();
                            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            sql = "UPDATE [" + DSNAME + "] SET Text = '" + text + "' WHERE DocEntry = " + DocEntry.ToString() + " AND LineNum = " + LineNum.ToString();
                            rs.DoQuery(sql);
                            if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            //oForm.Close();
                            //SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FUID);
                            //oSForm.Freeze(false);
                            //oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                            //oSForm.Select();

                        }
                        if (pVal.ItemUID == "1" || pVal.ItemUID == "2")
                        {
                            oFormId = oForm.DataSources.UserDataSources.Item("FUID").Value.ToString();
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || pVal.ItemUID == "2")
                            {
                                oForm.Close();
                                oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(oFormId);
                                oSForm.DataSources.UserDataSources.Item("cfluid").Value = "";
                                oSForm.Select();
                                SAP.SBOApplication.ActivateMenuItem("1289");
                                SAP.SBOApplication.ActivateMenuItem("1288");

                            }
                        }
                        break;
                    default:
                        break;
                }

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
