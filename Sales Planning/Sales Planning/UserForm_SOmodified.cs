using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_SOmodified
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            if (!validateRows(oForm))
                            {
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "1" || pVal.ItemUID == "2")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || pVal.ItemUID == "2")
                            {
                                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(oForm.DataSources.UserDataSources.Item("fuid").Value.ToString());
                                //SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).String);
                                
                                oSForm.Freeze(false);
                                oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                                oSForm.Select();

                                SAP.SBOApplication.ActivateMenuItem("1289");
                                SAP.SBOApplication.ActivateMenuItem("1288");
                            }
                        }
                        break;
                }
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
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            saveRows(oForm);
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processMenuEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                if (pVal.MenuUID == "1282" || pVal.MenuUID == "1287")
                    BubbleEvent = false;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processMenuEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal)
        {
            try
            {
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processRightClickEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processRightClickEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal)
        {
            try
            {
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        private static bool validateRows(SAPbouiCOM.Form oForm)
        {
            return true;
        }
        public static void saveRows(SAPbouiCOM.Form oForm)
        {
            string dscolname = "";
            string columnName = "";
            try
            {
                //if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                FT_ADDON.SAP.SBOCompany.StartTransaction();
                SAPbouiCOM.BoFieldsType datatype = SAPbouiCOM.BoFieldsType.ft_AlphaNumeric;
                string docentry = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                string linenum = oForm.DataSources.UserDataSources.Item("linenum").Value.ToString();
                string dsname = oForm.DataSources.UserDataSources.Item("ds").Value.ToString();
                //string docentry = ((SAPbouiCOM.EditText)oForm.Items.Item("DOCENTRY").Specific).String;
                //string linenum = ((SAPbouiCOM.EditText)oForm.Items.Item("LINENUM").Specific).String;
                //string dsname = ((SAPbouiCOM.EditText)oForm.Items.Item("DS").Specific).String;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "";
                oMatrix.FlushToDataSource();
                
                SAPbobsCOM.SBObob dt = (SAPbobsCOM.SBObob)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                DateTime date;
                string temp = "";
                string value = "";
                decimal dtemp = 0;
                for (int x = 1; x <= oMatrix.RowCount; x++)
                {
                    linenum = "";
                    sql = "";
                    for (int col = 0; col < oMatrix.Columns.Count; col++)
                    {
                        columnName = oMatrix.Columns.Item(col).UniqueID;
                        switch (columnName)
                        {
                            case "LineNum":
                                linenum = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).String;
                                break;
                            default:
                                if (columnName.Contains("U_"))
                                {
                                    dscolname = oMatrix.Columns.Item(col).DataBind.Alias;
                                    datatype = (SAPbouiCOM.BoFieldsType)oForm.DataSources.DBDataSources.Item(dsname).Fields.Item(dscolname).Type;
                                    switch (datatype)
                                    {
                                        case SAPbouiCOM.BoFieldsType.ft_AlphaNumeric:
                                        case SAPbouiCOM.BoFieldsType.ft_Text:
                                            if (oMatrix.Columns.Item(columnName).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                                value = ((SAPbouiCOM.ComboBox)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).Selected.Value.ToString();
                                            else
                                                value = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).String;

                                            value = value.Replace("'", "''");
                                            temp = dscolname + " = '" + value + "'";
                                            if (sql != "")
                                                sql += " , " + temp;
                                            else
                                                sql = temp;

                                            break;
                                        case SAPbouiCOM.BoFieldsType.ft_Date:
                                            value = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).String;
                                            if (value == "")
                                            {
                                                temp = dscolname + " = null";
                                            }
                                            else
                                            {
                                                rc = dt.Format_StringToDate(value);
                                                rc.MoveFirst();
                                                date = DateTime.Parse(rc.Fields.Item(0).Value.ToString());

                                                temp = dscolname + " = '" + date.Year.ToString() + "-" + date.Month.ToString() + "-" + date.Day.ToString() + "'";
                                                if (sql != "")
                                                    sql += " , " + temp;
                                                else
                                                    sql = temp;
                                            }
                                            break;
                                        default:
                                            if (oMatrix.Columns.Item(columnName).Type == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                                                value = ((SAPbouiCOM.ComboBox)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).Selected.Value.ToString();
                                            else
                                                value = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).String;
                                            if (value == "")
                                            {
                                                temp = dscolname + " = 0";
                                            }
                                            else
                                            {
                                                dtemp = decimal.Parse(value);
                                                temp = dscolname + " = " + dtemp.ToString();
                                                if (sql != "")
                                                    sql += " , " + temp;
                                                else
                                                    sql = temp;
                                            }
                                            break;
                                    }
                                }
                                break;
                        }
                    }
                    sql = "UPDATE " + dsname + " SET " + sql + " WHERE DOCENTRY = " + docentry.ToString() + " AND LINENUM = " + linenum.ToString();
                    rs.DoQuery(sql);
                }
                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox(columnName + " " + dscolname + " " + ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
