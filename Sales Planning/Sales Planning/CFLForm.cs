using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON
{

    class CFLForm
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            //try
            //{
                string oFormId = "";
                SAPbouiCOM.Form oSForm = null;
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        oFormId = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
                        try
                        {
                            oSForm = SAP.SBOApplication.Forms.Item(oFormId);
                        }
                        catch
                        { }

                        if (oSForm != null)
                            oSForm.DataSources.UserDataSources.Item("cfluid").Value = "";
                            //oSForm.Select();

                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "2")
                        {
                            oFormId = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
                            try
                            {
                                oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(oFormId);
                            }
                            catch
                            { }

                            if (oSForm != null)
                                oSForm.DataSources.UserDataSources.Item("cfluid").Value = "";
                            //oSForm.Select();

                        }
                        else if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;

                            if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Auto)
                            {
                                if (oGrid.Rows.IsSelected(pVal.Row))
                                    oGrid.Rows.SelectedRows.Remove(pVal.Row);
                                else
                                    oGrid.Rows.SelectedRows.Add(pVal.Row);
                                BubbleEvent = false;
                            }
                            else
                                oGrid.Rows.SelectedRows.Add(pVal.Row);
                            
                        }
                        break;
                }
            //}
            //catch (Exception ex)
            //{
            //    SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            //}
        }

        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                SAPbouiCOM.Grid oGrid = null;
                SAPbouiCOM.GridColumn oColumn = null;
                SAPbouiCOM.DataTable oDataTable = null;
                string rtncol = "";
                string value = "";
                SAPbouiCOM.Form oSForm = null;

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_KEY_DOWN:
                        //SAP.SBOApplication.MessageBox(pVal.CharPressed.ToString(), 1, "Ok", "", "");
                        if (pVal.ItemUID == "FIND" && pVal.CharPressed == 13) // Enter Keydown
                        {                            
                            string sql = oForm.DataSources.UserDataSources.Item("select").Value.ToString();
                            string col = oForm.DataSources.UserDataSources.Item("orderby").Value.ToString();
                            string chk = "";
                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
                            oDataTable = oGrid.DataTable;
                            //oForm.DataSources.UserDataSources.Item("rtnvalue").Value = "";
                            if (col != "")
                            {
                                string find = ((SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific).Value;
                                if (find.Trim() != "")
                                {
                                    for (int x = 0; x < oGrid.Rows.Count; x++ )
                                    {
                                        chk = oDataTable.GetValue(col, x).ToString().ToLower();
                                        if (chk.StartsWith(find.ToLower()))
                                        {
                                            oGrid.Rows.SelectedRows.Add(x);
                                            rtncol = oForm.DataSources.UserDataSources.Item("rtncol").Value.ToString();
                                            if (rtncol != "")
                                                oForm.DataSources.UserDataSources.Item("rtnvalue").Value = oForm.DataSources.DataTables.Item("cfl").GetValue(col, x).ToString();
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "ALL")
                        {
                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
                            for (int x = 0; x < oGrid.Rows.Count; x++)
                            {
                                oGrid.Rows.SelectedRows.Add(x);
                            }
                        }
                        if (pVal.ItemUID == "choose")
                        {
                            value = oForm.DataSources.UserDataSources.Item("rtnvalue").Value.ToString();
                            choose(oForm, value);
                            //SAPbouiCOM.Form targetForm = null;
                            //try
                            //{
                            //    targetForm = SAP.SBOApplication.Forms.Item(oForm.DataSources.UserDataSources.Item("FormUID").Value);
                            //}
                            //catch { }
                            //if (targetForm != null)
                            //{
                            //    targetForm.DataSources.UserDataSources.Item("cflFormUID").Value = "";
                            //    switch (targetForm.TypeEx)
                            //    {
                            //        case "142": // Purchase Order
                            //            //setPurchaseCFL(oCFLForm, targetForm);
                            //            break;
                            //        case "FT_TPPLAN":
                            //            //setCylinderLoadingCFL(oCFLForm, targetForm);
                            //            break;
                            //    }
                            //}
                        }
                        else if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            //rtncol = oForm.DataSources.UserDataSources.Item("rtncol").Value.ToString();
                            //if (rtncol != "")
                            //    oForm.DataSources.UserDataSources.Item("rtnvalue").Value = oForm.DataSources.DataTables.Item("cfl").GetValue(rtncol, pVal.Row).ToString();
                            //oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;

                            //if (pVal.ColUID != "RowsHeader")
                            //{
                            //    if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Auto)
                            //    {
                            //        if (oGrid.Rows.IsSelected(pVal.Row))
                            //            oGrid.Rows.SelectedRows.Remove(pVal.Row);
                            //        else
                            //            oGrid.Rows.SelectedRows.Add(pVal.Row);
                            //    }
                            //    else
                            //        oGrid.Rows.SelectedRows.Add(pVal.Row);
                            //}
                        }
                        else if (pVal.ItemUID == "grid")
                        {
                            if (oForm.DataSources.UserDataSources.Item("select").Value != "")
                            {
                                if (pVal.ColUID != "RowsHeader")
                                {
                                    oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
                                    oColumn = oGrid.Columns.Item(pVal.ColUID);
                                    oDataTable = oGrid.DataTable;

                                    if (oDataTable.Columns.Item(pVal.ColUID).Type == SAPbouiCOM.BoFieldsType.ft_AlphaNumeric || oDataTable.Columns.Item(pVal.ColUID).Type == SAPbouiCOM.BoFieldsType.ft_Text)
                                    {
                                        oForm.DataSources.UserDataSources.Item("find").Value = "";
                                        oForm.DataSources.UserDataSources.Item("orderby").Value = pVal.ColUID;
                                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st_1").Specific).Caption = oColumn.TitleObject.Caption;
                                        oForm.Items.Item("FIND").Enabled = true;
                                        oForm.DataSources.DataTables.Item("cfl").ExecuteQuery(oForm.DataSources.UserDataSources.Item("select").Value.ToString() + " order by " + pVal.ColUID);
                                    }
                                    else
                                    {
                                        oForm.DataSources.UserDataSources.Item("find").Value = "";
                                        oForm.DataSources.UserDataSources.Item("orderby").Value = "";
                                        ((SAPbouiCOM.StaticText)oForm.Items.Item("st_1").Specific).Caption = oColumn.TitleObject.Caption;
                                        oForm.Items.Item("FIND").Enabled = true;
                                        oForm.DataSources.DataTables.Item("cfl").ExecuteQuery(oForm.DataSources.UserDataSources.Item("select").Value.ToString() + " order by " + pVal.ColUID);
                                    }
                                    string temp = "";
                                    string sql = "";
                                    SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    for (int x = 0; x < oGrid.Columns.Count; x++)
                                    {
                                        temp = oGrid.Columns.Item(x).UniqueID;
                                        temp = temp.Replace("U_", "");
                                        sql = "select top 1 Descr from CUFD where AliasID = '" + temp + "'";
                                        rc.DoQuery(sql);
                                        if (rc.RecordCount > 0)
                                        {
                                            oGrid.Columns.Item(x).TitleObject.Caption = rc.Fields.Item(0).Value.ToString();
                                        }

                                    }
                                    oForm.DataSources.UserDataSources.Item("rtnvalue").Value = "";
                                    oGrid.Rows.SelectedRows.Clear();
                                    foreach (SAPbouiCOM.GridColumn column in oGrid.Columns)
                                    {
                                        column.Editable = false;
                                    }
                                }
                            }
                        }
                        else if (pVal.ItemUID == "2")
                        { }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
                            if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Single)
                                if (oGrid.Rows.Count > 0)
                                {
                                    rtncol = oForm.DataSources.UserDataSources.Item("rtncol").Value.ToString();
                                    if (rtncol != "")
                                        value = oForm.DataSources.DataTables.Item("cfl").GetValue(pVal.ColUID, pVal.Row).ToString();
                                    else
                                        value = oForm.DataSources.UserDataSources.Item("rtnvalue").Value.ToString();
                                    choose(oForm, value);
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

        private static void choose(SAPbouiCOM.Form oForm, string value)
        {
            string FormUID = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
            string ds = oForm.DataSources.UserDataSources.Item("ds").Value.ToString();
            string col = oForm.DataSources.UserDataSources.Item("col").Value.ToString();
            int row = int.Parse(oForm.DataSources.UserDataSources.Item("row").Value.ToString());
            string matrixname = oForm.DataSources.UserDataSources.Item("matrixname").Value.ToString();
            string rtncol = oForm.DataSources.UserDataSources.Item("rtncol").Value.ToString();

            CFLAfter.aftercfl(FormUID, ds, col, row, matrixname, value, oForm);

            oForm.Close();
        }

    }

    class CFLFormText
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
                            string DSNAME = oForm.DataSources.UserDataSources.Item("DSNAME").Value.ToString();
                            string col = oForm.DataSources.UserDataSources.Item("col").Value.ToString();
                            int row = int.Parse(oForm.DataSources.UserDataSources.Item("row").Value.ToString());
                            string text = ((SAPbouiCOM.EditText)oForm.Items.Item("TEXT").Specific).Value.ToString();
                            string matrixname = oForm.DataSources.UserDataSources.Item("matrix").ValueEx.Trim();

                            oFormId = oForm.DataSources.UserDataSources.Item("FUID").Value.ToString();

                            oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(oFormId);
                            oSForm.DataSources.DBDataSources.Item(DSNAME).SetValue(col, row, text);
                            if (matrixname != "")
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oSForm.Items.Item(matrixname).Specific;
                                oMatrix.LoadFromDataSource();
                                oMatrix.Columns.Item(col).Cells.Item(row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                            if (oSForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oSForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        if (pVal.ItemUID == "1" || pVal.ItemUID == "2")
                        {
                            oFormId = oForm.DataSources.UserDataSources.Item("FUID").Value.ToString();
                            oForm.Close();
                            oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(oFormId);
                            oSForm.DataSources.UserDataSources.Item("cfluid").Value = "";
                            oSForm.Select();

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