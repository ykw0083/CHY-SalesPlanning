using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON
{
    class CFLAfter
    {
        public static void buttoncflcustom(SAPbouiCOM.Form oForm, string ds, string matrixname, SAPbouiCOM.Form cflForm)
        {
            string col = "";
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)cflForm.Items.Item("grid").Specific;
            int rowsadd = 0;
            int size = -1;
            DateTime dtvalue;
            bool isok = false;

            if (oGrid.Rows.SelectedRows.Count > 0)
            {
                oForm.DataSources.UserDataSources.Item("cfluid").Value = "";

                if (matrixname != "")
                {
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixname).Specific;
                    oMatrix.FlushToDataSource();

                    for (int x = 0; x < oGrid.Rows.Count; x++)
                    {
                        if (oGrid.Rows.IsSelected(x))
                        {
                            rowsadd++;

                            size = oForm.DataSources.DBDataSources.Item(ds).Size;
                            if (size == 1)
                            {
                                if (oForm.DataSources.DBDataSources.Item(ds).GetValue("U_SOITEMCO", size - 1).Trim() != "")
                                    oForm.DataSources.DBDataSources.Item(ds).InsertRecord(oForm.DataSources.DBDataSources.Item(ds).Size);
                            }
                            else
                                oForm.DataSources.DBDataSources.Item(ds).InsertRecord(oForm.DataSources.DBDataSources.Item(ds).Size);

                            size = oForm.DataSources.DBDataSources.Item(ds).Size;

                            if (col == "")
                            {
                                string column = "";
                                string value = "";
                                for (int y = 0; y < cflForm.DataSources.DataTables.Item("cfl").Columns.Count; y++)
                                {
                                    column = cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Name;
                                    if (cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Type == SAPbouiCOM.BoFieldsType.ft_Date)
                                    {
                                        try
                                        {
                                            if (cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x) != null)
                                            {
                                                isok = false;
                                                isok = DateTime.TryParse(cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString(), out dtvalue);
                                                if (isok)
                                                    oForm.DataSources.DBDataSources.Item(ds).SetValue(column, size - 1, dtvalue.ToString("yyyyMMdd"));
                                            }
                                        }
                                        catch //(Exception ex)
                                        {
                                            //SAP.SBOApplication.MessageBox(column + " - Error 1", 1, "Ok", "", "");
                                        }
                                    }
                                    else
                                        try
                                        {
                                            value = cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString();
                                            oForm.DataSources.DBDataSources.Item(ds).SetValue(column, size - 1, value);
                                        }
                                        catch //(Exception ex)
                                        {
                                            //SAP.SBOApplication.MessageBox(column + " - Error 2", 1, "Ok", "", "");
                                        }
                                }
                            }
                            if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Single)
                                break;
                        }
                    }


                    oItem = oForm.Items.Item(matrixname);
                    oMatrix.LoadFromDataSource();
                    //oMatrix.Columns.Item(col).Cells.Item(row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                    CHY.UserForm_SalesPlanning.arrangematrix(oForm, oMatrix, ds);
                    CHY.UserForm_SalesPlanning.SetLTotal(oForm);
                    CHY.UserForm_SalesPlanning.SetRemarks(oForm);
                    CHY.UserForm_SalesPlanning.SPLANRemainingOpen(oForm, false);
                    if (rowsadd > 0 && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                }
                else
                {
                    for (int x = 0; x < oGrid.Rows.Count; x++)
                    {
                        if (oGrid.Rows.IsSelected(x))
                        {
                            if (col == "")
                            {
                                string column = "";
                                string value = "";
                                for (int y = 0; y < cflForm.DataSources.DataTables.Item("cfl").Columns.Count; y++)
                                {
                                    column = cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Name;
                                    if (cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Type == SAPbouiCOM.BoFieldsType.ft_Date)
                                    {
                                        try
                                        {
                                            if (cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x) != null)
                                            { 
                                                isok = false;
                                                isok = DateTime.TryParse(cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString(), out dtvalue);
                                                if (isok)
                                                    oForm.DataSources.DBDataSources.Item(ds).SetValue(column, 0, dtvalue.ToString("yyyyMMdd"));
                                            }
                                        }
                                        catch //(Exception ex)
                                        {
                                            //SAP.SBOApplication.MessageBox(column + " - Error 1", 1, "Ok", "", "");
                                        }
                                    }
                                    else
                                        try
                                        {
                                            value = cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString();
                                            oForm.DataSources.DBDataSources.Item(ds).SetValue(column, 0, value);
                                        }
                                        catch //(Exception ex)
                                        {
                                            //SAP.SBOApplication.MessageBox(column + " - Error 2", 1, "Ok", "", "");
                                        }
                                }
                            }
                            if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Single)
                                break;
                        }
                    }
                    CHY.UserForm_SalesPlanning.SetLTotal(oForm);
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                }
            }

        }
        public static void aftercflcustom(SAPbouiCOM.Form oForm, string ds, string col, int row, string matrixname, string rtnvalue, SAPbouiCOM.Form cflForm)
        {
            SAPbouiCOM.Item oItem = null;
            SAPbouiCOM.Matrix oMatrix = null;
            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)cflForm.Items.Item(matrixname).Specific;
            bool first = true;
            int rowsadd = 0;
            DateTime dtvalue;

            oForm.DataSources.UserDataSources.Item("cfluid").Value = "";

            if (matrixname != "")
            {
                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixname).Specific;
                oMatrix.FlushToDataSource();
                row = row - 1;

                for (int x = 0; x < oGrid.Rows.Count; x++)
                {
                    if (oGrid.Rows.IsSelected(x))
                    {
                        rowsadd++;

                        if (!first)
                        {
                            oForm.DataSources.DBDataSources.Item(ds).InsertRecord(oForm.DataSources.DBDataSources.Item(ds).Size - 1);
                            row++;
                        }
                        if (col == "U_itemcode" || col == "U_conno")
                        {
                            string column = "";

                            for (int y = 0; y < cflForm.DataSources.DataTables.Item("cfl").Columns.Count; y++)
                            {
                                column = cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Name;
                                if (col == "U_itemcode")
                                {
                                    if (column == "dscription")
                                    {
                                        oForm.DataSources.DBDataSources.Item(ds).SetValue("U_itemname", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                    }
                                    else if (column == "itemcode")
                                    {
                                        oForm.DataSources.DBDataSources.Item(ds).SetValue("U_itemcode", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                    }
                                    else
                                        try
                                        {
                                            oForm.DataSources.DBDataSources.Item(ds).SetValue(column, row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                        }
                                        catch (Exception ex)
                                        {
                                            SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                                        }
                                }
                                else
                                    if (column == "docentry")
                                    {
                                        oForm.DataSources.DBDataSources.Item(ds).SetValue("U_docentry", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                    }
                                    else if (column == "lineid")
                                    {
                                        oForm.DataSources.DBDataSources.Item(ds).SetValue("U_lineid", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                    }
                                    else
                                        if (cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Type == SAPbouiCOM.BoFieldsType.ft_Date)
                                        {
                                            try
                                            {
                                                dtvalue = DateTime.Parse(cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                                oForm.DataSources.DBDataSources.Item(ds).SetValue(column, row, dtvalue.ToString("yyyyMMdd"));
                                            }
                                            catch //(Exception ex)
                                            {
                                                //SAP.SBOApplication.MessageBox(column + " - " + ex.Message, 1, "Ok", "", "");
                                            }
                                        }
                                        else
                                            try
                                            {
                                                oForm.DataSources.DBDataSources.Item(ds).SetValue(column, row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                            }
                                            catch
                                            {
                                            }
                                //try
                                        //{
                                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue(column, row, cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString());
                                        //}
                                        //catch (Exception ex)
                                        //{
                                        //    SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                                        //}
                            }
                        }

                        //if (col == "U_itemcode")
                        //{
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_itemcode", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(0, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_itemname", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(1, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_size", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(2, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_jcclr", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(3, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_brand", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(4, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_perfcl", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(5, x).ToString());
                        //}
                        //else if (col == "U_conno")
                        //{
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue(col, row, rtnvalue);
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_conno", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(0, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_blno", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(1, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_vessel", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(2, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_sealno", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(3, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_consize", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(4, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_netw", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(5, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_grossw", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(6, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_measure", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(7, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_batchno", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(8, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_bmd", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(9, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_bed", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(10, x).ToString());
                        //    oForm.DataSources.DBDataSources.Item(ds).SetValue("U_batchrem", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(11, x).ToString());
                        //}
                        first = false;
                    }
                }


                oItem = oForm.Items.Item(matrixname);
                oMatrix.LoadFromDataSource();
                CHY.UserForm_ShipDoc.arrangematrix(oForm, oMatrix, ds);
                oMatrix.Columns.Item(col).Cells.Item(row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                if (rowsadd > 0 && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                //oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);                       
            }
            else
            {
                for (int x = 0; x < oGrid.Rows.Count; x++)
                {
                    if (oGrid.Rows.IsSelected(x))
                    {
                        if (col == "U_booking")
                        {
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_booking", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(0, x).ToString());
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_vessel", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(1, x).ToString());
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_loading", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(2, x).ToString());
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_discharg", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(3, x).ToString());
                            //oForm.DataSources.DBDataSources.Item(ds).SetValue("U_shipper", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(4, x).ToString());
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        if (col == "U_shipper")
                        {
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_shipper", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(3, x).ToString());
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_consigne", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(4, x).ToString());
                            oForm.DataSources.DBDataSources.Item(ds).SetValue("U_notify", row, cflForm.DataSources.DataTables.Item("cfl").GetValue(5, x).ToString());
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                    }
                }
            }

        }
        public static void aftercfl(string FormUID, string ds, string col, int row, string matrixname, string rtnvalue, SAPbouiCOM.Form cflForm)
        {
            bool overridescript = false;
            bool isok = false;
            DateTime dtvalue;

            SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);

            if (overridescript) return;

            if (col == "" && rtnvalue == "")
            {
                int size = oForm.DataSources.DBDataSources.Item(ds).Size;
                buttoncflcustom(oForm, ds, matrixname, cflForm);
                return;
            }

            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)cflForm.Items.Item("grid").Specific;

            oForm.DataSources.UserDataSources.Item("cfluid").Value = "";

            //if (matrixname != "")
            //{
            //    row = row - 1;
            //}
            bool change = false;
            switch (FormUID)
            {
                case "":
                    break;
                default:
                    if (ds == "")
                    {
                        if (rtnvalue != "")
                            oForm.DataSources.UserDataSources.Item(col).Value = rtnvalue;
                    }
                    else
                    {
                        if (rtnvalue != "")
                        {
                            oForm.DataSources.DBDataSources.Item(ds).SetValue(col, row, rtnvalue);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                        }
                        else
                        {
                            for (int x = 0; x < oGrid.Rows.Count; x++)
                            {
                                if (oGrid.Rows.IsSelected(x))
                                {
                                    if (col != "")
                                    {
                                        string column = "";
                                        string value = "";
                                        for (int y = 0; y < cflForm.DataSources.DataTables.Item("cfl").Columns.Count; y++)
                                        {
                                            column = cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Name;
                                            if (cflForm.DataSources.DataTables.Item("cfl").Columns.Item(y).Type == SAPbouiCOM.BoFieldsType.ft_Date)
                                            {
                                                try
                                                {
                                                    if (cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x) != null)
                                                    {
                                                        isok = false;
                                                        isok = DateTime.TryParse(cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString(), out dtvalue);
                                                        if (isok)
                                                        {
                                                            oForm.DataSources.DBDataSources.Item(ds).SetValue(column, row, dtvalue.ToString("yyyyMMdd"));
                                                            change = true;
                                                        }
                                                    }
                                                }
                                                catch //(Exception ex)
                                                {
                                                    SAP.SBOApplication.MessageBox(column + " - Error 1", 1, "Ok", "", "");
                                                }
                                            }
                                            else
                                                try
                                                {
                                                    value = cflForm.DataSources.DataTables.Item("cfl").GetValue(y, x).ToString();
                                                    oForm.DataSources.DBDataSources.Item(ds).SetValue(column, row, value);
                                                    change = true;
                                                }
                                                catch //(Exception ex)
                                                {
                                                    SAP.SBOApplication.MessageBox(column + " - Error 2", 1, "Ok", "", "");
                                                }
                                        }
                                    }
                                    if (oGrid.SelectionMode == SAPbouiCOM.BoMatrixSelect.ms_Single)
                                        break;
                                }
                            }
                            if (change)
                            {
                                SAPbouiCOM.Item oItem = oForm.Items.Item(matrixname);
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
                                oMatrix.LoadFromDataSource();
                                //oMatrix.Columns.Item(col).Cells.Item(row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                                oMatrix.Columns.Item(col).Cells.Item(row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                            }
                        }
                    }
                    //if (matrixname != "")
                    //{
                    //    oItem = oForm.Items.Item(matrixname);
                    //    oMatrix.LoadFromDataSource();
                    //    oMatrix.Columns.Item(col).Cells.Item(row + 1).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    //        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                    //    //oItem.Click(SAPbouiCOM.BoCellClickType.ct_Regular);                       
                    //}

                    break;
            }

            if (oForm.TypeEx == "FT_CHARGE" || oForm.TypeEx == "FT_TPPLAN")
                CHY.UserForm_SalesPlanning.SetLTotal(oForm);
            if (oForm.TypeEx == "FT_CHARGE" && ds == "@FT_CHARGE1" && col == "U_ITEMCODE")
            {
                CHY.UserForm_SalesPlanning.deletebatch(oForm, row);
            }
        }
    }
}
