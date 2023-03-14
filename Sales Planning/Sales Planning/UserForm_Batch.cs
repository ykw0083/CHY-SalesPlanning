using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_Batch
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        string sform1 = oForm.DataSources.UserDataSources.Item("sform").ValueEx.Trim();
                        SAPbouiCOM.Form oSForm1 = FT_ADDON.SAP.SBOApplication.Forms.Item(sform1);

                        oSForm1.Freeze(false);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            string sform = oForm.DataSources.UserDataSources.Item("sform").ValueEx.Trim();
                            SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(sform);

                            oSForm.Freeze(false);
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "grid3" && pVal.Row > 0)
                        {

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                            if (oMatrix.IsRowSelected(pVal.Row))
                            { }
                            else
                            {
                                oMatrix.SelectRow(pVal.Row, true, false);
                                retrieveBatch(oForm);
                            }

                            BubbleEvent = false;

                        }
                        else if (pVal.ItemUID == "grid2" && pVal.Row > 0)
                        {
                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;

                            if (pVal.ColUID != "V_-1")
                            {
                                oMatrix.SelectRow(pVal.Row, true, false);
                                BubbleEvent = false;
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        //if (pVal.ItemUID == "grid3" && pVal.Row >= 0)
                        //{
                        //    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                        //    if (oMatrix.IsRowSelected(pVal.Row))
                        //    { }
                        //    else
                        //    {
                        //        oMatrix.SelectRow(pVal.Row, true, false);
                        //        retrieveBatch(oForm);
                        //    }
                        //}
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:

                        if (pVal.ItemUID == "cb_add" || pVal.ItemUID == "cb_delete")
                        {

                            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                            string visorder = "";
                            string itemcode = "";
                            string whscode = "";
                            string batchnum = "";
                            string bin = "";
                            int binabs = 0;
                            int row = oMatrix.GetNextSelectedRow(0);
                            if (row > 0)
                            {
                                visorder = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("V_-1").Cells.Item(row).Specific).Value.Trim();
                                itemcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ITEMCODE").Cells.Item(row).Specific).Value.Trim();
                                whscode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(row).Specific).Value.Trim();

                                string u_onhand = "";
                                string u_quantity = "";
                                bool found = false;
                                double qty = 0;
                                double onhand = 0;
                                string sform = oForm.DataSources.UserDataSources.Item("sform").ValueEx.Trim();
                                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(sform);
                                string sbatchds = oForm.DataSources.UserDataSources.Item("sbatchds").ValueEx.Trim();
                                SAPbouiCOM.DBDataSource oSDS = oSForm.DataSources.DBDataSources.Item(sbatchds);
                                string sdetailds = oForm.DataSources.UserDataSources.Item("sdetailds").ValueEx.Trim();

                                SAPbouiCOM.DBDataSource oDS1 = oForm.DataSources.DBDataSources.Item("@FT_BATCH");
                                SAPbouiCOM.DBDataSource oDS2 = oForm.DataSources.DBDataSources.Item(sbatchds);

                                SAPbouiCOM.Matrix oMatrix1 = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                SAPbouiCOM.Matrix oMatrix2 = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;

                                if (pVal.ItemUID == "cb_add")
                                {
                                    #region add batch

                                    for (int x = 1; x <= oMatrix1.RowCount; x++)
                                    {
                                        u_onhand = ((SAPbouiCOM.EditText)oMatrix1.Columns.Item("U_ONHAND").Cells.Item(x).Specific).Value.Trim();
                                        u_quantity = ((SAPbouiCOM.EditText)oMatrix1.Columns.Item("U_QUANTITY").Cells.Item(x).Specific).Value.Trim();
                                        if (!double.TryParse(u_onhand, out onhand))
                                        {
                                            onhand = 0;
                                        }
                                        if (!double.TryParse(u_quantity, out qty))
                                        {
                                            qty = 0;
                                        }

                                        if (qty > 0 && onhand >= qty)
                                        {
                                            batchnum = ((SAPbouiCOM.EditText)oMatrix1.Columns.Item("U_BATCHNUM").Cells.Item(x).Specific).Value.Trim();
                                            bin = ((SAPbouiCOM.EditText)oMatrix1.Columns.Item("U_BIN").Cells.Item(x).Specific).Value.Trim();
                                            binabs = int.Parse(((SAPbouiCOM.EditText)oMatrix1.Columns.Item("U_BINABS").Cells.Item(x).Specific).Value.Trim());

                                            found = false;
                                            for (int y = 0; y < oSDS.Size; y++)
                                            {
                                                if (visorder == oSDS.GetValue("U_BASEVIS", y).Trim())
                                                {
                                                    if (itemcode == oSDS.GetValue("U_ITEMCODE", y).Trim())
                                                    {
                                                        if (whscode == oSDS.GetValue("U_WHSCODE", y).Trim())
                                                        {
                                                            if (batchnum == oSDS.GetValue("U_BATCHNUM", y).Trim())
                                                            {
                                                                if (bin == oSDS.GetValue("U_BIN", y).Trim())
                                                                {
                                                                    found = true;
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            if (!found)
                                            {
                                                if (oDS2.Size == 0)
                                                    oDS2.InsertRecord(0);
                                                else
                                                {
                                                    if (oDS2.GetValue("U_BATCHNUM", oDS2.Size - 1).Trim() == "" && oDS2.GetValue("U_BIN", oDS2.Size - 1).Trim() == "")
                                                    { }
                                                    else
                                                        oDS2.InsertRecord(oDS2.Size);
                                                }
                                                oDS2.SetValue("U_BASEVIS", oDS2.Size - 1, visorder);
                                                oDS2.SetValue("U_ITEMCODE", oDS2.Size - 1, itemcode);
                                                oDS2.SetValue("U_WHSCODE", oDS2.Size - 1, whscode);
                                                oDS2.SetValue("U_BIN", oDS2.Size - 1, bin);
                                                oDS2.SetValue("U_BINABS", oDS2.Size - 1, binabs.ToString());
                                                oDS2.SetValue("U_BATCHNUM", oDS2.Size - 1, batchnum);
                                                oDS2.SetValue("U_QUANTITY", oDS2.Size - 1, qty.ToString());

                                                if (oSDS.Size == 0)
                                                    oSDS.InsertRecord(0);
                                                else
                                                {
                                                    if (oSDS.GetValue("U_BATCHNUM", oSDS.Size - 1).Trim() == "" && oSDS.GetValue("U_BIN", oSDS.Size - 1).Trim() == "")
                                                    { }
                                                    else
                                                        oSDS.InsertRecord(oSDS.Size);
                                                }

                                                oSDS.SetValue("U_BASEVIS", oSDS.Size - 1, visorder);
                                                oSDS.SetValue("U_ITEMCODE", oSDS.Size - 1, itemcode);
                                                oSDS.SetValue("U_WHSCODE", oSDS.Size - 1, whscode);
                                                oSDS.SetValue("U_BIN", oSDS.Size - 1, bin);
                                                oSDS.SetValue("U_BINABS", oSDS.Size - 1, binabs.ToString());
                                                oSDS.SetValue("U_BATCHNUM", oSDS.Size - 1, batchnum);
                                                oSDS.SetValue("U_QUANTITY", oSDS.Size - 1, qty.ToString());

                                                UserForm_APSA.arrangevisorder(oSForm, sbatchds);

                                                oMatrix2.LoadFromDataSource();
                                                ((SAPbouiCOM.EditText)oMatrix1.Columns.Item("U_QUANTITY").Cells.Item(x).Specific).Value = "0";
                                            }
                                            else
                                            {
                                                SAP.SBOApplication.StatusBar.SetText("Cannot proceed. Batch Number found in source.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                                return;
                                            }

                                            
                                        }

                                    }
                                    #endregion
                                }
                                else
                                {
                                    #region delete batch
                                    row = oMatrix2.GetNextSelectedRow(0);
                                    if (row <= 0)
                                    {
                                        SAP.SBOApplication.StatusBar.SetText("Cannot Delete. Nothing to removed.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return;
                                    }
                                    else
                                    {
                                        visorder = ((SAPbouiCOM.EditText)oMatrix2.Columns.Item("V_-1").Cells.Item(row).Specific).Value.Trim();
                                        itemcode = ((SAPbouiCOM.EditText)oMatrix2.Columns.Item("U_ITEMCODE").Cells.Item(row).Specific).Value.Trim();
                                        whscode = ((SAPbouiCOM.EditText)oMatrix2.Columns.Item("U_WHSCODE").Cells.Item(row).Specific).Value.Trim();
                                        batchnum = ((SAPbouiCOM.EditText)oMatrix2.Columns.Item("U_BATCHNUM").Cells.Item(row).Specific).Value.Trim();
                                        bin = ((SAPbouiCOM.EditText)oMatrix2.Columns.Item("U_BIN").Cells.Item(row).Specific).Value.Trim();
                                        binabs = int.Parse(((SAPbouiCOM.EditText)oMatrix2.Columns.Item("U_BINABS").Cells.Item(row).Specific).Value.Trim());

                                        bool delete = false;
                                        for (int x = 0; x < oSDS.Size; x++)
                                        {
                                            if (delete) continue;
                                            if (oSDS.GetValue("U_BASEVIS", x).Trim() == visorder)
                                            {
                                                if (oSDS.GetValue("U_ITEMCODE", x).Trim() == itemcode)
                                                {
                                                    if (oSDS.GetValue("U_WHSCODE", x).Trim() == whscode)
                                                    {
                                                        if (oSDS.GetValue("U_BATCHNUM", x).Trim() == batchnum)
                                                        {
                                                            if (oSDS.GetValue("U_BIN", x).Trim() == bin)
                                                            {
                                                                oSDS.RemoveRecord(x);
                                                                delete = true;
                                                                UserForm_APSA.arrangevisorder(oSForm, sbatchds);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        if (delete)
                                        {
                                            oMatrix2.DeleteRow(row);
                                            oMatrix2.FlushToDataSource();
                                        }
                                        else
                                        {
                                            SAP.SBOApplication.StatusBar.SetText("Cannot Delete. Batch No found in source.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                            return;
                                        }
                                    }
                                    #endregion
                                }

                                double bqty = 0;
                                for (int x = 0; x < oDS2.Size; x++)
                                {
                                    bqty += double.Parse(oDS2.GetValue("U_QUANTITY", x));
                                }

                                oDS1 = oForm.DataSources.DBDataSources.Item(sdetailds);

                                for (int x = 0; x < oDS1.Size; x++)
                                {
                                    if (oDS1.GetValue("VisOrder", x) == visorder)
                                    {
                                        oDS1.SetValue("U_BQTY", x, bqty.ToString());
                                        oMatrix.LoadFromDataSource();

                                        oMatrix.SelectRow(x + 1, true, false);
                                        oSForm.DataSources.DBDataSources.Item(sdetailds).SetValue("U_BQTY", x, bqty.ToString());
                                        ((SAPbouiCOM.Matrix)oSForm.Items.Item("grid1").Specific).LoadFromDataSource();
                                    }
                                }

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

        public static void processMenuEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {
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
            string temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("Status", 0).ToString();
            if (temp != "O")
            {
                SAP.SBOApplication.StatusBar.SetText("Cannot proceed, Document is not OPEN.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_SODOCNUM", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("SO No is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            if (int.Parse(temp) <= 0)
            {
                SAP.SBOApplication.StatusBar.SetText("SO No is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("Booking No is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_DOCDATE", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("SA Date is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }
            temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_DODATE", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("DO Date is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            SAPbouiCOM.DBDataSource ods = null;
            SAPbouiCOM.Matrix oMatrix = null;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FT_APSA1");
            oMatrix.FlushToDataSource();
            int cnt = 0;
            double qty = 0;
            double openqty = 0;
            bool found = false;
            for (int x = 0; x < ods.Size; x++)
            {
                cnt = x + 1;
                if (ods.GetValue("U_QUANTITY", x) == null)
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for # " + cnt.ToString() + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (ods.GetValue("U_OPENQTY", x) == null)
                {
                    SAP.SBOApplication.StatusBar.SetText("Open Quantity for # " + cnt.ToString() + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(ods.GetValue("U_QUANTITY", x), out qty))
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for # " + cnt.ToString() + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(ods.GetValue("U_OPENQTY", x), out openqty))
                {
                    SAP.SBOApplication.StatusBar.SetText("Open Quantity for # " + cnt.ToString() + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                if (openqty == 0 && qty == 0)
                { }
                else if (openqty > 0 && qty > 0)
                { found = true; }
                else
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for # " + cnt.ToString() + " cannot less than zero.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

            }

            if (!found)
            {
                SAP.SBOApplication.StatusBar.SetText("Cannot proceed, no item found.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            oMatrix.LoadFromDataSource();
            //for (int x = 0; x < ods.Size; x++)
            //{
            //    ods.RemoveRecord(x);
            //}
            //oMatrix.FlushToDataSource();

            return true;
        }

        public static void retrieveBatch(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                SAPbouiCOM.DBDataSource oDS = oForm.DataSources.DBDataSources.Item("@FT_BATCH");

                int dsrow = oDS.Size;
                string sql = "";
                for (int x = dsrow - 1; x >= 0; x--)
                {
                    oDS.RemoveRecord(oDS.Size - 1);
                }

                //oDS.InsertRecord(oDS.Size);

                int row = oMatrix.GetNextSelectedRow(0);

                if (row > 0)
                {
                    string sform = oForm.DataSources.UserDataSources.Item("sform").ValueEx.Trim();
                    SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(sform);
                    string sbatchds = oForm.DataSources.UserDataSources.Item("sbatchds").ValueEx.Trim();
                    SAPbouiCOM.DBDataSource oSDS = oSForm.DataSources.DBDataSources.Item(sbatchds);
                    double onhand = 0;
                    double temp = 0;

                    SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Item("ITEMBATCH");
                    string visorder = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("V_-1").Cells.Item(row).Specific).Value.Trim();
                    string itemcode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ITEMCODE").Cells.Item(row).Specific).Value.Trim();
                    string whscode = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_WHSCODE").Cells.Item(row).Specific).Value.Trim();

                    SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    rc.DoQuery("select BinActivat from OWHS where WhsCode = '" + whscode + "'");
                    if (rc.RecordCount > 0)
                    {
                        rc.MoveFirst();
                        string bin = rc.Fields.Item(0).Value.ToString();
                        if (bin == "Y")
                            sql = oForm.DataSources.UserDataSources.Item("sqlbatch2").ValueEx.Trim();
                        else
                            sql = oForm.DataSources.UserDataSources.Item("sqlbatch1").ValueEx.Trim();

                        int cnt = 0;
                        if (itemcode == "" || whscode == "")
                        { }
                        else
                        {
                            string query = "";
                            if (bin == "Y")
                                query = sql + " where ItemCode = '" + itemcode + "' and WhsCode = '" + whscode + "' order by BinCode, BatchNum";
                            else
                                query = sql + " where ItemCode = '" + itemcode + "' and WhsCode = '" + whscode + "' and OnHand > 0 order by BinCode, BatchNum";

                            oDT.ExecuteQuery(query);

                            if (oDT.Rows.Count > 0)
                            {
                                cnt = 0;
                                for (int x = 0; x < oDT.Rows.Count; x++)
                                {
                                    if (oDT.GetValue(0, x).ToString().Trim() == "") continue;

                                    cnt++;

                                    if (oDS.Size == 0)
                                        oDS.InsertRecord(0);
                                    else
                                        oDS.InsertRecord(oDS.Size);

                                    oDS.SetValue("Code", oDS.Size - 1, cnt.ToString());
                                    oDS.SetValue("U_ITEMCODE", oDS.Size - 1, oDT.GetValue(0, x).ToString().Trim());
                                    oDS.SetValue("U_WHSCODE", oDS.Size - 1, oDT.GetValue(1, x).ToString().Trim());
                                    oDS.SetValue("U_BIN", oDS.Size - 1, oDT.GetValue(2, x).ToString().Trim());
                                    oDS.SetValue("U_BINABS", oDS.Size - 1, oDT.GetValue(3, x).ToString().Trim());
                                    oDS.SetValue("U_BATCHNUM", oDS.Size - 1, oDT.GetValue(4, x).ToString().Trim());
                                    onhand = double.Parse(oDT.GetValue(5, x).ToString().Trim());
                                    for (int y = 0; y < oSDS.Size; y++)
                                    {
                                        if (oDS.GetValue("U_ITEMCODE", x) == oSDS.GetValue("U_ITEMCODE", y).Trim())
                                        {
                                            if (oDS.GetValue("U_WHSCODE", x) == oSDS.GetValue("U_WHSCODE", y).Trim())
                                            {
                                                if (oDS.GetValue("U_BIN", x) == oSDS.GetValue("U_BIN", y).Trim())
                                                {
                                                    if (oDS.GetValue("U_BATCHNUM", x) == oSDS.GetValue("U_BATCHNUM", y).Trim())
                                                    {
                                                        if (visorder != oSDS.GetValue("U_BASEVIS", y).Trim())
                                                        {
                                                            temp = double.Parse(oSDS.GetValue("U_QUANTITY", y).Trim());
                                                            onhand = onhand - temp;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    oDS.SetValue("U_ONHAND", oDS.Size - 1, onhand.ToString());
                                    oDS.SetValue("U_QUANTITY", oDS.Size - 1, "0");
                                }
                            }
                        }
                    }
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                    oMatrix.LoadFromDataSource();

                    oDS = oForm.DataSources.DBDataSources.Item(sbatchds);

                    dsrow = oDS.Size;
                    for (int x = dsrow - 1; x >= 0; x--)
                    {
                        oDS.RemoveRecord(oDS.Size - 1);
                    }

                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;


                    if (oSDS.Size == 1)
                    {
                        if (oSDS.GetValue("U_BATCHNUM", 0).Trim() == "" && oSDS.GetValue("U_BIN", 0).Trim() == "")
                        {
                            oMatrix.LoadFromDataSource();
                            return;
                        }
                    }

                    for (int x = 0; x < oSDS.Size; x++)
                    {
                        if (oSDS.GetValue("U_BATCHNUM", x).Trim() == "" && oSDS.GetValue("U_BIN", x).Trim() == "")
                        {
                        }
                        else
                        {
                            if (visorder == oSDS.GetValue("U_BASEVIS", x).Trim())
                            {
                                if (itemcode == oSDS.GetValue("U_ITEMCODE", x).Trim())
                                {
                                    if (whscode == oSDS.GetValue("U_WHSCODE", x).Trim())
                                    {
                                        if (oDS.Size == 0)
                                        {
                                            oDS.InsertRecord(0);
                                        }
                                        else
                                        {
                                            oDS.InsertRecord(oDS.Size);
                                        }

                                        oDS.SetValue("U_BASEVIS", oDS.Size - 1, oSDS.GetValue("U_BASEVIS", x).Trim());
                                        oDS.SetValue("U_ITEMCODE", oDS.Size - 1, oSDS.GetValue("U_ITEMCODE", x).Trim());
                                        oDS.SetValue("U_WHSCODE", oDS.Size - 1, oSDS.GetValue("U_WHSCODE", x).Trim());
                                        oDS.SetValue("U_BIN", oDS.Size - 1, oSDS.GetValue("U_BIN", x).Trim());
                                        oDS.SetValue("U_BINABS", oDS.Size - 1, oSDS.GetValue("U_BINABS", x).Trim());
                                        oDS.SetValue("U_BATCHNUM", oDS.Size - 1, oSDS.GetValue("U_BATCHNUM", x).Trim());
                                        oDS.SetValue("U_QUANTITY", oDS.Size - 1, oSDS.GetValue("U_QUANTITY", x).Trim());
                                    }
                                }
                            }

                        }
                    }

                    oMatrix.LoadFromDataSource();

                }

                return;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
