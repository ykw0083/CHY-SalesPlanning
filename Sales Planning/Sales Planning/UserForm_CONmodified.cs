using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_CONmodified
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                //if (oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString() != "")
                //{
                //    string FUID = oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString();
                //    SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FUID);
                //    oSForm.Select();
                //    BubbleEvent = false;
                //}
                switch (pVal.EventType)
                {
                    //case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                    //    if (pVal.ItemUID == "grid1" && pVal.Row > 0 && pVal.ColUID == "U_CONNO")
                    //    {
                    //        CFLInit.usercfl(oForm.UniqueID, "@" + oForm.DataSources.UserDataSources.Item("ds").Value.ToString(), pVal.ColUID, pVal.Row, pVal.ItemUID, "itemcode", "select itemcode, itemname from oitm");
                    //        BubbleEvent = false;

                    //    }
                    //    break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            string mustcol = oForm.DataSources.UserDataSources.Item("mustcol").Value.ToString();
                            //string mustcol = ((SAPbouiCOM.EditText)oForm.Items.Item("MUSTCOL").Specific).String;
                            if (!validateRows(oForm, mustcol))
                            {
                                BubbleEvent = false;
                            }
                        }
                        //if (pVal.ItemUID == "1" || pVal.ItemUID == "2")
                        //{
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //    {
                        //        SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).String);
                        //        oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                        //    }
                        //}
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
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        FT_ADDON.SAP.currentMatrixRow = pVal.Row;
                        if (pVal.ItemUID == "grid1")
                        {
                            oForm.EnableMenu("1292", true);
                            oForm.EnableMenu("1293", true);
                            oForm.EnableMenu("1294", true);
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
                else if (pVal.MenuUID == "1292")
                {
                    addNewRow(oForm);
                    BubbleEvent = false;
                }
                else if (pVal.MenuUID == "1293")
                {
                    deleteRow(oForm, FT_ADDON.SAP.currentMatrixRow);
                    BubbleEvent = false;
                }
                else if (pVal.MenuUID == "1294")
                {
                    duplicateRow(oForm, FT_ADDON.SAP.currentMatrixRow);
                    BubbleEvent = false;
                }
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
        public static void retrieveRowOld(SAPbouiCOM.Form oForm, int docno, int lineno, string dsname)
        {
            try
            {
                FT_ADDON.SAP.SBOCompany.StartTransaction();
                SAPbouiCOM.DBDataSource oDS = null;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                SAPbobsCOM.SBObob dt = (SAPbobsCOM.SBObob)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                DateTime date;
                string temp = "";
                string columnName = "";
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "SELECT * FROM [@" + dsname + "] WHERE DocEntry = " + docno.ToString() + " AND U_LINENO = " + lineno.ToString() + " ORDER BY LineId";
                rs.DoQuery(sql);
                if (rs.RecordCount > 0)
                {
                    rs.MoveFirst();
                    for (int x = 1; x <= rs.RecordCount; x++)
                    {
                        //addNewRow(oForm);
                        oForm.DataSources.DBDataSources.Item("@" + dsname).InsertRecord(x - 1);
                        oForm.DataSources.DBDataSources.Item("@" + dsname).Offset = x - 1;
                        for (int col = 0; col < oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Count; col++)
                        {
                            oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
                            columnName = oDS.Fields.Item(col).Name;
                            switch (columnName)
                            {
                                case "Object":
                                case "LogInst":
                                    break;
                                case "VisOrder":
                                    oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue(columnName, x - 1, x.ToString());
                                    break;
                                default:
                                    datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oDS.Fields.Item(col).Type);
                                    switch (datatype)
                                    {
                                        case SAPbouiCOM.BoDataType.dt_DATE:
                                            date = DateTime.Parse(rs.Fields.Item(columnName).Value.ToString());
                                            rc = dt.Format_DateToString(date);
                                            rc.MoveFirst();
                                            temp = rc.Fields.Item(0).Value.ToString();
                                            oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue(columnName, x - 1, temp);
                                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).String = temp;
                                            break;
                                        default:
                                            oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue(columnName, x - 1, rs.Fields.Item(columnName).Value.ToString());
                                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnName).Cells.Item(x).Specific)).String = rs.Fields.Item(columnName).Value.ToString();
                                            break;
                                    }
                                    break;
                            }
                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("CODE").Cells.Item(x).Specific)).String = rs.Fields.Item("CODE").Value.ToString();
                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_ITEMNO").Cells.Item(x).Specific)).String = rs.Fields.Item("U_ITEMNO").Value.ToString();
                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_CONNO").Cells.Item(x).Specific)).String = rs.Fields.Item("U_CONNO").Value.ToString();
                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_REF").Cells.Item(x).Specific)).String = rs.Fields.Item("U_REF").Value.ToString();
                            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_CONDATE").Cells.Item(x).Specific)).String = temp;
                        }
                        rs.MoveNext();
                    }
                    oMatrix.LoadFromDataSource();
                    deleteRow(oForm, oMatrix.RowCount);
                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
                else
                {
                    addNewRow(oForm);
                }
                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void addNewRow(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            //string bookno = oForm.DataSources.UserDataSources.Item("bookno").ValueEx.Trim();
            
            oMatrix.FlushToDataSource();

            string dsname = oForm.DataSources.UserDataSources.Item("ds").Value.ToString();
            SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
            oDS.InsertRecord(oDS.Size);
            oDS.Offset = oDS.Size - 1;
            oDS.SetValue("LineId", oDS.Size - 1, "0");
            oDS.SetValue("VisOrder", oDS.Size - 1, oDS.Size.ToString());
            //if (bookno != "")
            //    oDS.SetValue("U_BookNo", oDS.Size - 1, bookno);

            oMatrix.LoadFromDataSource();
            //oMatrix.AddRow(1, oMatrix.RowCount);
            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.RowCount).Specific)).String = "0";
            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("VisOrder").Cells.Item(oMatrix.RowCount).Specific)).String = oMatrix.RowCount.ToString();

            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
        }

        public static void deleteRow(SAPbouiCOM.Form oForm, int row)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            oMatrix.DeleteRow(row);
            for (int x = 1; x <= oMatrix.RowCount; x++)
            {
                ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("VisOrder").Cells.Item(x).Specific)).String = x.ToString();
            }
            oMatrix.FlushToDataSource();
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
        }

        public static void duplicateRow(SAPbouiCOM.Form oForm, int row)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;

            oMatrix.FlushToDataSource();

            string dsname = oForm.DataSources.UserDataSources.Item("ds").Value.ToString();
            SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
            oDS.InsertRecord(oDS.Size);
            oDS.Offset = oDS.Size - 1;
            oDS.SetValue("LineId", oDS.Size - 1, "0");
            oDS.SetValue("VisOrder", oDS.Size - 1, oDS.Size.ToString());

            SAPbobsCOM.UserTable oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
            SAPbobsCOM.UserFields oUF = null;
            oUF = oUT.UserFields;

            for (int x = 0; x < oUF.Fields.Count; x++)
            {
                string columnname = "";
                columnname = oUF.Fields.Item(x).Name;

                oDS.SetValue(columnname, oDS.Size - 1, oDS.GetValue(columnname, row - 1));
            }

            //oDS.SetValue("U_LINENO", oDS.Size - 1, oDS.GetValue("U_LINENO", row - 1));
            //oDS.SetValue("U_BOOKNO", oDS.Size - 1, oDS.GetValue("U_BOOKNO",row-1));
            //oDS.SetValue("U_BLNO", oDS.Size - 1, oDS.GetValue("U_BLNO", row - 1));
            //oDS.SetValue("U_CONNO", oDS.Size - 1, oDS.GetValue("U_CONNO", row - 1));

            oMatrix.LoadFromDataSource();
            //oMatrix.AddRow(1, oMatrix.RowCount);
            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("LineId").Cells.Item(oMatrix.RowCount).Specific)).String = "0";
            //((SAPbouiCOM.EditText)(oMatrix.Columns.Item("VisOrder").Cells.Item(oMatrix.RowCount).Specific)).String = oMatrix.RowCount.ToString();
            //)
            oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
            oMatrix.Columns.Item(1).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
        }

        private static bool validateRows(SAPbouiCOM.Form oForm, string col)
        {
            if (col.Length > 0)
            {
                string[] cols = col.Split(',');
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                for (int x = 1; x <= oMatrix.RowCount; x++)
                {
                    foreach (string column in cols)
                    {
                        if (((SAPbouiCOM.EditText)(oMatrix.Columns.Item(column).Cells.Item(x).Specific)).String == "")
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(column).Title + " must filled in!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            oMatrix.Columns.Item(column).Cells.Item(x).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                            return false;
                        }
                    }
                }
            }
            return true;
        }
        public static void saveRows(SAPbouiCOM.Form oForm)
        {
            try
            {
                string columnName = "";
                string docno = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                string lineno = oForm.DataSources.UserDataSources.Item("linenum").Value.ToString();
                string dsname = oForm.DataSources.UserDataSources.Item("ds").Value.ToString();
                //string docno = ((SAPbouiCOM.EditText)oForm.Items.Item("DOCENTRY").Specific).String;
                //string lineno = ((SAPbouiCOM.EditText)oForm.Items.Item("LINENUM").Specific).String;
                //string dsname = ((SAPbouiCOM.EditText)oForm.Items.Item("DS").Specific).String;

                SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                string sql = "";
                int lineid = 0;
                int lineidnew = 0;
                //string conno = "";
                string temp = "";

                //string code = "";
                //int newcode = 0;
                //int len = 0;
                //rs.DoQuery("SELECT TOP 1 LEN(code) FROM [@" + dsname + "] ORDER BY LEN(code) DESC");
                //if (rs.RecordCount > 0)
                //{
                //    rs.MoveFirst();
                //    len = int.Parse(rs.Fields.Item(0).Value.ToString());
                //    rs.DoQuery("SELECT TOP 1 code FROM [@" + dsname + "] WHERE LEN(code) = " + len.ToString() + " ORDER BY code DESC");
                //    if (rs.RecordCount > 0)
                //    {
                //        rs.MoveFirst();
                //        newcode = int.Parse(rs.Fields.Item(0).Value.ToString());
                //    }
                //}
                //for (int x = 1; x <= oMatrix.RowCount; x++)
                //{
                //    code = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("CODE").Cells.Item(x).Specific)).String;
                //    if (code == "")
                //    {
                //        newcode++;
                //        code = newcode.ToString();
                //        ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("CODE").Cells.Item(x).Specific)).String = code;
                //    }
                //}
                oMatrix.FlushToDataSource();

                SAPbobsCOM.SBObob dt = (SAPbobsCOM.SBObob)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //DateTime date;

                SAPbobsCOM.Recordset rss = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rss.DoQuery("SELECT TOP 1 LineId FROM [@" + dsname + "] WHERE DocEntry = " + docno.ToString() + " ORDER BY LineId DESC");
                if (rss.RecordCount > 0)
                {
                    rss.MoveFirst();
                    lineidnew = int.Parse(rss.Fields.Item(0).Value.ToString());
                }
                else
                {
                    lineidnew = 0;
                }

                FT_ADDON.SAP.SBOCompany.StartTransaction();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("DELETE FROM [@" + dsname + "] WHERE DocEntry = " + docno.ToString() + " AND U_LINENO = " + lineno.ToString());
                int row = oDS.Size;
                int columns = oDS.Fields.Count;
                string sdate = "";
                decimal dtemp = 0;
                string value = "";
                for (int x = 0; x < row; x++)
                {
                    //code = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("CODE").Cells.Item(x).Specific)).String;
                    oForm.DataSources.DBDataSources.Item("@" + dsname).Offset = x;
                    temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("LineId", x);
                    if (temp == "") temp = "0";
                    lineid = int.Parse(temp);
                    if (lineid <= 0)
                    {
                        //rs.DoQuery("SELECT TOP 1 LineId FROM [@" + dsname + "] WHERE DocEntry = " + docno.ToString() + " ORDER BY LineId DESC");
                        //if (rs.RecordCount > 0)
                        //{
                        //    rs.MoveFirst();
                        //    lineid = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                        //}
                        //else
                        //{
                        //    lineid = 1;
                        //}
                        lineidnew++;
                        lineid = lineidnew;
                    }

                    //conno = ((SAPbouiCOM.EditText)(oMatrix.Columns.Item("U_CONNO").Cells.Item(x).Specific)).String;
                    sql = "INSERT INTO [@" + dsname + "] (DocEntry,LineId,U_LINENO) VALUES (" + docno.ToString() + "," + lineid.ToString() + "," + lineno.ToString() + ")";
                    rs.DoQuery(sql);
                    sql = "";
                    for (int col = 0; col < columns; col++)
                    {
                        columnName = oDS.Fields.Item(col).Name;
                        switch (columnName)
                        {
                            case "DocEntry":
                            case "VisOrder":
                            case "Object":
                            case "LogInst":
                            case "LineId":
                            case "U_LINENO":
                                break;
                            default:
                                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oDS.Fields.Item(col).Type);
                                switch (datatype)
                                {
                                    case SAPbouiCOM.BoDataType.dt_LONG_TEXT:
                                    case SAPbouiCOM.BoDataType.dt_SHORT_TEXT:
                                        value = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(columnName, x).Trim();
                                        value = value.Replace("'", "''");
                                        temp = columnName + " = '" + value + "'";
                                        if (sql != "")
                                            sql += " , " + temp;
                                        else
                                            sql = temp;

                                        break;
                                    case SAPbouiCOM.BoDataType.dt_DATE:
                                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(columnName, x) == "")
                                        {
                                            temp = columnName + " = null"; 
                                        }
                                        else
                                        {
                                            //rc = dt.Format_StringToDate(oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(columnName, x));
                                            //rc.MoveFirst();
                                            //date = DateTime.Parse(rc.Fields.Item(0).Value.ToString());
                                            //temp = columnName + " = '" + date.Year.ToString() + "-" + date.Month.ToString() + "-" + date.Day.ToString() + "'";
                                            sdate = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(columnName, x);
                                            temp = columnName + " = '" + sdate.Substring(0, 4) + "-" + sdate.Substring(4, 2) + "-" + sdate.Substring(6, 2) + "'";
                                            if (sql != "")
                                                sql += " , " + temp;
                                            else
                                                sql = temp;
                                        }
                                        break;

                                    default:
                                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(columnName, x) == "")
                                        {
                                            temp = columnName + " = 0";
                                        }
                                        else
                                        {
                                            dtemp = decimal.Parse(oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(columnName, x));
                                            temp = columnName + " = " + dtemp.ToString();
                                            if (sql != "")
                                                sql += " , " + temp;
                                            else
                                                sql = temp;
                                        }
                                        break;
                                }
                                break;
                        }
                    }
                    if (sql != "")
                    {
                        sql = "UPDATE [@" + dsname + "] SET " + sql + " WHERE DocEntry = " + docno.ToString() + " AND LineId = " + lineid.ToString() + " AND U_LINENO = " + lineno.ToString();
                        rs.DoQuery(sql);
                    }
                }
                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch (Exception ex)
            {
                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void retrieveRow(SAPbouiCOM.Form oForm, long docno, int lineno, string dsname)
        {
            try
            {
                SAPbouiCOM.Conditions oCons = (SAPbouiCOM.Conditions)SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;

                //string bookno = oForm.DataSources.UserDataSources.Item("bookno").ValueEx.Trim();

                oDS.Clear();
                SAPbouiCOM.Condition oCon = oCons.Add();
                oCon.BracketOpenNum = 1;
                oCon.Alias = "DocEntry";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = docno.ToString();
                oCon.BracketCloseNum = 1;
                oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCon = oCons.Add();
                oCon.BracketOpenNum = 1;
                oCon.Alias = "U_LINENO";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = lineno.ToString();
                oCon.BracketCloseNum = 1;

                oDS.Query(oCons);

                if (oDS.Size > 0)
                {
                    for (int x = 1; x <= oDS.Size; x++)
                    {
                        oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("VisOrder", x - 1, x.ToString());
                    }
                    oMatrix.LoadFromDataSource();
                }
                else
                {
                    oDS.InsertRecord(oDS.Size);
                    oDS.Offset = oDS.Size - 1;
                    oDS.SetValue("LineId", oDS.Size - 1, "0");
                    oDS.SetValue("VisOrder", oDS.Size - 1, oDS.Size.ToString());
                    //if (bookno != "")
                    //    oDS.SetValue("U_BookNo", oDS.Size - 1, bookno);
                    oMatrix.LoadFromDataSource();
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
