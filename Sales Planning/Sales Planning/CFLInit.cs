using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON
{
    class CFLInit
    {
        public static void usercfl(string FormUID, string ds, string col, int row, string matrixname, string rtncol, string sql, bool multi)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize CFL window...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_CFL";

                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = "Choose From List";
                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);
                oForm.Left = oSForm.Left;
                oForm.Width = 600;
                oForm.Top = oSForm.Top;
                oForm.Height = 500;

                oForm.DataSources.UserDataSources.Add("FormUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("ds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("col", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("matrixname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("rtncol", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("rtnvalue", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 0);

                oForm.DataSources.UserDataSources.Item("FormUID").Value = FormUID;
                oForm.DataSources.UserDataSources.Item("ds").Value = ds;
                oForm.DataSources.UserDataSources.Item("col").Value = col;
                oForm.DataSources.UserDataSources.Item("row").Value = row.ToString();
                oForm.DataSources.UserDataSources.Item("matrixname").Value = matrixname;
                oForm.DataSources.UserDataSources.Item("rtncol").Value = rtncol;
                oForm.DataSources.UserDataSources.Item("rtnvalue").Value = "";

                SAPbouiCOM.Button oButton = null;
                SAPbouiCOM.Grid oGrid = null;
                SAPbouiCOM.Item oItem = null;

                oItem = oForm.Items.Add("RTNVALUE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "rtnvalue");
                oItem.Left = oForm.Width - 140;
                oItem.Width = 100;
                oItem.Top = 5;
                oItem.Height = 15;
                oItem.Enabled = false;

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                oItem = oForm.Items.Add("choose", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Choose";

                oItem = oForm.Items.Add("grid", SAPbouiCOM.BoFormItemTypes.it_GRID);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = 25;
                oItem.Height = oForm.Height - 90;

                oGrid = (SAPbouiCOM.Grid)oItem.Specific;
                oForm.DataSources.DataTables.Add("cfl");

                if (col == "U_itemcode" && ds == "@FT_SHIP1")
                {
                    sql = "select rdr1.itemcode, rdr1.dscription, t1.name as size, rdr1.U_size, t2.name as jcclr, rdr1.U_jcclr, t3.name as brand, rdr1.U_brand, t4.name as perfcl, rdr1.U_perfcl from rdr1 left outer join [@0009] t1 on t1.code = rdr1.U_size left outer join [@0007] t2 on t2.code = rdr1.U_jcclr left outer join [@0004] t3 on t3.code = rdr1.U_brand left outer join [@0001] t4 on t4.code = rdr1.U_perfcl where " + sql;
                }
                else if (col == "U_conno" && ds == "@FT_SHIP2")
                {
                    sql = "select t1.U_conno, t0.itemcode as U_itemcode, t0.dscription as U_itemname, t4.name as size, t7.name as jcclr, t5.name as brand, t6.name as perfcl, t1.U_conqty, t1.U_conuom, t1.U_blno, t1.U_vessel, t1.U_sealno, t3.name as consize, t1.U_consize, t1.U_netw, t1.U_grossw, t1.U_measure, t1.U_batchno, t1.U_bmd, t1.U_bed, t1.U_batchrem, t1.docentry, t1.lineid, t4.code as U_size, t7.code as U_jcclr, t5.code as U_brand, t6.code as U_perfcl from [@FT_APSOC] t1 inner join rdr1 t0 on t1.DocEntry = t0.DocEntry and t1.U_LINENO = t0.LineNum left outer join [@0014] t3 on t1.U_consize = t3.code left outer join [@0009] t4 on t4.code = t0.U_size left outer join [@0004] t5 on t5.code = t0.U_brand left outer join [@0001] t6 on t6.code = t0.U_perfcl left outer join [@0007] t7 on t7.code = t0.U_jcclr left outer join [@FT_SHIP2] t2 on t1.docentry = t2.U_docentry and t1.lineid = t2.U_lineid where " + sql + " and t2.docentry is null";

                    //sql = "select t1.U_conno as conno, t0.itemcode as U_itemcode, t0.dscription as U_itemname, t0.U_desc, t4.name as qtyperfcl, t0.U_perfcl, t1.* from [@FT_APSOC] t1 inner join rdr1 t0 on t1.DocEntry = t0.DocEntry and t1.U_LINENO = t0.LineNum left outer join [@00011] t4 on t4.code = t0.U_perfcl left outer join [@FT_SHIP2] t2 on t1.docentry = t2.U_docentry and t1.lineid = t2.U_lineid where " + sql + " and t2.docentry is null";
                }

                oForm.DataSources.DataTables.Item("cfl").ExecuteQuery(sql);
                ////////////////////// pohling request
                oForm.DataSources.UserDataSources.Add("select", SAPbouiCOM.BoDataType.dt_LONG_TEXT, 0);
                oForm.DataSources.UserDataSources.Add("orderby", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                oForm.DataSources.UserDataSources.Add("find", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 100);
                oForm.DataSources.UserDataSources.Item("select").Value = "";
                oForm.DataSources.UserDataSources.Item("orderby").Value = "";

                if (col == "U_itemcode" || col == "U_conno" || col == "c_booking")
                {
                }
                else
                {
                    oForm.DataSources.UserDataSources.Item("select").Value = sql;
                }

                //oItem = oForm.Items.Item("RTNVALUE");
                //oItem.Visible = false;

                oItem = oForm.Items.Add("FIND", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "find");
                oItem.Left = 110;
                oItem.Width = 100;
                oItem.Top = 5;
                oItem.Height = 15;
                oItem.Enabled = false;

                oItem = oForm.Items.Add("st_find", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 5;
                oItem.Width = 100;
                oItem.Top = 5;
                oItem.Height = 15;
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Find";
                //oItem.LinkTo = "FIND";

                oItem = oForm.Items.Add("st_1", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                oItem.Left = 220;
                oItem.Width = 200;
                oItem.Top = 5;
                oItem.Height = 15;
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "";
                //oItem.LinkTo = "FIND";

                //////////////////////
                oGrid.DataTable = oForm.DataSources.DataTables.Item("cfl");

                SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string temp = "";
                foreach (SAPbouiCOM.GridColumn column in oGrid.Columns)
                {
                    column.Editable = false;
                    switch (column.UniqueID)
                    {
                        case "U_size":
                        case "U_jcclr":
                        case "U_brand":
                        case "U_perfcl":
                        case "U_consize":
                            column.Visible = false;
                            break;
                    }
                    temp = column.UniqueID;
                    temp = temp.Replace("U_", "");
                    sql = "select top 1 Descr from CUFD where AliasID = '" + temp + "'";
                    rc.DoQuery(sql);
                    if (rc.RecordCount > 0)
                    {
                        column.TitleObject.Caption = rc.Fields.Item(0).Value.ToString();
                    }
                }
                if (multi)
                {
                    oItem = oForm.Items.Add("ALL", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    ((SAPbouiCOM.Button)oItem.Specific).Caption = "Choose All";
                    oItem.Left = oForm.Width - 95;
                    oItem.Width = 65;
                    oItem.Top = oForm.Height - 60;
                    oItem.Height = 20;

                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                }
                else
                    oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

                oSForm.DataSources.UserDataSources.Item("cfluid").Value = oForm.UniqueID;

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
                oForm.Visible = true;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }

        }
        public static void CFLText(string FormUID, long docentry, int row, string dsname, string col, string matrixname, string value)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_None);

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_CFLTEXT";
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                //creationPackage.FormType 
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = "Text Detail...";
                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);
                oForm.Left = oSForm.Left;
                oForm.Width = 500;
                oForm.Top = oSForm.Top;
                oForm.Height = 400;

                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();

                //oItem = oForm.Items.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).String = docentry.ToString();

                //oItem.Left = 140;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;
                //oItem.Enabled = true;
                //oItem.Visible = false;

                //oItem = oForm.Items.Add("LineNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).String = linenum.ToString();

                //oItem.Left = 150;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;
                //oItem.Enabled = true;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("FUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "fuid");
                //oItem.Left = 160;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = row.ToString();

                uds = oForm.DataSources.UserDataSources.Add("DSNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = dsname;

                uds = oForm.DataSources.UserDataSources.Add("col", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = col;

                uds = oForm.DataSources.UserDataSources.Add("matrix", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = matrixname;

                //oItem = oForm.Items.Add("DSNAME", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).String = dsname;
                //oItem.Left = 160;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;

                uds = oForm.DataSources.UserDataSources.Add("text", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                uds.Value = value;

                oItem = oForm.Items.Add("TEXT", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);

                oItem.Left = 5;
                oItem.Width = oForm.Width - 30;
                oItem.Top = 5;
                oItem.Height = oForm.Height - 90;
                oItem.Enabled = true;
                oItem.Visible = true;
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "text");

                //oForm.Items.Item("FUID").Visible = false;
                //oForm.Items.Item("LineNo").Visible = false;
                //oForm.Items.Item("DocEntry").Visible = false;
                //oForm.Items.Item("DSNAME").Visible = false;

                //oForm.DataSources.DBDataSources.Add("INV1");
                oItem = oForm.Items.Item("1");
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                oForm.Visible = true;
                //oForm.Modal = true;
                //oForm
                //oSForm.State = SAPbouiCOM.BoFormStateEnum.;
                //oSForm.Freeze(true);

                //((SAPbouiCOM.EditText)oSForm.Items.Item("FUID").Specific).Value = oForm.UniqueID.ToString();
                oSForm.DataSources.UserDataSources.Item("cfluid").Value = oForm.UniqueID;

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
