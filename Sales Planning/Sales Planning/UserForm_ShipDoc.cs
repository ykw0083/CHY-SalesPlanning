using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_ShipDoc
    {
        static int set = 0;
        //static long docnum = 0;

        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                string docentry = "";
                string blno = "";
                int cnt = 0;
                string value = "";
                SAPbouiCOM.Form oSForm = null;
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbobsCOM.Recordset rs = null;

                if (oForm.DataSources.UserDataSources.Count > 0)
                {
                    if (oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString() != "")
                    {
                        string FUID = oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString();
                        oSForm = SAP.SBOApplication.Forms.Item(FUID);
                        oSForm.Select();
                        BubbleEvent = false;
                    }
                }
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ColUID == "")
                        {
                            if (oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Fields.Item(pVal.ItemUID).Type == SAPbouiCOM.BoFieldsType.ft_Text)
                            {
                                docentry = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_sdoc", 0);
                                value = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue(pVal.ItemUID, 0).ToString();
                                if (value == null)
                                    value = "";
                                CFLInit.CFLText(oForm.UniqueID, long.Parse(docentry), 0, "@FT_SHIPD", pVal.ItemUID, "", value);
                                BubbleEvent = false;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_VALIDATE:
                        rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        if (pVal.ColUID == "U_conno")
                        {
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
                            value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).Value.ToString();

                            docentry = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_sdoc", 0);
                            rs.DoQuery("select count(*) from [@FT_APSOC] where docentry = " + docentry + " and U_conno = '" + value + "'");
                            if (rs.RecordCount > 0)
                            {
                                rs.MoveFirst();

                                cnt = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                                if (cnt == 0)
                                    BubbleEvent = false;
                                else
                                    addmatrix(oForm, oMatrix, pVal.ColUID, "@FT_SHIP2");
                            }
                            else
                                BubbleEvent = false;

                        }
                        else if (pVal.ColUID == "U_itemcode")
                        {
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                            value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_itemcode").Cells.Item(pVal.Row).Specific).Value.ToString();

                            docentry = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_sdoc", 0);
                            rs.DoQuery("select count(*) from rdr1 where docentry = " + docentry + " and itemcode = '" + value + "'");
                            if (rs.RecordCount > 0)
                            {
                                rs.MoveFirst();

                                cnt = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                                if (cnt == 0)
                                    BubbleEvent = false;
                                else
                                    addmatrix(oForm, oMatrix, pVal.ColUID, "@FT_SHIP1");
                            }
                            else
                                BubbleEvent = false;
                        }
                        else if (pVal.ColUID == "U_prodtype")
                        {
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                            addmatrix(oForm, oMatrix, pVal.ColUID, "@FT_SHIP3");
                        }

                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "c_itemdesc")
                        {
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_itemdesc", 0, "", "text", "select textcode, convert(nvarchar(1000),text) as text from OPDT", false);
                            BubbleEvent = false;
                        }
                        string address = "";
                        if (pVal.ItemUID == "c_shipper")
                        {
                            address = ", (case when U_ship1 <> '' then U_ship1 else '' end)";
                            address = address + " + (case when U_ship2 <> '' then CHAR(13) + CHAR(10) + U_ship2 else '' end)";
                            address = address + " + (case when U_ship3 <> '' then CHAR(13) + CHAR(10) + U_ship3 else '' end)";
                            address = address + " + (case when U_ship4 <> '' then CHAR(13) + CHAR(10) + U_ship4 else '' end)";
                            address = address + " + (case when U_ship5 <> '' then CHAR(13) + CHAR(10) + U_ship5 else '' end) as shipper";
                            address = address + ", (case when U_con1 <> '' then U_con1 else '' end)";
                            address = address + " + (case when U_con2 <> '' then CHAR(13) + CHAR(10) + U_con2 else '' end)";
                            address = address + " + (case when U_con3 <> '' then CHAR(13) + CHAR(10) + U_con3 else '' end)";
                            address = address + " + (case when U_con4 <> '' then CHAR(13) + CHAR(10) + U_con4 else '' end)";
                            address = address + " + (case when U_con5 <> '' then CHAR(13) + CHAR(10) + U_con5 else '' end) as consigne";
                            address = address + ", (case when U_notify1 <> '' then U_notify1 else '' end)";
                            address = address + " + (case when U_notify2 <> '' then CHAR(13) + CHAR(10) + U_notify2 else '' end)";
                            address = address + " + (case when U_notify3 <> '' then CHAR(13) + CHAR(10) + U_notify3 else '' end)";
                            address = address + " + (case when U_notify4 <> '' then CHAR(13) + CHAR(10) + U_notify4 else '' end)";
                            address = address + " + (case when U_notify5 <> '' then CHAR(13) + CHAR(10) + U_notify5 else '' end) as notify_party";
                            //CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_shipper", 0, "", "fulladd", "select code, name + CHAR(13) + CHAR(10) + U_add1 + CHAR(13) + CHAR(10) + U_add2 as fulladd from [@0012]", false);
                            //CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_shipper", 0, "", "fulladd", "select code, name, " + address + " as fulladd from [@0020]", false);
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_shipper", 0, "", "", "select code as buyer, U_bname as name, U_discharg as port_of_discharge " + address + " from [@FT_BUYER] where code like '" + ((SAPbouiCOM.EditText)oForm.Items.Item("U_shipper").Specific).Value.ToString() + "%'", false);
                            BubbleEvent = false;
                        }
                        address = "(case when U_add1 <> '' then U_add1 else '' end)";
                        address = address + " + (case when U_add2 <> '' then CHAR(13) + CHAR(10) + U_add2 else '' end)";
                        address = address + " + (case when U_add3 <> '' then CHAR(13) + CHAR(10) + U_add3 else '' end)";
                        address = address + " + (case when U_add4 <> '' then CHAR(13) + CHAR(10) + U_add4 else '' end)";
                        address = address + " + (case when U_add5 <> '' then CHAR(13) + CHAR(10) + U_add5 else '' end)";
                        if (pVal.ItemUID == "c_consigne")
                        {
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_consigne", 0, "", "fulladd", "select code, name, " + address + " as fulladd from [@0018]", false);
                            BubbleEvent = false;
                        }
                        if (pVal.ItemUID == "c_notify")
                        {
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_notify", 0, "", "fulladd", "select code, name, " + address + " as fulladd from [@0019]", false);
                            BubbleEvent = false;
                        }
                        if (pVal.ItemUID == "c_loading")
                        {
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_loading", 0, "", "fulladd", "select code, name as fulladd from [@0008]", false);
                            BubbleEvent = false;
                        }
                        if (pVal.ItemUID == "c_discharg")
                        {
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_discharg", 0, "", "fulladd", "select code, name as fulladd from [@0017]", false);
                            BubbleEvent = false;
                        }
                        if (pVal.ItemUID == "c_country")
                        {
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_country", pVal.Row, "", "name", "select code, name from [@0013]", false);
                            BubbleEvent = false;
                        }
                        if (pVal.ItemUID == "c_booking")
                        {
                            address = "(case when t4.U_add1 <> '' then t4.U_add1 else '' end)";
                            address = address + " + (case when t4.U_add2 <> '' then CHAR(13) + CHAR(10) + t4.U_add2 else '' end)";
                            address = address + " + (case when t4.U_add3 <> '' then CHAR(13) + CHAR(10) + t4.U_add3 else '' end)";
                            address = address + " + (case when t4.U_add4 <> '' then CHAR(13) + CHAR(10) + t4.U_add4 else '' end)";
                            address = address + " + (case when t4.U_add5 <> '' then CHAR(13) + CHAR(10) + t4.U_add5 else '' end)";
                            docentry = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_sdoc", 0);
                            //CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_booking", pVal.Row, "", "U_bookno", "select distinct t1.U_bookno, t1.U_notify, t1.U_vessel, t2.name, t3.name from [@FT_APSOC] t1 left outer join [@0008] t2 on t1.U_portload = t2.code left outer join [@0017] t3 on t1.U_dest = t3.code where docentry = " + docentry, false);
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIPD", "U_booking", pVal.Row, "", "U_bookno", "select distinct t1.U_bookno, t1.U_vessel, t2.name as port_of_loading, t3.name as port_of_discharge from [@FT_APSOC] t1 left outer join [@0008] t2 on t1.U_portload = t2.code left outer join [@0017] t3 on t1.U_dest = t3.code left outer join [@0020] t4 on t1.U_shipper = t4.code where docentry = " + docentry, false);
                            BubbleEvent = false;
                        }
                        if (pVal.ColUID == "U_conno")
                        {
                            docentry = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_sdoc", 0);
                            blno = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_booking", 0);
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
                            if (oMatrix.RowCount == pVal.Row)
                                CFLInit.usercfl(oForm.UniqueID, "@FT_SHIP2", pVal.ColUID, pVal.Row, pVal.ItemUID, "U_conno", "t1.docentry = " + docentry + " and t1.U_BOOKNO = '" + blno + "'", true);
                            else
                                CFLInit.usercfl(oForm.UniqueID, "@FT_SHIP2", pVal.ColUID, pVal.Row, pVal.ItemUID, "U_conno", "t1.docentry = " + docentry + " and t1.U_BOOKNO = '" + blno + "'", false);

                            BubbleEvent = false;

                        }
                        if (pVal.ColUID == "U_item1" || pVal.ColUID == "U_item2")
                        {
                            docentry = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                            CFLInit.usercfl(oForm.UniqueID, "@FT_SHIP2", pVal.ColUID, pVal.Row, pVal.ItemUID, "dscription", "select itemcode, dscription from rdr1 where docentry = " + docentry, false);
                            BubbleEvent = false;
                        }
                        else if (pVal.ColUID == "U_itemcode")
                        {
                            docentry = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                            if (oMatrix.RowCount == pVal.Row)
                                CFLInit.usercfl(oForm.UniqueID, "@FT_SHIP1", pVal.ColUID, pVal.Row, pVal.ItemUID, "itemcode", "rdr1.docentry = " + docentry, true);
                            else
                                CFLInit.usercfl(oForm.UniqueID, "@FT_SHIP1", pVal.ColUID, pVal.Row, pVal.ItemUID, "itemcode", "rdr1.docentry = " + docentry, false);
                            BubbleEvent = false;
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "fgrid1")
                            oForm.PaneLevel = 1;
                        else if (pVal.ItemUID == "fgrid2")
                            oForm.PaneLevel = 2;
                        else if (pVal.ItemUID == "fgrid3")
                            oForm.PaneLevel = 3;

                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (!validateRows(oForm))
                                {
                                    BubbleEvent = false;
                                }
                                else
                                {
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        set = int.Parse(oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_set", 0).ToString());
                                        //rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                        //docentry = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                                        //oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_sdoc", 0, docentry);

                                        //rs.DoQuery("select max(U_set) from [@FT_SHIPD] where U_sdoc = " + docentry.ToString());
                                        //if (rs.RecordCount > 0)
                                        //{
                                        //    rs.MoveFirst();
                                        //    set = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                                        //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_set", 0, set.ToString());
                                        //}
                                        //rs.DoQuery("select max(docentry) from [@FT_SHIPD]");
                                        //if (rs.RecordCount > 0)
                                        //{
                                        //    docnum = long.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                                        //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("docnum", 0, docnum.ToString());
                                        //}
                                    }
                                }
                            }
                        }
                        if (pVal.ItemUID == "1" || pVal.ItemUID == "2")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE || pVal.ItemUID == "2")
                            {
                                oSForm = SAP.SBOApplication.Forms.Item(oForm.DataSources.UserDataSources.Item("fuid").Value.ToString());
                                //SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).String);
                                
                                oSForm.Freeze(false);
                                oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                                oSForm.Select();
                            }
                        }
                        break;

                }
                rs = null;
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
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        SAPbouiCOM.Form oSForm = SAP.SBOApplication.Forms.Item(oForm.DataSources.UserDataSources.Item("fuid").Value.ToString());
                        //SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).String);

                        oSForm.Freeze(false);
                        oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                        oSForm.Select();
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                        oForm.EnableMenu("1287", true);//duplicate
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                        oForm.EnableMenu("1287", false);//duplicate
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {

                            //if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                set = set + 1;
                                //docnum = docnum + 1;

                                oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_set", 0, set.ToString());
                                //oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("docnum", 0, docnum.ToString());
                                oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_sdoc", 0, oForm.DataSources.UserDataSources.Item("docentry").Value.ToString());
                                oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_pino", 0, oForm.DataSources.UserDataSources.Item("docnum").Value.ToString());
                                
                                SAPbouiCOM.Matrix oMatrix = null;

                                oForm.DataSources.DBDataSources.Item("@FT_SHIP1").SetValue("VisOrder", 0, "1");
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                oMatrix.LoadFromDataSource();

                                oForm.DataSources.DBDataSources.Item("@FT_SHIP2").SetValue("VisOrder", 0, "1");
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
                                oMatrix.LoadFromDataSource();

                                oForm.DataSources.DBDataSources.Item("@FT_SHIP3").SetValue("VisOrder", 0, "1");
                                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                                oMatrix.LoadFromDataSource();
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        SAP.currentMatrixRow = pVal.Row;
                        if (pVal.ItemUID == "grid1" || pVal.ItemUID == "grid2" || pVal.ItemUID == "grid3")
                        {
                            oForm.EnableMenu("1292", true);//add row
                            oForm.EnableMenu("1293", true);//delete row
                            oForm.EnableMenu("1294", true);//duplicate row
                            oForm.EnableMenu("1283", false);//duplicate row
                            oForm.EnableMenu("1287", false);//duplicate row
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if (pVal.ItemUID == "grid1" || pVal.ItemUID == "grid2" || pVal.ItemUID == "grid3")
                        {
                            oForm.EnableMenu("1292", false);//add row
                            oForm.EnableMenu("1293", false);//delete row
                            oForm.EnableMenu("1294", false);//duplicate row
                            oForm.EnableMenu("1283", true);//duplicate row
                            oForm.EnableMenu("1287", true);//duplicate row
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
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                long docentry = 0;
                long U_sdoc = 0;
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCon = null;
                string sql = "";
                string value = "";

                if (pVal.MenuUID == "1281")
                    BubbleEvent = false;
                if (pVal.MenuUID == "1282" || pVal.MenuUID == "1287" || pVal.MenuUID == "1283")
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Save this Document first.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }

                    if (pVal.MenuUID == "1283")
                    {

                        if (FT_ADDON.SAP.SBOApplication.MessageBox("Are you sure you want to remove this Document?", 2, "Yes", "No") == 2)
                            BubbleEvent = false;
                    }
                }

                if (pVal.MenuUID == "1288" || pVal.MenuUID == "1289" || pVal.MenuUID == "1290" || pVal.MenuUID == "1291")
                {
                    BubbleEvent = false;

                    if (long.TryParse(oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("docentry", 0), out docentry) == false)
                        docentry = 0;
                    if (long.TryParse(oForm.DataSources.UserDataSources.Item("docentry").Value.ToString(), out U_sdoc) == false)
                        U_sdoc = 0;

                    if (U_sdoc > 0)
                    {
                        if (docentry > 0)
                        {
                            if (pVal.MenuUID == "1288")//next record
                                sql = "select top 1 docentry from [@FT_SHIPD] where docentry > " + docentry.ToString() + " and U_sdoc = " + U_sdoc.ToString() + " order by docentry";
                            if (pVal.MenuUID == "1289")//previous record
                                sql = "select top 1 docentry from [@FT_SHIPD] where docentry < " + docentry.ToString() + " and U_sdoc = " + U_sdoc.ToString() + " order by docentry desc";
                            if (pVal.MenuUID == "1290")//first record
                                sql = "select top 1 docentry from [@FT_SHIPD] where U_sdoc = " + U_sdoc.ToString() + " order by docentry";
                            if (pVal.MenuUID == "1291")//last record
                                sql = "select top 1 docentry from [@FT_SHIPD] where U_sdoc = " + U_sdoc.ToString() + " order by docentry desc";
                        }
                        else
                        {
                            if (pVal.MenuUID == "1288")//next record
                                sql = "select top 1 docentry from [@FT_SHIPD] where U_sdoc = " + U_sdoc.ToString() + " order by docentry";
                            if (pVal.MenuUID == "1289")//previous record
                                sql = "select top 1 docentry from [@FT_SHIPD] where U_sdoc = " + U_sdoc.ToString() + " order by docentry desc";
                            if (pVal.MenuUID == "1290")//first record
                                sql = "select top 1 docentry from [@FT_SHIPD] where U_sdoc = " + U_sdoc.ToString() + " order by docentry";
                            if (pVal.MenuUID == "1291")//last record
                                sql = "select top 1 docentry from [@FT_SHIPD] where U_sdoc = " + U_sdoc.ToString() + " order by docentry desc";
                        }
                        rs.DoQuery(sql);
                        if (rs.RecordCount > 0)
                        {
                            rs.MoveFirst();
                            docentry = long.Parse(rs.Fields.Item(0).Value.ToString());

                            oCon = oCons.Add();
                            oCon.BracketOpenNum = 1;
                            oCon.Alias = "docentry";
                            oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCon.CondVal = docentry.ToString();
                            oCon.BracketCloseNum = 1;
                            oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons);

                            oForm.DataSources.DBDataSources.Item("@FT_SHIP1").Query(oCons);
                            oForm.DataSources.DBDataSources.Item("@FT_SHIP2").Query(oCons);
                            oForm.DataSources.DBDataSources.Item("@FT_SHIP3").Query(oCons);
                            ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                            ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                            ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();
                            oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                        }
                    }
                }

                
                if (pVal.MenuUID == "1294")
                {
                    if (oForm.ActiveItem == "grid1")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_itemcode").Cells.Item(SAP.currentMatrixRow).Specific).Value.ToString();
                        if (value == "")
                            BubbleEvent = false;
                        value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_itemcode").Cells.Item(oMatrix.RowCount).Specific).Value.ToString();
                        if (value != "")
                        {
                            BubbleEvent = false;
                        }
                    }
                    else if (oForm.ActiveItem == "grid2")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_conno").Cells.Item(SAP.currentMatrixRow).Specific).Value.ToString();
                        if (value == "")
                            BubbleEvent = false;
                        value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_conno").Cells.Item(oMatrix.RowCount).Specific).Value.ToString();
                        if (value != "")
                        {
                            BubbleEvent = false;
                        }
                    }
                    else if (oForm.ActiveItem == "grid3")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ffa").Cells.Item(SAP.currentMatrixRow).Specific).Value.ToString();
                        if (value == "")
                            BubbleEvent = false;
                        value = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_ffa").Cells.Item(oMatrix.RowCount).Specific).Value.ToString();
                        if (value != "")
                        {
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.MenuUID == "1292")
                {
                    if (oForm.ActiveItem == "grid1")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        addmatrix(oForm, oMatrix, "U_itemcode", "@FT_SHIP1");
                    }
                    else if (oForm.ActiveItem == "grid2")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        addmatrix(oForm, oMatrix, "U_conno", "@FT_SHIP2");
                    }
                    else if (oForm.ActiveItem == "grid3")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        addmatrix(oForm, oMatrix, "U_ffa", "@FT_SHIP3");
                    }
                }
                else if (pVal.MenuUID == "1293")
                {
                    if (oForm.ActiveItem == "grid1")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        if (oMatrix.RowCount == 1)
                        {
                            oMatrix.DeleteRow(1);
                            addmatrix(oForm, oMatrix, "", "@FT_SHIP1");
                        }
                        else
                        {
                            oMatrix.DeleteRow(SAP.currentMatrixRow);
                            arrangematrix(oForm, oMatrix, "@FT_SHIP1");
                        }
                        BubbleEvent = false;
                    }
                    else if (oForm.ActiveItem == "grid2")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        if (oMatrix.RowCount == 1)
                        {
                            oMatrix.DeleteRow(1);
                            addmatrix(oForm, oMatrix, "", "@FT_SHIP2");
                        }
                        else
                        {
                            oMatrix.DeleteRow(SAP.currentMatrixRow);
                            arrangematrix(oForm, oMatrix, "@FT_SHIP2");
                        }
                        BubbleEvent = false;

                    }
                    else if (oForm.ActiveItem == "grid3")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(oForm.ActiveItem).Specific;
                        if (oMatrix.RowCount == 1)
                        {
                            oMatrix.DeleteRow(1);
                            addmatrix(oForm, oMatrix, "", "@FT_SHIP3");
                        }
                        else
                        {
                            oMatrix.DeleteRow(SAP.currentMatrixRow);
                            arrangematrix(oForm, oMatrix, "@FT_SHIP3");
                        }
                        BubbleEvent = false;
                    }
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                rs = null;
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
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string docentry = "";
                string docnum = "";

                if (pVal.MenuUID == "1283" || pVal.MenuUID == "1282" || pVal.MenuUID == "1287")
                {
                    if (pVal.MenuUID == "1283")
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                    }

                    docentry = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                    docnum = oForm.DataSources.UserDataSources.Item("docnum").Value.ToString();
                    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_sdoc", 0, docentry);
                    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_pino", 0, docnum);

                    rs.DoQuery("select max(U_set) from [@FT_SHIPD] where U_sdoc = " + docentry.ToString());
                    if (rs.RecordCount > 0)
                    {
                        rs.MoveFirst();
                        set = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_set", 0, set.ToString());
                    }
                    //rs.DoQuery("select max(docentry) from [@FT_SHIPD]");
                    //if (rs.RecordCount > 0)
                    //{
                    //    docnum = long.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                    //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("docnum", 0, docnum.ToString());
                    //}

                    SAPbouiCOM.Matrix oMatrix = null;

                    oForm.DataSources.DBDataSources.Item("@FT_SHIP1").SetValue("VisOrder", 0, "1");
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                    oMatrix.LoadFromDataSource();

                    oForm.DataSources.DBDataSources.Item("@FT_SHIP2").SetValue("VisOrder", 0, "1");
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
                    oMatrix.LoadFromDataSource();

                    oForm.DataSources.DBDataSources.Item("@FT_SHIP3").SetValue("VisOrder", 0, "1");
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                    oMatrix.LoadFromDataSource();
                }
                rs = null;
                
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
            string temp = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_pino",0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("PI No is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_docdate", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("PI Date is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_set", 0).ToString();
            if (temp == null || temp.Trim() == "" || temp == "0")
            {
                SAP.SBOApplication.StatusBar.SetText("Set is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("U_booking", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("Booking No is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            SAPbouiCOM.DBDataSource ods = null;
            SAPbouiCOM.Matrix oMatrix = null;
            int cnt = 0;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FT_SHIP1");
            oMatrix.FlushToDataSource();
            for (int x = 0; x < ods.Size; x++)
            {
                cnt = x + 1;
                ods.SetValue("VisOrder", x, cnt.ToString());
                ods.SetValue("LineId", x, cnt.ToString());
            }
            oMatrix.LoadFromDataSource();
            //for (int x = 0; x < ods.Size; x++)
            //{
            //    ods.RemoveRecord(x);
            //}
            //oMatrix.FlushToDataSource();

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FT_SHIP2");
            oMatrix.FlushToDataSource();
            for (int x = 0; x < ods.Size; x++)
            {
                cnt = x + 1;
                ods.SetValue("VisOrder", x, cnt.ToString());
                ods.SetValue("LineId", x, cnt.ToString());
            }
            oMatrix.LoadFromDataSource();
            //for (int x = 0; x < ods.Size; x++)
            //{
            //    ods.RemoveRecord(x);
            //}
            //oMatrix.FlushToDataSource();

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FT_SHIP3");
            oMatrix.FlushToDataSource();
            for (int x = 0; x < ods.Size; x++)
            {
                cnt = x + 1;
                ods.SetValue("VisOrder", x, cnt.ToString());
                ods.SetValue("LineId", x, cnt.ToString());
            }
            oMatrix.LoadFromDataSource();
            //for (int x = 0; x < ods.Size; x++)
            //{
            //    ods.RemoveRecord(x);
            //}
            //oMatrix.FlushToDataSource();

            return true;
        }
        private static void addmatrix(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix matrix, string col, string ds)
        {
            try
            {
                SAPbouiCOM.DBDataSource ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item(ds);
                matrix.FlushToDataSource();                
                if (col == "" && ods.Size == 1)
                {
                    ods.SetValue("VisOrder", ods.Size - 1, ods.Size.ToString());
                    //oForm.DataSources.DBDataSources.Item(ds).SetValue("VisOrder", ods.Size, ods.Size.ToString());
                    matrix.LoadFromDataSource();
                }
                else
                {
                        if (ods.GetValue(col, ods.Size - 1) != null)
                        {
                            if (ods.GetValue(col, ods.Size - 1) != "")
                            {
                                ods.InsertRecord(ods.Size);
                                ods.SetValue("VisOrder", ods.Size - 1, ods.Size.ToString());
                                //oForm.DataSources.DBDataSources.Item(ds).SetValue("VisOrder", ods.Size, ods.Size.ToString());
                                matrix.LoadFromDataSource();
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void arrangematrix(SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, string ds)
        {
            try
            {
                SAPbouiCOM.DBDataSource ods = null;
                int cnt = 0;

                ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item(ds);
                oMatrix.FlushToDataSource();
                for (int x = 0; x < ods.Size; x++)
                {
                    cnt = x + 1;
                    ods.SetValue("VisOrder", x, cnt.ToString());
                }
                oMatrix.LoadFromDataSource();

                
                //if (ds == "@FT_SHIP1")
                //{
                //    string itemdesc = "";
                //    string temp = "";
                //    string brand = "";
                //    string size = "";
                //    string jcclr = "";
                //    string itemname = "";

                //    for (int xx = 1; xx <= oMatrix.RowCount; xx++)
                //    {
                //        if (itemdesc != "")
                //        {
                //            itemdesc = itemdesc + Environment.NewLine;
                //        }
                //        brand = "";
                //        if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_itemcode").Cells.Item(xx).Specific).Value.ToString() != "")
                //        {
                //            temp = "1 X 20' GP CONTRS STC:";
                //            itemname = ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_itemname").Cells.Item(xx).Specific).Value.ToString();
                //            brand = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_brand").Cells.Item(xx).Specific).Selected.Description.ToString();
                //            size = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_size").Cells.Item(xx).Specific).Selected.Description.ToString();
                //            jcclr = ((SAPbouiCOM.ComboBox)oMatrix.Columns.Item("U_jcclr").Cells.Item(xx).Specific).Selected.Description.ToString();
                //        }
                //        if (temp != "" && brand != "")
                //        {
                //            itemdesc = itemdesc + temp + Environment.NewLine;
                //            itemdesc = itemdesc + "1 PCS OF " + size + " " + jcclr + " JERRYCANS" + Environment.NewLine;
                //            itemdesc = itemdesc + brand + " BRAND " + itemname;
                //        }
                //    }
                //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_itemdesc", 0, itemdesc.ToUpper());
                //}
            }
            
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
