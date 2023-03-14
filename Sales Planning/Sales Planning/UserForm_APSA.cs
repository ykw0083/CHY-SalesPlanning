using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_APSA
    {
        public static void processDataEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Data Event Before " + ex.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }

        }
        public static void processDataEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //oRS.DoQuery("select AutoKey from ONNM where ObjectCode = 'FT_APSA'");
                        oRS.DoQuery("select T0.DocNum, T0.U_BookNo, T0.U_SODOCNUM, T1.U_SOENTRY from ONNM inner join [@FT_APSA] T0 on ONNM.AutoKey - 1 = T0.DocEntry inner join [@FT_APSA1] T1 on T0.DocEntry = T1.DocEntry where ONNM.ObjectCode = 'FT_APSA'");
                        string DocNum = oRS.Fields.Item(0).Value.ToString().Trim();
                        string BookNo = oRS.Fields.Item(1).Value.ToString().Trim();
                        string SODOCNUM = oRS.Fields.Item(2).Value.ToString().Trim();
                        string SOENTRY = oRS.Fields.Item(3).Value.ToString().Trim();

                        //sendMsg("A", DocNum, BookNo, SODOCNUM, SOENTRY);
                        string emailmsg = "Base on Booking No " + BookNo + ".";

                        EmailClass email = new EmailClass();
                        email.EmailMsg = emailmsg;
                        email.EmailSubject = "Draft SA No " + DocNum + " from SC No " + SODOCNUM + " Generated.";
                        email.ObjType = "FT_APSA";

                        email.SendEmail();

                        break;
                }

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "grid1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.Row >= 0)
                                {
                                    if (oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("Status", 0) != "O")
                                    {
                                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status = 'O' allowed Add Booking Details.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                if (validateRows(oForm))
                                {
                                    //string value = "";
                                    //string sadate = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_SADATE", 0);
                                    //if (rs.RecordCount > 0)
                                    //{
                                    //    rs.MoveFirst();
                                    //    value = rs.Fields.Item(0).Value.ToString();
                                    //    oForm.DataSources.DBDataSources.Item("@FT_APSA").SetValue("Period", 0, value);

                                    //}
                                    //else
                                    //{
                                    //    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Period is no valid.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    //    BubbleEvent = false;
                                    //    return;
                                    //}
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    string DocNum = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocNum", 0).Trim();
                                    string BookNo = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).Trim();
                                    string SODOCNUM = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_SODOCNUM", 0).Trim();
                                    string SOENTRY = oForm.DataSources.DBDataSources.Item("@FT_APSA1").GetValue("U_SOENTRY", 0).Trim();
                                    sendMsg("U", DocNum, BookNo, SODOCNUM, SOENTRY);
                                }

                            }
                        }
                        if (pVal.ItemUID == "cb_gendo")
                        {

                            //if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("Status", 0) != "O")
                                {
                                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status = 'O' allowed Generate DO.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }

                            }
                            else
                            {
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only OK mode allowed Generate DO.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                            }
                        }
                        if (pVal.ItemUID == "cb_batch")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("Status", 0) != "O")
                                {
                                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status = 'O' allowed Generate DO.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                }
                                else
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                    oMatrix.FlushToDataSource();
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                                }

                            }
                            else
                            {
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Manage Batch is not allowed in this mode.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                            }

                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Item Event Before " + ex.Message, 1, "Ok", "", "");
                BubbleEvent = false;
            }
        }

        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "grid1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.Row >= 0)
                                {
                                    int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("@FT_APSA1").GetValue("DocEntry", pVal.Row - 1).ToString());
                                    int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("@FT_APSA1").GetValue("LineId", pVal.Row - 1).ToString());
                                    string bookno = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).Trim();
                                    if (pVal.ColUID == "U_BookNo")
                                    {
                                        InitForm.CONM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APSAC", "U_CONNO,U_BookNo", bookno);
                                    }
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        if (oCFLEvento.SelectedObjects != null)
                        {
                            SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                            string val = oDataTable.GetValue(0, 0).ToString();
                            if (pVal.ColUID == "U_WHSCODE")
                            {
                                try
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String = val;
                                }
                                catch { }
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                        oForm.EnableMenu("1284", true);//cancel
                        oForm.EnableMenu("1286", false);//close
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                        oForm.EnableMenu("1284", false);//cancel
                        oForm.EnableMenu("1286", false);//close
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "cb_gendo")
                        {

                            int docentry = genDO(oForm);
                            if (docentry < 0)
                            {
                                return;
                            }

                            SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                            if (!oDoc.GetByKey(docentry))
                            {
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Cannot find Deliery Order from DocEntry " + docentry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return;
                            }

                            int docnum = oDoc.DocNum;

                            string sadocentry = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocEntry", 0).Trim();
                            string sono = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_SODOCNUM", 0).Trim();
                            string bookno = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).Trim();
                            string sadocnum = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocNum", 0).Trim();

                            string emailmsg = "Base on" + Environment.NewLine + "Booking No " + bookno + Environment.NewLine + "Shipping Advise No " + sadocnum + ".";

                            EmailClass email = new EmailClass();
                            email.EmailMsg = emailmsg;
                            email.EmailSubject = "Delivery Order No " + docnum.ToString() + " from SC No " + sono + " Generated.";
                            email.ObjType = "FT_APSA";

                            email.SendEmail();
                            //oForm.DataSources.DBDataSources.Item("@FT_APSA").SetValue("Status", 0, "C");
                            //oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                            //oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("DO Generated please check from Delivery Order Screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            oForm.Close();
                            return;
                        }
                        else if (pVal.ItemUID == "cb_batch")
                        {

                            InitForm.batchno(oForm.UniqueID, "@FT_APSA1", "@FT_APSA2");
                            oForm.Freeze(true);
                        }
                        else if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocNum", 0) == "0")
                                { }
                                else
                                {
                                    SAP.SBOApplication.ActivateMenuItem("1289");
                                    SAP.SBOApplication.ActivateMenuItem("1288");
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        if (pVal.ItemUID == "grid1")
                        {
                            SAP.currentMatrixRow = pVal.Row;
                            if (pVal.Row > 0)
                               oForm.EnableMenu("1293", true);//delete row
                            else
                                oForm.EnableMenu("1293", false);//delete row
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if (pVal.ItemUID == "grid1")
                        {
                            oForm.EnableMenu("1293", false);//delete row
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Item Event After " + ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processMenuEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            try
            {

                SAPbouiCOM.DBDataSource oDS = oForm.DataSources.DBDataSources.Item("@FT_APSA");
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                long docentry = 0;
                string sql = "";
                int value = 0;

                if (pVal.MenuUID == "1284") // cancel doc
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only OK mode allowed Cancel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("Status", 0) != "O")
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status O allowed Cancel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                    docentry = int.Parse(oDS.GetValue("DocEntry", 0));
                    sql = "select count(*) from DLN1 where U_SAENTRY = " + docentry;
                    rs.DoQuery(sql);
                    rs.MoveFirst();
                    value = int.Parse(rs.Fields.Item(0).Value.ToString());
                    if (value > 0)
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Delivery Order found, Cancel is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                }
                else if (pVal.MenuUID == "1293") // delete row
                {
                    if (oForm.ActiveItem == "grid1")
                    {
                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                        if (oMatrix.RowCount == 1)
                        {
                            BubbleEvent = false;
                        }
                        else
                        {
                            oMatrix.DeleteRow(SAP.currentMatrixRow);
                            arrangematrix(oForm, oMatrix, "@FT_APSA1");

                            oDS = oForm.DataSources.DBDataSources.Item("@FT_APSA2");
                            int row = oDS.Size;
                            for (int x = row - 1; x >= 0; x--)
                            {
                                oDS.RemoveRecord(oDS.Size - 1);
                            }

                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Delete Row will reset Batch information. Please reassign Batch No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                        }
                    }
                    BubbleEvent = false;
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                else if (pVal.MenuUID == "1285")
                {
                    if (oDS.GetValue("Canceled", 0) == "Y")
                    {

                    }
                    else
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document Closed, Restore is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                }

                rs = null;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Menu Event Before " + ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processMenuEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal)
        {
            try
            {
                if (pVal.MenuUID == "1281") // find
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item("sodocnum").Specific).Value = "";
                    ((SAPbouiCOM.EditText)oForm.Items.Item("dodocnum").Specific).Value = "";
                    //((SAPbouiCOM.EditText)oForm.Items.Item("status").Specific).Value = "";

                }
                if (pVal.MenuUID == "1284") // cancel doc
                {

                }

                //SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //if (pVal.MenuUID == "1283" || pVal.MenuUID == "1282" || pVal.MenuUID == "1287")
                //{
                //    if (pVal.MenuUID == "1283")
                //    {
                //        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                //    }

                //    docentry = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                //    docnum = oForm.DataSources.UserDataSources.Item("docnum").Value.ToString();
                //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_sdoc", 0, docentry);
                //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_pino", 0, docnum);

                //    rs.DoQuery("select max(U_set) from [@FT_SHIPD] where U_sdoc = " + docentry.ToString());
                //    if (rs.RecordCount > 0)
                //    {
                //        rs.MoveFirst();
                //        set = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_set", 0, set.ToString());
                //    }
                //    //rs.DoQuery("select max(docentry) from [@FT_SHIPD]");
                //    //if (rs.RecordCount > 0)
                //    //{
                //    //    docnum = long.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                //    //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("docnum", 0, docnum.ToString());
                //    //}

                //    SAPbouiCOM.Matrix oMatrix = null;

                //    oForm.DataSources.DBDataSources.Item("@FT_SHIP1").SetValue("VisOrder", 0, "1");
                //    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                //    oMatrix.LoadFromDataSource();

                //    oForm.DataSources.DBDataSources.Item("@FT_SHIP2").SetValue("VisOrder", 0, "1");
                //    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
                //    oMatrix.LoadFromDataSource();

                //    oForm.DataSources.DBDataSources.Item("@FT_SHIP3").SetValue("VisOrder", 0, "1");
                //    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
                //    oMatrix.LoadFromDataSource();
                //}
                //rs = null;

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Menu Event After " + ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processRightClickEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Right Click Event Before " + ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processRightClickEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal)
        {
            try
            {
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Right Click Event After " + ex.Message, 1, "Ok", "", "");
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
            //temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_DODATE", 0).ToString();
            //if (temp == null || temp.Trim() == "")
            //{
            //    SAP.SBOApplication.StatusBar.SetText("DO Date is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return false;
            //}

            SAPbouiCOM.DBDataSource ods = null;
            SAPbouiCOM.Matrix oMatrix = null;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FT_APSA1");
            SAPbouiCOM.DBDataSource oDS2 = oForm.DataSources.DBDataSources.Item("@FT_APSA2");

            oMatrix.FlushToDataSource();
            double qty = 0;
            double openqty = 0;
            bool found = false;
            bool wrongbatch = false;
            string whscode = "";
            string itemcode = "";
            string visorder = "";
            for (int x = 0; x < ods.Size; x++)
            {
                visorder = ods.GetValue("VisOrder", x).Trim();

                if (ods.GetValue("U_QUANTITY", x) == null)
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for #" + visorder + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (ods.GetValue("U_OPENQTY", x) == null)
                {
                    SAP.SBOApplication.StatusBar.SetText("Open Quantity for #" + visorder + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(ods.GetValue("U_QUANTITY", x), out qty))
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(ods.GetValue("U_OPENQTY", x), out openqty))
                {
                    SAP.SBOApplication.StatusBar.SetText("Open Quantity for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                whscode = ods.GetValue("U_WHSCODE", x).Trim();
                itemcode = ods.GetValue("U_ITEMCODE", x).Trim();

                if (openqty == 0 && qty == 0)
                { }
                else if (openqty < qty)
                {
                    SAP.SBOApplication.StatusBar.SetText("Open Quantity for #" + visorder + " is not enough.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                else if (openqty > 0 && qty > 0)
                {
                    found = true;
                    if (whscode == "")
                    {
                        SAP.SBOApplication.StatusBar.SetText("Warehouse Code #" + visorder + " is emtpy.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }

                    int row = oDS2.Size;
                    for (int y = row - 1; y >= 0; y--)
                    {
                        if (oDS2.GetValue("U_BASEVIS", y).Trim() == visorder)
                        {
                            if (oDS2.GetValue("U_WHSCODE", y).Trim() == whscode && oDS2.GetValue("U_ITEMCODE", y).Trim() == itemcode)
                            {
                            }
                            else
                            {
                                wrongbatch = true;
                                oDS2.RemoveRecord(oDS2.Size - 1);
                            }
                        }
                    }

                }
                else
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for #" + visorder + " cannot less than zero.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

            }

            if (wrongbatch)
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Edit Warehouse will reset Batch information. Please assign Batch No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
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
        private static bool validateBatches(SAPbouiCOM.Form oForm)
        {
            string temp = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_DODATE", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                SAP.SBOApplication.StatusBar.SetText("DO Date is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            SAPbouiCOM.DBDataSource ods = null;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@FT_APSA1");
            SAPbouiCOM.DBDataSource oDS2 = oForm.DataSources.DBDataSources.Item("@FT_APSA2");

            double qty = 0;
            double batchqty = 0;
            double tempqty = 0;
            string whscode = "";
            string itemcode = "";
            string visorder = "";
            for (int x = 0; x < ods.Size; x++)
            {
                visorder = ods.GetValue("VisOrder", x).Trim();

                if (!double.TryParse(ods.GetValue("U_QUANTITY", x), out qty))
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                whscode = ods.GetValue("U_WHSCODE", x).Trim();
                itemcode = ods.GetValue("U_ITEMCODE", x).Trim();

                
                if (qty == 0)
                { }
                else if (qty > 0)
                {
                    oRS.DoQuery("select ManBtchNum from OITM where ItemCode = '" + itemcode + "'");
                    oRS.MoveFirst();
                    if (oRS.Fields.Item(0).Value.ToString().Trim() == "Y")
                    {
                        batchqty = 0;
                        for (int y = 0; y < oDS2.Size; y++)
                        {
                            if (oDS2.GetValue("U_BASEVIS", y).Trim() == visorder)
                            {
                                if (oDS2.GetValue("U_WHSCODE", y).Trim() == whscode && oDS2.GetValue("U_ITEMCODE", y).Trim() == itemcode)
                                {
                                    if (!double.TryParse(oDS2.GetValue("U_QUANTITY", y), out tempqty))
                                    {
                                        SAP.SBOApplication.StatusBar.SetText("Batch Quantity for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        return false;
                                    }
                                    batchqty += tempqty;

                                }
                            }
                        }

                        if (batchqty != qty)
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("#" + visorder +" Quantity and Batch Quantity is not tally. Please assign Batch No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }

                    }
                }
                else
                {
                    SAP.SBOApplication.StatusBar.SetText("Quantity for #" + visorder + " cannot less than zero.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

            }

            return true;
        }
        public static void arrangevisorder(SAPbouiCOM.Form oForm, string ds)
        {
            try
            {
                SAPbouiCOM.DBDataSource ods = null;
                int cnt = 0;

                ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item(ds);
                for (int x = 0; x < ods.Size; x++)
                {
                    cnt++;
                    ods.SetValue("VisOrder", x, cnt.ToString());
                }

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("arrangevisorder " + ex.Message, 1, "Ok", "", "");
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
                    cnt++;
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
                SAP.SBOApplication.MessageBox("arrangematrix " + ex.Message, 1, "Ok", "", "");
            }
        }

        private static int genDO (SAPbouiCOM.Form oForm)
        {
            try
            {
                if (!validateRows(oForm)) return -1;
                if (!validateBatches(oForm)) return -1;

                int rtn = 0;
                int docentry = 0;
                int lineid = 0;

                string temp = "";
                DateTime docdate;
                string dsname = "";
                string columnname = "";
                string BookNo = "";
                int rows = 0;

                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                dsname = "FT_APSA";
                SAPbouiCOM.DBDataSource oDBs = oForm.DataSources.DBDataSources.Item("@" + dsname);

                SAPbobsCOM.UserTable oUT = null;
                SAPbobsCOM.UserFields oUF = null;
                //SAPbobsCOM.UserFieldsMD oUFMD = (SAPbobsCOM.UserFieldsMD)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;

                docentry = int.Parse(oDBs.GetValue("DocEntry", 0).Trim());
                oDoc.UserFields.Fields.Item("U_SADOCNUM").Value = oDBs.GetValue("DocNum", 0).Trim();
                oDoc.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                //oDoc.Series = 3;

                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    //oUFMD.GetByKey("@" + dsname, oUF.Fields.Item(x).FieldID);
                    //columnname = "U_" + oUFMD.Name;
                    columnname = oUF.Fields.Item(x).Name;

                    switch (columnname)
                    {
                        case "U_CARDNAME":
                        case "U_SODOCENTRY":
                        case "U_SODOCNUM":
                        case "U_DOCDATE":
                        case "U_DODOCNUM":
                            break;
                        case "U_CARDCODE":
                            oDoc.CardCode = oDBs.GetValue(columnname, 0).Trim();
                            break;
                        case "U_NUMATCARD":
                            oDoc.NumAtCard = oDBs.GetValue(columnname, 0).Trim();
                            break;
                        case "U_DODATE":
                            temp = oDBs.GetValue(columnname, 0).Substring(0, 4);
                            temp += "-" + oDBs.GetValue(columnname, 0).Substring(4, 2);
                            temp += "-" + oDBs.GetValue(columnname, 0).Substring(6, 2);

                            docdate = DateTime.Parse(temp);
                            oDoc.DocDate = docdate;
                            break;
                        default:
                            if (columnname == "U_BookNo")
                                BookNo = oDBs.GetValue(columnname, 0).Trim();
                            else
                                oDoc.UserFields.Fields.Item(columnname).Value = oDBs.GetValue(columnname, 0).Trim();
                            break;
                    }
                    //dbs.SetValue(columnname, 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue(columnname, 0));
                }

                string visorder = "";
                string itemcode = "";
                string whscode = "";
                dsname = "FT_APSA1";
                oDBs = oForm.DataSources.DBDataSources.Item("@" + dsname);

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;

                rows = oDBs.Size;
                int cnt = 0;
                int cntbatch = 0;
                for (int y = 0; y < rows; y++)
                {
                    cnt++;
                    if (cnt > 1)
                    {
                        oDoc.Lines.Add();
                        oDoc.Lines.SetCurrentLine(cnt - 1);
                    }

                    visorder = oDBs.GetValue("VisOrder", y).Trim();
                    oDoc.Lines.UserFields.Fields.Item("U_SAENTRY").Value = oDBs.GetValue("DocEntry", y).Trim();
                    oDoc.Lines.UserFields.Fields.Item("U_SALINE").Value = oDBs.GetValue("LineId", y).Trim();
                    oDoc.Lines.UserFields.Fields.Item("U_BookNo").Value = BookNo;
                    oDoc.Lines.BaseType = 17;

                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        //oUFMD.GetByKey("@" + dsname, oUF.Fields.Item(x).FieldID);
                        //columnname = "U_" + oUFMD.Name;
                        columnname = oUF.Fields.Item(x).Name;

                        switch (columnname)
                        {
                            case "U_BookNo":
                            case "U_ITEMNAME":
                            case "U_OPENQTY":
                            case "U_BQTY":
                                break;
                            case "U_ITEMCODE":
                                itemcode = oDBs.GetValue(columnname, y).Trim();
                                oDoc.Lines.ItemCode = oDBs.GetValue(columnname, y).Trim();
                                break;
                            case "U_WHSCODE":
                                whscode = oDBs.GetValue(columnname, y).Trim();
                                oDoc.Lines.WarehouseCode = oDBs.GetValue(columnname, y).Trim();
                                break;
                            case "U_SOENTRY":
                                oDoc.Lines.BaseEntry = int.Parse(oDBs.GetValue(columnname, y).Trim());
                                break;
                            case "U_SOLINE":
                                oDoc.Lines.BaseLine = int.Parse(oDBs.GetValue(columnname, y).Trim());
                                break;
                            case "U_QUANTITY":
                                oDoc.Lines.Quantity = double.Parse(oDBs.GetValue(columnname, y).Trim());
                                break;
                            default:
                                oDoc.Lines.UserFields.Fields.Item(columnname).Value = oDBs.GetValue(columnname, y).Trim();
                                break;
                        }
                    }
                    #region batchno
                    cntbatch = 0;
                    oRS.DoQuery("select U_BATCHNUM, U_QUANTITY from [@FT_APSA2] where DocEntry = " + docentry.ToString() + " and U_BASEVIS = " + visorder + " and U_ITEMCODE = '" + itemcode + "' and U_WHSCODE = '" + whscode + "' and U_QUANTITY > 0");
                    oRS.MoveFirst();
                    while (!oRS.EoF)
                    {
                        if (oRS.Fields.Item(0).Value.ToString().Trim() != "")
                        {
                            cntbatch++;
                            if (cntbatch > 1)
                            {
                                oDoc.Lines.BatchNumbers.Add();
                                oDoc.Lines.BatchNumbers.SetCurrentLine(cntbatch - 1);
                            }

                            oDoc.Lines.BatchNumbers.BatchNumber = oRS.Fields.Item(0).Value.ToString().Trim();
                            oDoc.Lines.BatchNumbers.Quantity = double.Parse(oRS.Fields.Item(1).Value.ToString().Trim());

                        }

                        oRS.MoveNext();
                    }
                    #endregion

                }

                FT_ADDON.SAP.SBOCompany.StartTransaction();
                if (oDoc.Add() != 0)
                {
                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText(FT_ADDON.SAP.SBOCompany.GetLastErrorCode() + " " + FT_ADDON.SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return -1;
                }

                rtn = int.Parse(FT_ADDON.SAP.SBOCompany.GetNewObjectKey());

                oDoc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                if (!oDoc.GetByKey(rtn))
                {
                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Cannot find Deliery Order from DocEntry " + docentry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return -1;
                }
                int docnum = oDoc.DocNum;

                oRS.DoQuery("update [@FT_APSA] set Status = 'C', U_DODOCNUM = " + docnum.ToString() + " where DocEntry = " + docentry.ToString());

                #region booking
                int dolinenum = 0;
                string sql = "";
                int dodocentry = rtn;
                for (int y = 0; y < oDBs.Size; y++)
                {
                    lineid = int.Parse(oDBs.GetValue("LineId", y).Trim());
                    oRS.DoQuery("select LineNum from DLN1 where U_SAENTRY = " + docentry.ToString() + " and U_SALINE = " + lineid.ToString());
                    oRS.MoveFirst();

                    if (!int.TryParse(oRS.Fields.Item(0).Value.ToString(), out dolinenum))
                    {
                        if (FT_ADDON.SAP.SBOCompany.InTransaction)
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Unable to get DO info.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            return -1;
                        }                        
                    }
                    sql = "insert into [@FT_APDOC] (DocEntry, LineId, VisOrder, U_LINENO, U_CONNO, U_BookNo)";
                    sql = sql + " select " + dodocentry.ToString() + ", LineId, VisOrder, " + dolinenum.ToString() + ", U_CONNO, U_BookNo from [@FT_APSAC] where DocEntry = " + docentry.ToString() + " and U_LINENO = " + lineid.ToString();

                    oRS.DoQuery(sql);

                }

                #endregion

                if (FT_ADDON.SAP.SBOCompany.InTransaction)
                    FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

                return rtn;
            }
            catch (Exception ex)
            {
                if (FT_ADDON.SAP.SBOCompany.InTransaction)
                {
                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText(FT_ADDON.SAP.SBOCompany.GetLastErrorCode() + " " + FT_ADDON.SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                SAP.SBOApplication.MessageBox("genDO " + ex.Message, 1, "Ok", "", "");

                return -1;
            }
        }

        private static void sendMsg(string trxtype, string DocNum, string BookNo, string SODOCNUM, string SOENTRY)
        {

            SystemMsgClass oSystemMsgClass = new SystemMsgClass();

            oSystemMsgClass.ObjType = "FT_APSA";
            oSystemMsgClass.TrxType = trxtype;

            string temp = oSystemMsgClass.TrxType == "A" ? "Added." : "Updated.";
            oSystemMsgClass.MsgSubject = "Draft SA " + temp;
            oSystemMsgClass.Msg = "Booking #" + BookNo + " Draft SA #" + DocNum;
            oSystemMsgClass.ColumnName = "Description";

            oSystemMsgClass.LinkObj = "17";
            oSystemMsgClass.LineValue = "SO #" + SODOCNUM;
            oSystemMsgClass.LinkKey = SOENTRY;

            oSystemMsgClass.SendMsg();
            
        }
    }
}
