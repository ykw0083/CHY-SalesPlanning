using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_Approval
    {
        public static void retrievedata(SAPbouiCOM.Form oForm)
        {
            string sql = "";
            string sptype = oForm.DataSources.UserDataSources.Item("sptype").ValueEx;
            string sysobj = oForm.DataSources.UserDataSources.Item("sysobj").ValueEx;
            string table = oForm.DataSources.UserDataSources.Item("table").ValueEx;
            string table1 = oForm.DataSources.UserDataSources.Item("table1").ValueEx;
            string appcode = oForm.DataSources.UserDataSources.Item("appcode").ValueEx;
            string objecttype = oForm.DataSources.UserDataSources.Item("objecttype").ValueEx;

            //if (sysobj == "Y")
            //    sql = "select T0.DocEntry, T0.ObjType, T0.DocNum, T0.DocDate, case T0.U_APPROVAL when 'P' then 'Pending' when 'A' then 'Approved' else 'Reject' end as U_APPROVAL, T0.U_APPRE, T0.U_APPBY, T0.U_APPDATE, T0.U_APPTIME from [" + table + "] T0 inner join (select DocEntry, sum(isnull(U_SPLANQTY,0)) as U_SPLANQTY from [" + table1 + "] group by DocEntry) T1 on T0.DocEntry = T1.DocEntry and T1.U_SPLANQTY = 0 where T0.U_APPROVAL = '" + status + "' and T0.DocStatus = 'O' and T0.U_SPLAN = 'Y' order by T0.DocNum";
            //else
            //    sql = "select T0.DocEntry, T0.Object, T0.DocNum, T0.U_DocDate, case T0.U_APPROVAL when 'P' then 'Pending' when 'A' then 'Approved' else 'Reject' end as U_APPROVAL, T0.U_APPRE, T0.U_APPBY, T0.U_APPDATE, T0.U_APPTIME from [" + table + "] T0 where T0.U_APPROVAL = '" + status + "' and T0.Status = 'O' order by T0.DocNum";

            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            sql = "select U_AppSQL from [@FT_APPSQL] where Code = '" + appcode + "'";
            oRS.DoQuery(sql);
            if (oRS.RecordCount > 0)
            {
                oRS.MoveFirst();
                sql = oRS.Fields.Item(0).Value.ToString();
                sql = ObjectFunctions.ReplaceParams(oForm, sql, 0);
            }

            oForm.DataSources.DataTables.Item("cfl").ExecuteQuery(sql);

            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;
            oGrid.DataTable = oForm.DataSources.DataTables.Item("cfl");

            SAPbobsCOM.UserTable oUT = null;
            SAPbobsCOM.UserFields oUF = null;
            SAPbobsCOM.Documents doc = null;

            if (sysobj == "Y")
            {
                doc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), objecttype));
                oUF = doc.UserFields;
            }
            else
            {
                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(objecttype);
                oUF = oUT.UserFields;
            }
            string columnname = "";
            foreach (SAPbouiCOM.GridColumn column in oGrid.Columns)
            {
                column.Editable = false;
                columnname = column.UniqueID;
                switch (columnname)
                {
                    case "CardCode":
                        column.TitleObject.Caption = "BP Code";
                        break;
                    case "CardName":
                        column.TitleObject.Caption = "BP Name";
                        break;
                    case "DocTotal":
                        column.TitleObject.Caption = "Document Total";
                        break;
                    case "DocNum":
                        column.TitleObject.Caption = "Doc Num";
                        break;
                    case "DocDate":
                    case "U_DocDate":
                        column.TitleObject.Caption = "Doc Date";
                        break;
                    case "Object":
                    case "ObjType":
                    case "WddCode":
                    case "T_OBJECT":
                        column.Visible = false;
                        break;
                    case "DocEntry":
                        ((SAPbouiCOM.EditTextColumn)column).LinkedObjectType = objecttype;
                        column.Width = 15;
                        column.TitleObject.Caption = "";
                        break;
                    case "T_DOCNUM":
                        if (sptype == "FT_TPPLAN")
                            column.TitleObject.Caption = "TP Doc Num";
                        else
                            column.TitleObject.Caption = "Target Doc Num";
                        break;
                    case "T_DOCDATE":
                        if (sptype == "FT_TPPLAN")
                            column.TitleObject.Caption = "TP Doc Date";
                        else
                            column.TitleObject.Caption = "Target Doc Date";
                        break;
                    case "T_ENTRY":
                        ((SAPbouiCOM.EditTextColumn)column).LinkedObjectType = sptype;
                        column.Width = 15;
                        column.TitleObject.Caption = "";
                        break;
                    default:
                        if (columnname.Substring(0,2) == "U_")
                            column.TitleObject.Caption = oUF.Fields.Item(column.UniqueID).Description;
                        break;
                }
            }


        }

        public static void processDataEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        //string table = oForm.DataSources.UserDataSources.Item("table").ValueEx;
                        //SAPbouiCOM.DBDataSource oDS = oForm.DataSources.DBDataSources.Item(table);
                        //oDS.SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                        //oDS.SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                        //oDS.SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));

                        //string docentry = oDS.GetValue("DocEntry", 0);
                        //string sql = "Update [" + table + "] set U_APPROVAL = '" + oDS.GetValue("U_APPROVAL", 0) + "' " +
                        //            ", U_APPRE = '" + oDS.GetValue("U_APPRE", 0) + "' " +
                        //            ", U_APPBY = '" + SAP.SBOCompany.UserName + "' " +
                        //            ", U_APPDATE = '" + DateTime.Today.ToString("yyyy-MM-dd") + "' " +
                        //            ", U_APPTIME = " + DateTime.Now.ToString("HHmm") + " " +
                        //            " where DocEntry = " + docentry;

                        //SAPbobsCOM.Recordset oRC = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        //oRC.DoQuery(sql);

                        //BubbleEvent = false;
                        //oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;

                        //SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Status").Specific;
                        //string temp = oComboBox.Selected.Value.ToString();
                        //retrievedata(oForm, temp);

                        //oForm.PaneLevel = 1;
                        break;
                }
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
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        if (BusinessObjectInfo.ActionSuccess)
                        {
                            SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Status").Specific;
                            string temp = oComboBox.Selected.Value.ToString();
                            retrievedata(oForm);

                            oForm.PaneLevel = 1;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                string sptype = oForm.DataSources.UserDataSources.Item("sptype").ValueEx;

                                string objecttype = oForm.DataSources.UserDataSources.Item("objecttype").ValueEx;

                                string table = oForm.DataSources.UserDataSources.Item("table").ValueEx;
                                string sysobj = oForm.DataSources.UserDataSources.Item("sysobj").ValueEx;

                                SAPbouiCOM.DBDataSource oDS = oForm.DataSources.DBDataSources.Item(table);
                                oDS.SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                oDS.SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                oDS.SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));

                                string docentry = oDS.GetValue("DocEntry", 0);
                                string docnum = oDS.GetValue("DocNum", 0);

                                string status = "";
                                string app = oForm.DataSources.UserDataSources.Item("U_APP").ValueEx;
                                switch (app)
                                {
                                    case "Y":
                                        status = "Approval Done";
                                        break;
                                    case "N":
                                        status = "Document Rejected";
                                        break;
                                    case "W":
                                        status = "Document Pending";
                                        break;

                                }

                                if (objecttype == "112") // draft sales order
                                {
                                    //sendMsg(objecttype, "U", docentry, oForm.Title + " " + docnum, status, "Draft Document #" + docnum);

                                    SAPbobsCOM.ApprovalRequestsService AppReqService;
                                    SAPbobsCOM.ApprovalRequest AppReq;
                                    SAPbobsCOM.ApprovalRequestsParams AppReqParams;
                                    SAPbobsCOM.ApprovalRequestParams AppReqParam;

                                    AppReqService = (SAPbobsCOM.ApprovalRequestsService)SAP.SBOCompany.GetCompanyService().GetBusinessService(SAPbobsCOM.ServiceTypes.ApprovalRequestsService);
                                    AppReq = (SAPbobsCOM.ApprovalRequest)AppReqService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequest);
                                    AppReqParams = (SAPbobsCOM.ApprovalRequestsParams)AppReqService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestsParams);
                                    AppReqParam = (SAPbobsCOM.ApprovalRequestParams)AppReqService.GetDataInterface(SAPbobsCOM.ApprovalRequestsServiceDataInterfaces.arsApprovalRequestParams);

                                    AppReqParam.Code = int.Parse(oForm.DataSources.UserDataSources.Item("WddCode").ValueEx);
                                    AppReq = AppReqService.GetApprovalRequest(AppReqParam);
                                    //AppReqParams = AppReqService.GetAllApprovalRequestsList();

                                    //AppReParam = AppReqParams.Item(AppReqParams.Count - 1);
                                    string appcode = AppReqParam.Code.ToString();

                                    AppReq = AppReqService.GetApprovalRequest(AppReqParam);
                                    AppReq.ApprovalRequestDecisions.Add();
                                    AppReq.ApprovalRequestDecisions.Item(0).Remarks = oDS.GetValue("U_APPRE", 0);
                                    switch (app)
                                    {
                                        case "Y":
                                            AppReq.ApprovalRequestDecisions.Item(0).Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardApproved;
                                            break;
                                        case "W":
                                            AppReq.ApprovalRequestDecisions.Item(0).Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardPending;
                                            break;
                                        case "N":
                                            AppReq.ApprovalRequestDecisions.Item(0).Status = SAPbobsCOM.BoApprovalRequestDecisionEnum.ardNotApproved;
                                            break;

                                    }
                                    AppReqService.UpdateRequest(AppReq);

                                    //string sql = "Update [" + table + "] set " + appcol + " = '" + oDS.GetValue(appcol, 0) + "' " +
                                    //        " where DocEntry = " + docentry;

                                    //SAPbobsCOM.Recordset oRC = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    //oRC.DoQuery(sql);


                                }
                                else
                                {
                                    string sql = "Update [" + table + "] set U_APP = '" + app + "' " +
                                                ", U_APPRE = '" + oDS.GetValue("U_APPRE", 0) + "' " +
                                                ", U_APPBY = '" + SAP.SBOCompany.UserName + "' " +
                                                ", U_APPDATE = '" + DateTime.Today.ToString("yyyy-MM-dd") + "' " +
                                                ", U_APPTIME = " + DateTime.Now.ToString("HHmm") + " " +
                                                " where DocEntry = " + docentry;

                                    SAPbobsCOM.Recordset oRC = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                                    oRC.DoQuery(sql);

                                    oRC.DoQuery("Exec FT_AfterAPP '" + table + "', " + docentry + ", '" + app + "'");
                                    //if (sptype == "FT_TPPLAN")
                                    //{
                                    //    oRC.DoQuery("Exec FT_Update_Approval '" + table + "', " + docentry + ", '" + sptype + "', '" + app + "'");
                                    //}

                                    sendAppMsg(objecttype, "U", docentry, oForm.Title + " " + docnum, status, oForm.Title.Replace("Approval", "") + docnum);

                                }
                                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Status").Specific;
                                string temp = oComboBox.Selected.Value.ToString();
                                retrievedata(oForm);

                                oForm.PaneLevel = 1;

                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Operation Completed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                                BubbleEvent = false;
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                BubbleEvent = false;
                            }

                        }
                        else if (pVal.ItemUID == "2")
                        {
                            if (oForm.PaneLevel == 2)
                            {
                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                {
                                    oForm.PaneLevel = 1;
                                }
                                else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    SAP.SBOApplication.StatusBar.SetText("Please Update in order to proceed.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                }

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
                string sptype = "", sysobj = "", objecttype = "";
                
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED:
                        sysobj = oForm.DataSources.UserDataSources.Item("sysobj").ValueEx;
                        sptype = oForm.DataSources.UserDataSources.Item("sptype").ValueEx;

                        objecttype = oForm.DataSources.UserDataSources.Item("objecttype").ValueEx;

                        if (pVal.ColUID == "T_ENTRY")
                        {
                            if (sptype == "FT_SPLAN" || sptype == "FT_TPPLAN" || sptype == "FT_CHARGE")
                            { }
                            else
                                break;
                        }
                        else if (pVal.ColUID == "DocEntry")
                        {
                            if (objecttype == "FT_SPLAN" || objecttype == "FT_TPPLAN" || objecttype == "FT_CHARGE")
                            { }
                            else
                                break;
                        }

                        if (sysobj != "Y")
                        {
                            /*
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("cfl");
                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(pVal.ItemUID).Specific;
                            oGrid.Rows.SelectedRows.Add(pVal.Row);

                            if (pVal.ColUID == "T_ENTRY")
                            {
                                oForm.DataSources.UserDataSources.Item("row").Value = pVal.Row.ToString();

                                string docentry = oDataTable.GetValue("T_ENTRY", pVal.Row).ToString();
                                string docnum = oDataTable.GetValue("T_DOCNUM", pVal.Row).ToString();

                                InitForm.FT_SPLANscreenpainter(sptype, sptype + "1", "");

                                SAPbouiCOM.Form oNewForm = SAP.SBOApplication.Forms.ActiveForm;

                                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                                SAPbouiCOM.Condition oCondition = oConditions.Add();

                                oCondition.Alias = "DocEntry";
                                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCondition.CondVal = docentry;

                                SAPbouiCOM.DBDataSource ds = oNewForm.DataSources.DBDataSources.Item("@" + sptype);
                                ds.Query(oConditions);
                                ds = oNewForm.DataSources.DBDataSources.Item("@" + sptype + "1");
                                ds.Query(oConditions);
                                ((SAPbouiCOM.Matrix)oNewForm.Items.Item("grid1").Specific).LoadFromDataSource();

                                UserForm_SalesPlanning.checkrow(oNewForm);
                                oNewForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            }
                            else if (pVal.ColUID == "DocEntry")
                            {
                                oForm.DataSources.UserDataSources.Item("row").Value = pVal.Row.ToString();

                                string docentry = oDataTable.GetValue("DocEntry", pVal.Row).ToString();
                                string docnum = oDataTable.GetValue("DocNum", pVal.Row).ToString();

                                InitForm.FT_SPLANscreenpainter(objecttype, objecttype + "1", "");

                                SAPbouiCOM.Form oNewForm = SAP.SBOApplication.Forms.ActiveForm;

                                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                                SAPbouiCOM.Condition oCondition = oConditions.Add();

                                oCondition.Alias = "DocEntry";
                                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCondition.CondVal = docentry;

                                SAPbouiCOM.DBDataSource ds = oNewForm.DataSources.DBDataSources.Item("@" + objecttype);
                                ds.Query(oConditions);
                                ds = oNewForm.DataSources.DBDataSources.Item("@" + objecttype + "1");
                                ds.Query(oConditions);
                                ((SAPbouiCOM.Matrix)oNewForm.Items.Item("grid1").Specific).LoadFromDataSource();

                                UserForm_SalesPlanning.checkrow(oNewForm);
                                oNewForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            }
                            */
                        }
                        else
                        {
                            /*
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("cfl");
                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(pVal.ItemUID).Specific;
                            oGrid.Rows.SelectedRows.Add(pVal.Row);
                            oForm.DataSources.UserDataSources.Item("WddCode").Value = oDataTable.GetValue("WddCode", pVal.Row).ToString();

                            if (pVal.ColUID == "T_ENTRY")
                            {
                                oForm.DataSources.UserDataSources.Item("row").Value = pVal.Row.ToString();

                                string docentry = oDataTable.GetValue("T_ENTRY", pVal.Row).ToString();
                                string docnum = oDataTable.GetValue("T_DOCNUM", pVal.Row).ToString();

                                InitForm.FT_SPLANscreenpainter(sptype, sptype + "1", "");

                                SAPbouiCOM.Form oNewForm = SAP.SBOApplication.Forms.ActiveForm;

                                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                                SAPbouiCOM.Condition oCondition = oConditions.Add();

                                oCondition.Alias = "DocEntry";
                                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCondition.CondVal = docentry;

                                SAPbouiCOM.DBDataSource ds = oNewForm.DataSources.DBDataSources.Item("@" + sptype);
                                ds.Query(oConditions);
                                ds = oNewForm.DataSources.DBDataSources.Item("@" + sptype + "1");
                                ds.Query(oConditions);
                                ((SAPbouiCOM.Matrix)oNewForm.Items.Item("grid1").Specific).LoadFromDataSource();

                                UserForm_SalesPlanning.checkrow(oNewForm);
                                oNewForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            }
                            else if (pVal.ColUID == "DocEntry")
                            {
                                oForm.DataSources.UserDataSources.Item("row").Value = pVal.Row.ToString();

                                string docentry = oDataTable.GetValue("DocEntry", pVal.Row).ToString();
                                string docnum = oDataTable.GetValue("DocNum", pVal.Row).ToString();

                                InitForm.FT_SPLANscreenpainter(objecttype, objecttype + "1", "");

                                SAPbouiCOM.Form oNewForm = SAP.SBOApplication.Forms.ActiveForm;

                                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                                SAPbouiCOM.Condition oCondition = oConditions.Add();

                                oCondition.Alias = "DocEntry";
                                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                                oCondition.CondVal = docentry;

                                SAPbouiCOM.DBDataSource ds = oNewForm.DataSources.DBDataSources.Item("@" + objecttype);
                                ds.Query(oConditions);
                                ds = oNewForm.DataSources.DBDataSources.Item("@" + objecttype + "1");
                                ds.Query(oConditions);
                                ((SAPbouiCOM.Matrix)oNewForm.Items.Item("grid1").Specific).LoadFromDataSource();

                                UserForm_SalesPlanning.checkrow(oNewForm);
                                oNewForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                            }
                            */
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("Status").Specific;
                        string temp = oComboBox.Selected.Value.ToString();
                        retrievedata(oForm);
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(pVal.ItemUID).Specific;
                            oGrid.Rows.SelectedRows.Add(pVal.Row);
                            oForm.DataSources.UserDataSources.Item("row").Value = pVal.Row.ToString();
                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("cfl");
                            oForm.DataSources.UserDataSources.Item("WddCode").Value = oDataTable.GetValue("WddCode", pVal.Row).ToString();
                            
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            sptype = oForm.DataSources.UserDataSources.Item("sptype").ValueEx;
                            SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(pVal.ItemUID).Specific;
                            oGrid.Rows.SelectedRows.Add(pVal.Row);
                            oForm.DataSources.UserDataSources.Item("row").Value = pVal.Row.ToString();

                            SAPbouiCOM.DataTable oDataTable = oForm.DataSources.DataTables.Item("cfl");
                            oForm.DataSources.UserDataSources.Item("WddCode").Value = oDataTable.GetValue("WddCode", pVal.Row).ToString();

                            string docentry = oDataTable.GetValue("DocEntry", pVal.Row).ToString();
                            SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                            SAPbouiCOM.Condition oCondition = oConditions.Add();

                            oCondition.Alias = "DocEntry";
                            oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                            oCondition.CondVal = docentry;

                            string table = oForm.DataSources.UserDataSources.Item("table").ValueEx;
                            SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item(table);
                            ds.Query(oConditions);

                            oForm.DataSources.UserDataSources.Item("U_APP").Value = oForm.DataSources.UserDataSources.Item("Status").ValueEx;


                            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            objecttype = oForm.DataSources.UserDataSources.Item("objecttype").ValueEx;

                            string sql = "select 1 from [@FT_APPUSER] where U_userid = '" + SAP.SBOCompany.UserName + "' and U_objcode = '" + objecttype + "'";
                            oRS.DoQuery(sql);
                            if (oRS.RecordCount == 0)
                            {
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("You are not authorized for approval.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                break;
                            }

                            oForm.PaneLevel = 2;
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                //SAP.SBOApplication.MessageBox("Item Event After " + ex.Message, 1, "Ok", "", "");
                SAP.SBOApplication.MessageBox($"Item Event After {ex.Source} {ex.StackTrace} {System.Environment.NewLine} {ex.Message}", 1, "Ok", "", "");
            }
        }

        private static void sendAppMsg(string ObjType, string TrxType, string DocEntry, string msgsubject, string msg, string linkvalue)
        {

            ApprovalMsgClass oSystemMsgClass = new ApprovalMsgClass();

            oSystemMsgClass.IsApprove = true;
            oSystemMsgClass.ObjType = ObjType;
            oSystemMsgClass.TrxType = TrxType;

            string temp = oSystemMsgClass.TrxType == "A" ? "Added." : "Updated.";
            oSystemMsgClass.MsgSubject = msgsubject;
            oSystemMsgClass.Msg = msg;
            oSystemMsgClass.ColumnName = "Description";

            int obj = 0;
            if (int.TryParse(ObjType, out obj))
            {
                if (obj > 0)
                {

                    oSystemMsgClass.LinkObj = ObjType;
                    oSystemMsgClass.LineValue = linkvalue;
                    oSystemMsgClass.LinkKey = DocEntry;
                }
            }

            oSystemMsgClass.SendMsg();

        }

    }
}
