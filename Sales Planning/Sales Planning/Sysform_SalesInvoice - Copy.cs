using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.AYS
{
    class Sysform_SalesInvoice
    {
        public static void processRightClickEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
            try
            {
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processRightClickEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal)
        {
        }
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD)
                {
                    SAPbouiCOM.Item oItem = null;
                    SAPbouiCOM.EditText oEdit = null;
                    oItem = oForm.Items.Add("U_ADDONIND", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oItem.Left = oForm.Width + 100;
                    oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                    oEdit.DataBind.SetBound(true, "OINV", "U_ADDONIND");

                    //SAPbouiCOM.UserDataSource uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    //uds.Value = "";
                }
                //if (oForm.DataSources.UserDataSources.Count > 0)
                //{
                //    if (oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString() != "")
                //    {
                //        string FUID = oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString();
                //        SAPbouiCOM.Form oSForm = SAP.SBOApplication.Forms.Item(FUID);
                //        oSForm.Select();
                //        BubbleEvent = false;
                //    }
                //}
                //if (((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).Value != "")
                //{
                //    string FUID = ((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).Value.ToString();
                //    SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FUID);
                //    oSForm.Select();
                //    BubbleEvent = false;
                //}
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        //if (pVal.ItemUID == "38")
                        //{
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //    {
                        //        if (pVal.Row >= 0 && pVal.ColUID == "256")
                        //        {
                        //            int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("INV1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                        //            int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("INV1").GetValue("LINENUM", pVal.Row - 1).ToString());

                        //            InitForm.TEXT(oForm.UniqueID, docentry, linenum, pVal.Row, "INV1", "38", oForm.DataSources.DBDataSources.Item("INV1").GetValue("TEXT", pVal.Row - 1).ToString());
                        //            BubbleEvent = false;
                        //        }
                        //    }
                        //}
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("U_ADDONIND").Specific;
                                oEditText.String = "I";
                            }
                            
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
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
                        //if (pVal.ItemUID == "38")
                        //{
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //    {
                        //        if (pVal.Row >= 0 && pVal.ColUID != "256")
                        //        {
                        //            int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("INV1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                        //            int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("INV1").GetValue("LINENUM", pVal.Row - 1).ToString());
                        //            //if (pVal.ColUID == "U_BookNo")
                        //            //{
                        //            //    InitForm.CONM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APDOC", "U_CONNO,U_BookNo");
                        //            //}
                        //            //else if (pVal.ColUID == "U_ProdType")
                        //            //{
                        //            //    InitForm.DOPTM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APDOPT", "");
                        //            //}
                        //            //else
                        //            //{
                        //                InitForm.SDM(oForm.UniqueID, docentry, linenum, pVal.Row, "INV1", "38");
                        //            //}
                        //        }
                        //    }
                        //}
                        break;
                    default:
                        break;
                }
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processMenuEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            return;            
        }
        public static void processMenuEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.MenuEvent pVal)
        {
            return;
        }
        //public static void processDataEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        //{
        //    try
        //    {
        //        switch (BusinessObjectInfo.EventType)
        //        {
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
        //                if (oForm.TypeEx == "FT_CHARGE")
        //                {
        //                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                    SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //                    rs.DoQuery("Select top 1 T0.*,T3.U_BASEOBJ as [BASETYPE],T4.U_Type from [@FT_CHARGE] T0  " +
        //                    " INNER JOIN[@FT_CHARGE1] T1 ON T0.DOCENTRY = T1.DOCENTRY " +
        //                    " INNER JOIN[@FT_TPPLAN1] T2 ON T1.U_BASEENT = T2.DocEntry AND T1.U_BASEOBJ = T2.Object " +
        //                    " INNER JOIN[@FT_SPLAN1] T3 ON T2.U_BASEENT = T3.DocEntry AND T2.U_BASEOBJ = T3.Object " +
        //                    " INNER JOIN OCRD T4 on  T0.U_CARDCODE = T4.CARDCODE ORDER BY T0.DOCENTRY DESC");
        //                    int docEntry = 0, cnt = 0, retcode = 0;
        //                    string oIssueKey = "", oReceiptKey = "", oJEKey = "", oDelKey = "", cardcode = "", docnum = "", baseType = "", oDeldocnum = "", bpType = "", cogsAcct = "";
        //                    DateTime date = new DateTime();
        //                    double different = 0;

        //                    if (rs.RecordCount > 0)
        //                    {
        //                        docEntry = int.Parse(rs.Fields.Item("DocEntry").Value.ToString());
        //                        docnum = rs.Fields.Item("Docnum").Value.ToString();
        //                        date = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
        //                        cardcode = rs.Fields.Item("U_CARDCODE").Value.ToString();
        //                        baseType = rs.Fields.Item("BASETYPE").Value.ToString();

        //                        bpType = rs.Fields.Item("U_Type").Value.ToString();
        //                        switch (bpType.ToUpper())
        //                        {
        //                            case "LOCAL":
        //                                cogsAcct = "500001";
        //                                break;
        //                            case "OVERSEA":
        //                                cogsAcct = "500100";
        //                                break;
        //                            case "INTER-COMPANY":
        //                                cogsAcct = "500200";
        //                                break;
        //                        }
        //                    }
        //                    SAPbobsCOM.Documents oReceipt = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
        //                    SAPbobsCOM.Documents oIssue = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
        //                    SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

        //                    rs1.DoQuery("select * from [@FT_CHARGE1] T0 inner join OITM  T1 on T0.U_SOITEMCO = T1.itemcode " +
        //                        " inner join OITM T2  on T0.U_itemcode = T2.itemcode where T0.docentry = " + docEntry + " and T0.U_SOITEMCO <> T0.U_ITEMCODE and isnull(T1.invntitem,'N') = 'Y' and isnull(T2.invntitem,'N') = 'Y' ");
        //                    if (rs1.RecordCount > 0)
        //                    {
        //                        if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();

        //                        System.Data.DataTable dt = new System.Data.DataTable();
        //                        dt.Columns.Add("OriItemCode", typeof(string));
        //                        dt.Columns.Add("RevItemCode", typeof(string));
        //                        dt.Columns.Add("LineId", typeof(int));
        //                        dt.Columns.Add("OriItemCost", typeof(double));
        //                        dt.Columns.Add("RevItemCost", typeof(double));
        //                        dt.Columns.Add("Different", typeof(double));


        //                        rs1.MoveFirst();


        //                        oReceipt.DocDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
        //                        oReceipt.TaxDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
        //                        oReceipt.UserFields.Fields.Item("U_IsType").Value = "CHRG_OUT";

        //                        oIssue.DocDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
        //                        oIssue.TaxDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
        //                        oIssue.UserFields.Fields.Item("U_IsType").Value = "CHRG_OUT";
        //                        while (!rs1.EoF)
        //                        {
        //                            if (cnt > 0)
        //                            {
        //                                oIssue.Lines.Add();
        //                                oIssue.Lines.SetCurrentLine(cnt);
        //                            }
        //                            rs.DoQuery("select * from oitw where itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' and whscode='" + rs1.Fields.Item("U_SOWHSCOD").Value.ToString() + "'");
        //                            if (rs.RecordCount > 0)
        //                            {
        //                                System.Data.DataRow dr = dt.Rows.Add();
        //                                dr["OriItemCode"] = rs1.Fields.Item("U_SOITEMCO").Value.ToString();
        //                                dr["LineId"] = rs1.Fields.Item("LineId").Value.ToString();
        //                                dr["OriItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
        //                                rs.DoQuery("select * from oitw where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' and whscode='" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
        //                                if (rs.RecordCount > 0)
        //                                {
        //                                    dr["RevItemCode"] = rs1.Fields.Item("U_ITEMCODE").Value.ToString();
        //                                    dr["RevItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
        //                                    dr["Different"] = double.Parse(dr["OriItemCost"].ToString()) - double.Parse(rs.Fields.Item("AvgPrice").Value.ToString());
        //                                }

        //                            }
        //                            oIssue.Lines.ItemCode = rs1.Fields.Item("U_ITEMCODE").Value.ToString();
        //                            oIssue.Lines.ItemDescription = rs1.Fields.Item("U_ITEMNAME").Value.ToString();
        //                            oIssue.Lines.Quantity = double.Parse(rs1.Fields.Item("U_QUANTITY").Value.ToString());
        //                            oIssue.Lines.WarehouseCode = rs1.Fields.Item("U_WHSCODE").Value.ToString();

        //                            cnt++;
        //                            rs1.MoveNext();
        //                        }
        //                        retcode = oIssue.Add();
        //                        if (retcode != 0)
        //                        {
        //                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                            SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                            return;
        //                        }
        //                        SAP.SBOCompany.GetNewObjectCode(out oIssueKey);

        //                        //rs.DoQuery("select * from ige1 where docentry =" + oIssueKey);
        //                        cnt = 0;
        //                        rs1.MoveFirst();
        //                        while (!rs1.EoF)
        //                        {
        //                            if (cnt > 0)
        //                            {
        //                                oReceipt.Lines.Add();
        //                                oReceipt.Lines.SetCurrentLine(cnt);
        //                            }
        //                            rs.DoQuery("select * from oitw where itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' and whscode='" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");

        //                            oReceipt.Lines.ItemCode = rs1.Fields.Item("U_SOITEMCO").Value.ToString();
        //                            oReceipt.Lines.ItemDescription = rs1.Fields.Item("U_SOITEMNA").Value.ToString();
        //                            oReceipt.Lines.Quantity = double.Parse(rs1.Fields.Item("U_QUANTITY").Value.ToString());
        //                            oReceipt.Lines.WarehouseCode = rs1.Fields.Item("U_WHSCODE").Value.ToString();
        //                            oReceipt.Lines.UnitPrice = double.Parse(rs.Fields.Item("AVGPrice").Value.ToString());

        //                            cnt++;
        //                            rs1.MoveNext();
        //                        }
        //                        retcode = oReceipt.Add();
        //                        if (retcode != 0)
        //                        {
        //                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                            SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                            return;
        //                        }
        //                        SAP.SBOCompany.GetNewObjectCode(out oReceiptKey);
        //                        foreach (System.Data.DataRow dr in dt.Rows)
        //                        {
        //                            different += double.Parse(dr["different"].ToString());
        //                        }
        //                        if (baseType == "13")
        //                        {
        //                            if (different != 0)
        //                            {
        //                                rs.MoveFirst();
        //                                oJE.ReferenceDate = date;
        //                                oJE.Memo = "DO Charge Out";



        //                                if (different > 0)
        //                                {
        //                                    oJE.Lines.AccountCode = "150400";
        //                                    oJE.Lines.Debit = different;
        //                                    oJE.Lines.Add();
        //                                    oJE.Lines.SetCurrentLine(1);
        //                                    oJE.Lines.AccountCode = cogsAcct;
        //                                    oJE.Lines.Credit = different;
        //                                }
        //                                else if (different < 0)
        //                                {
        //                                    oJE.Lines.AccountCode = cogsAcct;
        //                                    oJE.Lines.Debit = -different;
        //                                    oJE.Lines.Add();
        //                                    oJE.Lines.SetCurrentLine(1);
        //                                    oJE.Lines.AccountCode = "150400";
        //                                    oJE.Lines.Credit = -different;
        //                                }
        //                                retcode = oJE.Add();
        //                                if (retcode != 0)
        //                                {
        //                                    if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                                    SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                                    //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                                    return;
        //                                }
        //                                SAP.SBOCompany.GetNewObjectCode(out oJEKey);
        //                            }
        //                        }
        //                    }
        //                    rs.DoQuery("select T0.U_TPCODE, T0.U_LORRY, T0.U_DRIVER, T0.U_DRIVERIC,U_AREA, T1.*, T3.U_BASEOBJ as [ObjType], T3.U_BASEENT AS [SODocEntry], T3.U_BASELINE AS [SOBaseLine] " +
        //                        " from  [@FT_CHARGE] T0 inner join [@FT_CHARGE1] T1 on T0.docentry = T1.docentry " +
        //                        " inner join[@FT_TPPLAN1] T2 on T1.U_BASEENT = T2.DocEntry and T1.U_BASELINE = T2.LineId " +
        //                        " inner join[@FT_SPLAN1] T3 on T2.U_BASEENT = T3.DocEntry and T2.U_BASELINE = T3.LineId WHERE T1.DOCENTRY =" + docEntry + " order by T3.U_BASEOBJ ");
        //                    if (rs.RecordCount > 0)
        //                    {
        //                        cnt = 0;
        //                        if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();
        //                        SAPbobsCOM.Documents oDel = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
        //                        oDel.CardCode = cardcode;
        //                        oDel.DocDate = date;
        //                        oDel.UserFields.Fields.Item("U_ChargeNo").Value = docnum;
        //                        oDel.UserFields.Fields.Item("U_Diff").Value = different;

        //                        if (rs.Fields.Item("U_TPCODE").Value.ToString() != "")
        //                        {
        //                            oDel.UserFields.Fields.Item("U_Transporter").Value = rs.Fields.Item("U_TPCODE").Value.ToString();
        //                        }
        //                        if (rs.Fields.Item("U_LORRY").Value.ToString() != "")
        //                        {
        //                            oDel.UserFields.Fields.Item("U_LorryNo").Value = rs.Fields.Item("U_LORRY").Value.ToString();
        //                        }
        //                        if (rs.Fields.Item("U_DRIVER").Value.ToString() != "")
        //                        {
        //                            oDel.UserFields.Fields.Item("U_Driver").Value = rs.Fields.Item("U_DRIVER").Value.ToString();
        //                        }
        //                        if (rs.Fields.Item("U_DRIVERIC").Value.ToString() != "")
        //                        {
        //                            oDel.UserFields.Fields.Item("U_ICNo").Value = rs.Fields.Item("U_DRIVERIC").Value.ToString();
        //                        }
        //                        if (rs.Fields.Item("U_AREA").Value.ToString() != "")
        //                        {
        //                            oDel.UserFields.Fields.Item("U_Area").Value = rs.Fields.Item("U_AREA").Value.ToString();
        //                        }
        //                        rs.MoveFirst();
        //                        while (!rs.EoF)
        //                        {
        //                            if (cnt > 0)
        //                            {
        //                                oDel.Lines.Add();
        //                                oDel.Lines.SetCurrentLine(cnt);
        //                            }
        //                            oDel.Lines.ItemCode = rs.Fields.Item("U_SOITEMCO").Value.ToString();
        //                            oDel.Lines.UserFields.Fields.Item("U_RevisedItemCode").Value = rs.Fields.Item("U_ITEMCODE").Value.ToString();
        //                            oDel.Lines.UserFields.Fields.Item("U_RevisedItemDesc").Value = rs.Fields.Item("U_ITEMNAME").Value.ToString();
        //                            oDel.Lines.Quantity = double.Parse(rs.Fields.Item("U_QUANTITY").Value.ToString());
        //                            oDel.Lines.WarehouseCode = rs.Fields.Item("U_WHSCODE").Value.ToString();
        //                            if (rs.Fields.Item("ObjType").Value.ToString() == "17")
        //                            {
        //                                oDel.Lines.BaseType = 17;
        //                                rs1.DoQuery("select * from rdr1 where docentry = " + rs.Fields.Item("SODocEntry").Value.ToString() + " and linenum = " + rs.Fields.Item("SOBaseLine").Value.ToString());
        //                                oDel.Lines.PriceAfterVAT = double.Parse(rs1.Fields.Item("PriceAfVAT").Value.ToString());
        //                                oDel.Lines.VatGroup = rs1.Fields.Item("VatGroup").Value.ToString();
        //                                oDel.Lines.LineTotal = double.Parse(rs.Fields.Item("U_QUANTITY").Value.ToString()) * double.Parse(rs1.Fields.Item("Price").Value.ToString());
        //                            }
        //                            else if (rs.Fields.Item("ObjType").Value.ToString() == "13")
        //                            {
        //                                oDel.Lines.BaseType = 13;
        //                                rs1.DoQuery("select * from inv1 where docentry = " + rs.Fields.Item("SODocEntry").Value.ToString() + " and linenum = " + rs.Fields.Item("SOBaseLine").Value.ToString());
        //                                oDel.Lines.PriceAfterVAT = double.Parse(rs1.Fields.Item("PriceAfVAT").Value.ToString());
        //                                oDel.Lines.VatGroup = rs1.Fields.Item("VatGroup").Value.ToString();
        //                                oDel.Lines.LineTotal = double.Parse(rs.Fields.Item("U_QUANTITY").Value.ToString()) * double.Parse(rs1.Fields.Item("Price").Value.ToString());
        //                            }
        //                            oDel.Lines.BaseLine = int.Parse(rs.Fields.Item("SOBaseLine").Value.ToString());
        //                            oDel.Lines.BaseEntry = int.Parse(rs.Fields.Item("SODocEntry").Value.ToString());
        //                            cnt++;
        //                            rs.MoveNext();
        //                        }
        //                        retcode = oDel.Add();
        //                        if (retcode != 0)
        //                        {
        //                            string test = SAP.SBOCompany.GetLastErrorDescription();
        //                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                            SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                            return;
        //                        }
        //                        SAP.SBOCompany.GetNewObjectCode(out oDelKey);
        //                    }

        //                    if (retcode != 0)
        //                    {
        //                        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                        SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                        return;
        //                        //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                    }
        //                    rs.DoQuery("Select T0.docnum, T0.docdate,T1.basetype,T1.baseentry, T1.itemcode, sum(T1.quantity) as [quantity], T1.whscode from ODLN T0 inner join DLN1 T1 on T0.docentry = T1.docentry " +
        //                        " where T0.docentry =" + oDelKey + " group by  T0.docnum, T0.docdate,T1.basetype, T1.baseentry,T1.itemcode,T1.whscode ");
        //                    if (rs.RecordCount > 0)
        //                    {
        //                        oDeldocnum = rs.Fields.Item("docnum").Value.ToString();

        //                        if (rs.Fields.Item("basetype").Value.ToString() == "13")
        //                        {
        //                            double total = 0, RInv_Total = 0;
        //                            string baseentry = "", RInv_docnum = "";
        //                            rs.MoveFirst();
        //                            baseentry = rs.Fields.Item("baseentry").Value.ToString();
        //                            if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();
        //                            oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

        //                            date = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
        //                            oJE.TaxDate = date;
        //                            oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                            while (!rs.EoF)
        //                            {
        //                                rs1.DoQuery("select avgprice from oitw where itemcode='" + rs.Fields.Item("itemcode").Value.ToString() + "' and whscode='" + rs.Fields.Item("whscode").Value.ToString() + "'");
        //                                if (rs1.RecordCount > 0)
        //                                {
        //                                    total += double.Parse(rs.Fields.Item("quantity").Value.ToString()) * double.Parse(rs1.Fields.Item("avgprice").Value.ToString());
        //                                }
        //                                rs.MoveNext();
        //                            }

        //                            oJE.Lines.AccountCode = "150400";
        //                            oJE.Lines.Debit = total;
        //                            oJE.Lines.Add();
        //                            oJE.Lines.SetCurrentLine(1);
        //                            oJE.Lines.AccountCode = cogsAcct;
        //                            oJE.Lines.Credit = total;

        //                            retcode = oJE.Add();

        //                            rs.DoQuery("select * from oinv where docentry=" + baseentry);
        //                            if (rs.RecordCount > 0)
        //                            {
        //                                RInv_docnum = rs.Fields.Item("docnum").Value.ToString();

        //                                rs.DoQuery("select sum(debit) as [total] from ojdt T0 inner join JDT1 T1 on T0.transid = T1.transid where U_RInvNo='" + RInv_docnum + "'");
        //                                if (rs.RecordCount > 0)
        //                                {
        //                                    RInv_Total = double.Parse(rs.Fields.Item("total").Value.ToString());
        //                                    different = RInv_Total - total;
        //                                    if (different > 0)
        //                                    {
        //                                        SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

        //                                        oJE1.TaxDate = date;
        //                                        oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;

        //                                        oJE1.Lines.AccountCode = "530900";
        //                                        oJE1.Lines.Debit = different;
        //                                        oJE1.Lines.Add();
        //                                        oJE1.Lines.SetCurrentLine(1);
        //                                        oJE1.Lines.AccountCode = cogsAcct;
        //                                        oJE1.Lines.Credit = different;

        //                                        retcode = oJE1.Add();
        //                                    }
        //                                    else if (different < 0)
        //                                    {
        //                                        SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

        //                                        oJE1.TaxDate = date;
        //                                        oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;

        //                                        oJE1.Lines.AccountCode = cogsAcct;
        //                                        oJE1.Lines.Debit = different;
        //                                        oJE1.Lines.Add();
        //                                        oJE1.Lines.SetCurrentLine(1);
        //                                        oJE1.Lines.AccountCode = "530900";
        //                                        oJE1.Lines.Credit = different;

        //                                        retcode = oJE1.Add();
        //                                    }
        //                                }
        //                            }
        //                            if (retcode != 0)
        //                            {
        //                                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                                return;
        //                            }
        //                        }
        //                        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //                        if (oReceiptKey != "")
        //                        {
        //                            oReceipt = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
        //                            oReceipt.GetByKey(int.Parse(oReceiptKey));
        //                            oReceipt.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                            oReceipt.Update();
        //                        }
        //                        if (oIssueKey != "")
        //                        {
        //                            oIssue = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
        //                            oIssue.GetByKey(int.Parse(oIssueKey));
        //                            oIssue.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                            oIssue.Update();
        //                        }
        //                        if (oJEKey != "")
        //                        {
        //                            oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
        //                            oJE.GetByKey(int.Parse(oJEKey));
        //                            oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                            oJE.Update();
        //                        }
        //                    }

        //                }
        //                //SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
        //                ////oRS.DoQuery("select AutoKey from ONNM where ObjectCode = 'FT_APSA'");
        //                //oRS.DoQuery("select T0.DocNum, T0.U_BookNo, T0.U_SODOCNUM, T1.U_SOENTRY from ONNM inner join [@FT_APSA] T0 on ONNM.AutoKey - 1 = T0.DocEntry inner join [@FT_APSA1] T1 on T0.DocEntry = T1.DocEntry where ONNM.ObjectCode = 'FT_APSA'");
        //                //string DocNum = oRS.Fields.Item(0).Value.ToString().Trim();
        //                //string BookNo = oRS.Fields.Item(1).Value.ToString().Trim();
        //                //string SODOCNUM = oRS.Fields.Item(2).Value.ToString().Trim();
        //                //string SOENTRY = oRS.Fields.Item(3).Value.ToString().Trim();

        //                ////sendMsg("A", DocNum, BookNo, SODOCNUM, SOENTRY);
        //                //string emailmsg = "Base on Booking No " + BookNo + ".";

        //                //EmailClass email = new EmailClass();
        //                //email.EmailMsg = emailmsg;
        //                //email.EmailSubject = "Draft SA No " + DocNum + " from SC No " + SODOCNUM + " Generated.";
        //                //email.ObjType = "FT_APSA";

        //                //email.SendEmail();

        //                break;
        //            //case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //            //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
        //            //    string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
        //            //    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;

        //            //    SAPbouiCOM.DBDataSource ods = null;
        //            //    ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
        //            //    if (ods.GetValue("Status", 0).ToUpper().Trim() == "C")
        //            //        if (ods.GetValue("Canceled", 0).ToUpper().Trim() == "Y")
        //            //            oForm.DataSources.UserDataSources.Item("docstatus").Value = "CANCELLED";
        //            //        else
        //            //            oForm.DataSources.UserDataSources.Item("docstatus").Value = "CLOSED";
        //            //    else if (ods.GetValue("Status", 0).ToUpper().Trim() == "O")
        //            //        oForm.DataSources.UserDataSources.Item("docstatus").Value = "OPEN";

        //            //    ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname1);
        //            //    for (int x = 0; x < ods.Size; x++)
        //            //    {
        //            //        ods.SetValue("U_ORIQTY", x, ods.GetValue("U_QUANTITY", x));
        //            //    }
        //            //    oMatrix.LoadFromDataSource();

        //            //    arrangematrix(oForm, oMatrix, "@" + dsname1);
        //            //    break;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
        //    }
        //}
        public static void processDataEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            //JEHeader jehdr = new JEHeader();
            //JEDetails jedtl;
            //List<JEDetails> JEdtls = new List<JEDetails>();
            string docnum = "";
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        if (!BusinessObjectInfo.ActionSuccess) return;

                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);

                        UserForm_FT_GENSPERR.SalesInvoice(xmlDoc.InnerText);

                        string bpType = "", cogsAcct = "";
                        DateTime date = new DateTime();
                        int retcode = 0;
                        double temp = 0;
                        double invtotaldiff = 0;
                        int currentline = 0;
                        string productgroup = "";
                        string ChargeNo = "";
                        string sql = "";
                        string delno = "";


                        //sql = "select T0.docdate, T0.docnum, T4.U_Type, min(T1.docnum) as [DelNo] " +
                        //    "from oinv T0 inner join " +
                        //    "( " +
                        //    "select oinv.docentry, isnull(odln.U_Diff, 0) as U_Diff, odln.docnum " +
                        //    "from oinv inner join inv1 on oinv.docentry = inv1.docentry " +
                        //    "inner join dln1 on inv1.baseentry = dln1.docentry and inv1.basetype = 15 " +
                        //    "and inv1.BaseLine = dln1.LineNum " +
                        //    "inner join odln  on dln1.docentry = odln.docentry and odln.CANCELED = 'N' " +
                        //    "where oinv.DocEntry = " + xmlDoc.InnerText + " " +
                        //    "group by oinv.docentry, odln.U_Diff, odln.docnum " +
                        //    ") T1 on T0.docentry = T1.docentry " +
                        //    "inner join ocrd T4 on T0.cardcode = T4.cardcode " +
                        //    "where T0.docentry = " + xmlDoc.InnerText + " " +
                        //    "group by T0.docdate, T0.docnum, T4.U_Type";
                        sql = "select T0.docdate, T0.docnum, T4.U_Type, min(T3.docnum) as [DelNo] " +
                            " from oinv T0 inner join inv1 T1 on T0.docentry = T1.docentry " +
                            " inner join dln1 T2 on T1.baseentry = T2.docentry and T1.basetype = 15 " +
                            " and T1.BaseLine = T2.LineNum " +
                            " inner join odln T3 on T2.docentry = T3.docentry " +
                            " inner join ocrd T4 on T0.cardcode = T4.cardcode " +
                            " where T0.CANCELED = 'N' and T0.docentry = " + xmlDoc.InnerText + " " +
                            " group by T0.docdate, T0.docnum, T4.U_Type";
                        rs.DoQuery(sql);
                        if (rs.RecordCount > 0)
                        {
                            rs.MoveFirst();

                            //different = double.Parse(rs.Fields.Item("U_Diff").Value.ToString());

                            //if (different != 0)
                            {
                                docnum = rs.Fields.Item("docnum").Value.ToString();
                                SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                                date = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
                                oJE.ReferenceDate = date;
                                oJE.Memo = "DO Charge Out";
                                oJE.UserFields.Fields.Item("U_ARInvNo").Value = docnum;
                                oJE.UserFields.Fields.Item("U_DelNo").Value = rs.Fields.Item("DelNo").Value.ToString();

                                //jehdr.DocType = "ARIV";
                                //jehdr.RefDate = oJE.ReferenceDate;
                                //jehdr.U_ARInvNo = oJE.UserFields.Fields.Item("U_ARInvNo").Value.ToString();
                                //jehdr.U_DelNo = oJE.UserFields.Fields.Item("U_DelNo").Value.ToString();
                                //jehdr.Memo = oJE.Memo;

                                bpType = rs.Fields.Item("U_Type").Value.ToString();
                                switch (bpType.ToUpper())
                                {
                                    case "LOCAL":
                                        cogsAcct = "500001";
                                        break;
                                    case "OVERSEA":
                                        cogsAcct = "500100";
                                        break;
                                    case "INTER-COMPANY":
                                        cogsAcct = "500200";
                                        break;
                                }


                                rs.DoQuery("select T3.U_ChargeNo, T3.docnum" +
                                    " from oinv T0 inner join inv1 T1 on T0.docentry = T1.docentry " +
                                    " inner join dln1 T2 on T1.baseentry = T2.docentry and T1.basetype = 15 " +
                                    " and T1.BaseLine = T2.LineNum " +
                                    " inner join odln T3 on T2.docentry = T3.docentry " +
                                    " where T0.docentry = " + xmlDoc.InnerText +
                                    " group by T3.U_ChargeNo, T3.docnum");
                                if (rs.RecordCount > 0)
                                {
                                    rs.MoveFirst();

                                    while (!rs.EoF)
                                    {
                                        //productgroup = rs.Fields.Item("U_CostCenter").Value.ToString();
                                        ChargeNo = rs.Fields.Item("U_ChargeNo").Value.ToString();
                                        delno = rs.Fields.Item("docnum").Value.ToString();
                                        //if (string.IsNullOrEmpty(productgroup)) productgroup = "";
                                        if (!string.IsNullOrEmpty(ChargeNo))
                                        {
                                            sql = " select T1.LineId, T5.U_CostCenter, (T1.U_QUANTITY * isnull(T6.AvgPrice,0)) - (T1.U_QUANTITY * isnull(T7.AvgPrice,0)) as total " +
                                                " from [@FT_CHARGE] T0 inner join [@FT_CHARGE1] T1 on T0.docentry = T1.docentry and T1.U_SOITEMCO <> T1.U_ITEMCODE" +
                                                " inner join oitm T4 on T1.U_SOITEMCO = T4.itemcode and isnull(T4.invntitem,'N') = 'Y'" +
                                                " inner join oitb T5 on T4.itmsgrpcod = T5.itmsgrpcod" +
                                                " inner join oitm T8 on T1.U_ITEMCODE = T8.itemcode and isnull(T8.invntitem,'N') = 'Y'" +
                                                " inner join ( select T1.itemcode, avg(T1.stockprice) as AvgPrice from OIGN T0 inner join IGN1 T1 on T0.DocEntry = T1.DocEntry where T0.U_DelNo = " + delno + " group by T1.itemcode ) T6 on T1.U_SOITEMCO = T6.itemcode" +
                                                " inner join ( select T1.itemcode, avg(T1.stockprice) as AvgPrice from OIGE T0 inner join IGE1 T1 on T0.DocEntry = T1.DocEntry where T0.U_DelNo = " + delno + " group by T1.itemcode ) T7 on T1.U_ITEMCODE = T7.itemcode" +
                                                " where T0.DocNum = " + ChargeNo +
                                                " order by T1.LineId";
                                            rs1.DoQuery(sql);
                                            if (rs1.RecordCount > 0)
                                            {
                                                rs1.MoveFirst();
                                                while (!rs1.EoF)
                                                {
                                                    productgroup = rs1.Fields.Item("U_CostCenter").Value.ToString();
                                                    temp = double.Parse(rs1.Fields.Item("total").Value.ToString());
                                                    temp = Math.Round(temp, 2);
                                                    currentline++;
                                                    if (currentline > 1)
                                                    {
                                                        oJE.Lines.Add();
                                                        oJE.Lines.SetCurrentLine(currentline - 1);
                                                    }
                                                    oJE.Lines.AccountCode = cogsAcct;
                                                    if (temp > 0)
                                                        oJE.Lines.Credit = temp;
                                                    else if (temp < 0)
                                                        oJE.Lines.Debit = -temp;
                                                    if (!string.IsNullOrEmpty(productgroup))
                                                        oJE.Lines.CostingCode = productgroup;

                                                    //jedtl = new JEDetails();
                                                    //jedtl.AccountCode = oJE.Lines.AccountCode;
                                                    //jedtl.CostingCode = oJE.Lines.CostingCode;
                                                    //jedtl.Debit = oJE.Lines.Debit;
                                                    //jedtl.Credit = oJE.Lines.Credit;
                                                    //JEdtls.Add(jedtl);

                                                    invtotaldiff = invtotaldiff + temp;

                                                    rs1.MoveNext();
                                                }
                                            }

                                        }
                                        rs.MoveNext();
                                    }

                                    if (invtotaldiff != 0)
                                    {
                                        currentline++;
                                        if (currentline > 1)
                                        {
                                            oJE.Lines.Add();
                                            oJE.Lines.SetCurrentLine(currentline - 1);
                                        }
                                        oJE.Lines.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                        if (invtotaldiff > 0)
                                            oJE.Lines.Debit = invtotaldiff;
                                        else if (invtotaldiff < 0)
                                            oJE.Lines.Credit = -invtotaldiff;

                                        //jedtl = new JEDetails();
                                        //jedtl.AccountCode = oJE.Lines.AccountCode;
                                        //jedtl.Credit = oJE.Lines.Credit;
                                        //jedtl.Debit = oJE.Lines.Debit;
                                        //JEdtls.Add(jedtl);
                                    }

                                }
                                if (invtotaldiff != 0)
                                {
                                    retcode = oJE.Add();
                                    if (retcode != 0)
                                    {
                                        //jehdr.ErrMsg = SAP.SBOCompany.GetLastErrorDescription();
                                        //jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmss");
                                        //ObjectFunctions.ErrorLog(jehdr, JEdtls);
                                        //if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        rs.DoQuery("update [@FT_SPERRLOG] set U_ErrMsg = '" + SAP.SBOCompany.GetLastErrorDescription() + "' where U_ARInvNo = '" + docnum + "'");
                                        SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                        //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    }
                                    rs.DoQuery("update [@FT_SPERRLOG] set Status = 'C', Canceled = 'Y' where U_ARInvNo = '" + docnum + "'");
                                }
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                //jehdr.ErrMsg = "Data Event After " + ex.Message;
                //jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                //ObjectFunctions.ErrorLog(jehdr, JEdtls);
                //if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAPbobsCOM.Recordset rs0 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs0.DoQuery("update [@FT_SPERRLOG] set U_ErrMsg = '" + ex.Message + "' where U_ARInvNo = '" + docnum + "'");
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
