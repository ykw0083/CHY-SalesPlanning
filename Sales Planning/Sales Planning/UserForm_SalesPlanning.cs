using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class UserForm_SalesPlanning
    {
        public static void deletebatch(SAPbouiCOM.Form oForm, int dsrow)
        {
            if (oForm.TypeEx == "FT_CHARGE")
            {
                string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
                SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname1);
                ods.SetValue("U_BQTY", dsrow, "0");
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                oMatrix.LoadFromDataSource();
                string dsnameb = oForm.DataSources.UserDataSources.Item("dsnameb").Value.Trim();
                SAPbouiCOM.DBDataSource oDS2 = oForm.DataSources.DBDataSources.Item("@" + dsnameb);
                string visorder = ods.GetValue("VisOrder", dsrow);
                if (oDS2 != null)
                {
                    int row = oDS2.Size;
                    for (int y = row - 1; y >= 0; y--)
                    {
                        if (oDS2.GetValue("U_BASEVIS", y).Trim() == visorder)
                        {
                            oDS2.RemoveRecord(y);
                        }
                    }
                }
            }

        }
        public static void checkrow(SAPbouiCOM.Form oForm)
        {
            try
            {
                string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                SAPbouiCOM.DBDataSource ods = null;
                ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);


                if (ods.GetValue("Status", 0).ToUpper().Trim() == "C")
                {
                    if (ods.GetValue("Canceled", 0).ToUpper().Trim() == "Y")
                        oForm.DataSources.UserDataSources.Item("docstatus").Value = "CANCELLED";
                    else
                        oForm.DataSources.UserDataSources.Item("docstatus").Value = "CLOSED";
                }
                else if (ods.GetValue("Status", 0).ToUpper().Trim() == "O")
                {
                    if (dsname == "FT_SPLAN" || dsname == "FT_TPPLAN")
                    {
                        string approval = ods.GetValue("U_APP", 0).ToUpper().Trim();
                        if (string.IsNullOrEmpty(approval)) approval = "O";
                        switch (approval)
                        {
                            case "O":
                                oForm.DataSources.UserDataSources.Item("docstatus").Value = "OPEN";
                                if (ods.GetValue("U_RELEASE", 0).ToUpper().Trim() == "Y")
                                    oForm.DataSources.UserDataSources.Item("docstatus").Value = "RELEASED";
                                break;
                            case "Y":
                                oForm.DataSources.UserDataSources.Item("docstatus").Value = "APPROVED";
                                if (ods.GetValue("U_RELEASE", 0).ToUpper().Trim() == "Y")
                                    oForm.DataSources.UserDataSources.Item("docstatus").Value = "RELEASED-A";
                                break;
                            case "N":
                                oForm.DataSources.UserDataSources.Item("docstatus").Value = "REJECTED";
                                if (ods.GetValue("U_RELEASE", 0).ToUpper().Trim() == "Y")
                                    oForm.DataSources.UserDataSources.Item("docstatus").Value = "RELEASED-R";
                                break;
                            case "W":
                                oForm.DataSources.UserDataSources.Item("docstatus").Value = "PENDING";
                                if (ods.GetValue("U_RELEASE", 0).ToUpper().Trim() == "Y")
                                    oForm.DataSources.UserDataSources.Item("docstatus").Value = "RELEASED-P";
                                break;
                        }
                    }
                    else
                    {
                        oForm.DataSources.UserDataSources.Item("docstatus").Value = "OPEN";
                    }

                }
                string status = ods.GetValue("Status", 0).ToUpper().Trim();
                string release = "N";
                string comparecolumn = "";
                switch (dsname)
                {
                    case "FT_SPLAN":
                        comparecolumn = "U_TPQTY";
                        release = ods.GetValue("U_RELEASE", 0).ToUpper().Trim();
                        break;
                    case "FT_TPPLAN":
                        comparecolumn = "U_CMQTY";
                        release = ods.GetValue("U_RELEASE", 0).ToUpper().Trim();
                        break;
                }
                double comparevalue = 0;
                ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname1);
                if (oMatrix.RowCount >= ods.Size)
                {

                    for (int x = 0; x < ods.Size; x++)
                    {
                        ods.SetValue("U_ORIQTY", x, ods.GetValue("U_QUANTITY", x));
                        if (status == "O" && release == "N" && ods.GetValue("U_LSTATUS", x) == "O")
                        {
                            if (comparecolumn != "")
                            {
                                comparevalue = double.Parse(ods.GetValue(comparecolumn, x));
                                if (comparevalue != 0)
                                    oMatrix.CommonSetting.SetRowEditable(x + 1, false);
                                else
                                    oMatrix.CommonSetting.SetRowEditable(x + 1, true);
                            }
                            else
                                oMatrix.CommonSetting.SetRowEditable(x + 1, true);
                        }
                        else
                            oMatrix.CommonSetting.SetRowEditable(x + 1, false);
                    }
                }
                ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dsname1);
                oMatrix.LoadFromDataSource();

                arrangematrix(oForm, oMatrix, "@" + dsname1);

            }
            catch (Exception ex)
            {
                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox("checkrow " + ex.Message, 1, "Ok", "", "");
            }
        }
        public static void AddNew(SAPbouiCOM.Form oForm)
        {
            try
            {

                string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
                int docentry = 0;
                if (!int.TryParse(oForm.DataSources.DBDataSources.Item('@' + dsname).GetValue("DocEntry", 0), out docentry))
                {
                    docentry = 0;
                }
                if (docentry == 0)
                {
                    oForm.DataSources.UserDataSources.Item("docstatus").Value = "OPEN";
                    oForm.DataSources.DBDataSources.Item('@' + dsname).SetValue("U_DOCDATE", 0, DateTime.Today.ToString("yyyyMMdd"));

                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("series").Specific;
                    string defseries = oForm.DataSources.UserDataSources.Item("defseries").Value;
                    oCombo.Select(defseries, SAPbouiCOM.BoSearchKey.psk_Index);
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("AddNew " + ex.Message, 1, "Ok", "", "");
            }

        }
        public static void processDataEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, ref bool BubbleEvent)
        {
            try
            {
                string ds = "";
                string ds1 = "";
                SAPbouiCOM.DBDataSource ods = null;
                SAPbouiCOM.DBDataSource ods1 = null;
                string errMsg = "", limitType = "";
                double different = 0, c_usage = 0, t_limit = 0, c_limit = 0;

                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
/*
                        if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN")
                        {
                            ds = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                            ds1 = oForm.DataSources.UserDataSources.Item("dsname1").ValueEx;
                            ods = oForm.DataSources.DBDataSources.Item("@" + ds);
                            ods1 = oForm.DataSources.DBDataSources.Item("@" + ds1);
                            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            if (ft_Functions.CheckCreditTerm(oForm, ods, ods1, ref errMsg))
                            {
                                if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN")
                                {
                                    if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_RELEASE", 0) == "Y")
                                    {
                                        if (ObjectFunctions.Approval("17"))
                                        {
                                            if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_APP", 0) == "O")
                                            {
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APP", 0, "W");
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));

                                                //SAP.SBOApplication.SetStatusBarMessage("Approval Required. There is/are invoices overdue for this customer", SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            }
                                        }
                                    }
                                }
                            }
                            errMsg = "";
                            int cnt = 0;
                            cnt = ft_Functions.CheckCreditLimit(oForm, ods, ods1, ref errMsg, ref limitType, ref different, ref c_usage, ref t_limit, ref c_limit);
                            if (cnt == -1)
                            {
                                BubbleEvent = false;
                                break;
                            }
                            else if (cnt >= 1)
                            {
                                if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN")
                                {
                                    if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_RELEASE", 0) == "Y")
                                    {
                                        if (ObjectFunctions.Approval("17"))
                                        {
                                            if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_APP", 0) == "O")
                                            {
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APP", 0, "W");
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                                oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));
                                                //SAP.SBOApplication.SetStatusBarMessage("Approval Required. Credit Limit Exceeded " + Environment.NewLine + "Limit Type - " +
                                                //limitType + Environment.NewLine + " Over Limit Amount - RM " + different.ToString("#,###,###,###.00"), SAPbouiCOM.BoMessageTime.bmt_Medium, false);
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        //BubbleEvent = false;
*/
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN")
                        {
                            ds = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                            ds1 = oForm.DataSources.UserDataSources.Item("dsname1").ValueEx;
                            ods = oForm.DataSources.DBDataSources.Item("@" + ds);
                            ods1 = oForm.DataSources.DBDataSources.Item("@" + ds1);
                            bool appreq = false;
                            if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_APP", 0) == "N" ||
                                    oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_APP", 0) == "W")
                            { appreq = true; }

                            if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_RELEASE", 0) == "Y" && appreq)
                            {
                                //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_RELEASE", 0, "N");
                                SAP.SBOApplication.MessageBox("Cannot Release." + Environment.NewLine + "This document approval required.", 1, "Ok", "", "");
                                //BubbleEvent = false;
                            }

                            //SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                            //rs.DoQuery("select count(*) from [@FT_RAPPMSG] where U_objcode = '" + oForm.TypeEx + "'");
                            /*
                            int cnt = 0;
                            cnt = ft_Functions.CheckCreditTerm(oForm, ods, ods1, ref errMsg);
                            if (cnt == -1)
                            {
                                BubbleEvent = false;
                                break;
                            }
                            else if (cnt >= 1)
                            {
                                //if (ObjectFunctions.Approval(oForm.TypeEx))
                                //{
                                if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_RELEASE", 0) == "Y")
                                    {
                                        oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_RELEASE", 0, "N");
                                        SAP.SBOApplication.MessageBox("Cannot Release." + Environment.NewLine + "There is/are invoices overdue for this customer", 1, "Ok", "", "");
                                        BubbleEvent = false;
                                        break;
                                    }
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APP", 0, "W");
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));
                                //}
                            }
                            errMsg = "";
                            cnt = ft_Functions.CheckCreditLimit(oForm, ods, ods1, ref errMsg, ref limitType, ref different, ref c_usage, ref t_limit, ref c_limit);
                            if (cnt == -1)
                            {
                                BubbleEvent = false;
                                break;
                            }
                            else if (cnt >= 1)
                            {
                                //if (ObjectFunctions.Approval(oForm.TypeEx))
                                //{
                                    if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_RELEASE", 0) == "Y")
                                    {
                                        oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_RELEASE", 0, "N");
                                        SAP.SBOApplication.MessageBox("Cannot Release." + Environment.NewLine + "Credit Limit Exceeded " + Environment.NewLine + "Limit Type - " +
                                            limitType + Environment.NewLine + " Over Limit Amount - RM " + different.ToString("#,###,###,###.00"), 1, "Ok", "", "");
                                        BubbleEvent = false;
                                        break;
                                    }
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APP", 0, "W");
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPBY", 0, SAP.SBOCompany.UserName);
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPDATE", 0, DateTime.Today.ToString("yyyyMMdd"));
                                    //oForm.DataSources.DBDataSources.Item("@" + ds).SetValue("U_APPTIME", 0, DateTime.Now.ToString("HHmm"));
                                //}
                            }
                            */
                        }
                        //BubbleEvent = false;
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
                string docnum = "";
                string approval = "";
                string docentry = "";
                string cardcode = "";

                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                        if (!BusinessObjectInfo.ActionSuccess) break;
                        string ds = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                        /*
                        docnum = oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("DocNum", 0);
                        docentry = oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("DocEntry", 0);
                        cardcode = oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_CARDCODE", 0);
                        if (ds == "FT_SPLAN" || ds == "FT_TPPLAN")
                        {
                            approval = oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_APP", 0);
                            if (approval == "W" || approval == "N")
                            {
                                sendAppMsg(ds, "U", docentry, oForm.Title + " " + docnum, "Pending for Approval (U) of Customer[" + cardcode + "]", oForm.Title + " " + docnum);
                            }
                            else
                            {
                                sendMsg(ds, "U", docentry, oForm.Title + " " + docnum, "Document Updated of Customer[" + cardcode + "]", oForm.Title + " " + docnum);
                            }
                            if (oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("U_APP", 0) == "W")
                            {
                                SAP.SBOApplication.MessageBox("Credit Control Approval Required.");
                            }

                        }
                        else
                        {
                            sendMsg(ds, "U", docentry, oForm.Title + " " + docnum, "Document Updated of Customer[" + cardcode + "]", oForm.Title + " " + docnum);
                        }
                        */
                        genDO(oForm, int.Parse(oForm.DataSources.DBDataSources.Item("@" + ds).GetValue("DocEntry", 0)));

                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        if (!BusinessObjectInfo.ActionSuccess) break;

                        string dsn = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                        SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        bool appreq = false;
                        if (dsn == "FT_SPLAN" || dsn == "FT_TPPLAN")
                        {
                            oRS.DoQuery("select top 1 DocNum, DocEntry, U_CARDCODE, U_APP from [@" + dsn + "] order by DocEntry desc");

                            if (oRS.RecordCount > 0)
                            {
                                docnum = oRS.Fields.Item(0).Value.ToString();
                                docentry = oRS.Fields.Item(1).Value.ToString();
                                cardcode = oRS.Fields.Item(2).Value.ToString();
                                approval = oRS.Fields.Item(3).Value.ToString().Trim();

                            }
                            if (approval == "W")
                            {
                                sendAppMsg(dsn, "A", docentry, oForm.Title + " " + docnum, "Pending for Approval of Customer[" + cardcode + "]", oForm.Title + " " + docnum);
                                appreq = true;

                            }
                            else
                            {
                                sendMsg(dsn, "A", docentry, oForm.Title + " " + docnum, "Document Created of Customer[" + cardcode + "]", oForm.Title + " " + docnum);

                            }
                        }
                        else
                        {
                            oRS.DoQuery("select top 1 DocNum, DocEntry, U_CARDCODE from [@" + dsn + "] order by DocEntry desc");

                            if (oRS.RecordCount > 0)
                            {
                                docnum = oRS.Fields.Item(0).Value.ToString();
                                docentry = oRS.Fields.Item(1).Value.ToString();
                                cardcode = oRS.Fields.Item(2).Value.ToString();
                            }
                            sendMsg(dsn, "A", docentry, oForm.Title + " " + docnum, "Document Created of Customer[" + cardcode + "]", oForm.Title + " " + docnum);
                        }
                        genDO(oForm, 0);

                        if (docnum != "")
                        {
                            //if (appreq)
                            //    SAP.SBOApplication.MessageBox("Doc No = " + docnum + System.Environment.NewLine + "Credit Control Approval Required.");
                            //else
                                SAP.SBOApplication.MessageBox("Doc No = " + docnum);
                        }
                        //SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        ////oRS.DoQuery("select AutoKey from ONNM where ObjectCode = 'FT_APSA'");
                        //oRS.DoQuery("select T0.DocNum, T0.U_BookNo, T0.U_SODOCNUM, T1.U_SOENTRY from ONNM inner join [@FT_APSA] T0 on ONNM.AutoKey - 1 = T0.DocEntry inner join [@FT_APSA1] T1 on T0.DocEntry = T1.DocEntry where ONNM.ObjectCode = 'FT_APSA'");
                        //string DocNum = oRS.Fields.Item(0).Value.ToString().Trim();
                        //string BookNo = oRS.Fields.Item(1).Value.ToString().Trim();
                        //string SODOCNUM = oRS.Fields.Item(2).Value.ToString().Trim();
                        //string SOENTRY = oRS.Fields.Item(3).Value.ToString().Trim();

                        ////sendMsg("A", DocNum, BookNo, SODOCNUM, SOENTRY);
                        //string emailmsg = "Base on Booking No " + BookNo + ".";

                        //EmailClass email = new EmailClass();
                        //email.EmailMsg = emailmsg;
                        //email.EmailSubject = "Draft SA No " + DocNum + " from SC No " + SODOCNUM + " Generated.";
                        //email.ObjType = "FT_APSA";

                        //email.SendEmail();

                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
                        checkrow(oForm);
                        break;
                }

            }
            catch (Exception ex)
            {
                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
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

        //                    rs.DoQuery("Select top 1 * from [@FT_CHARGE] T0 ORDER BY T0.DOCENTRY DESC");
        //                    int docEntry = 0, cnt = 0, retcode = 0;
        //                    string oIssueKey = "", oReceiptKey = "", oJEKey="",oDelKey="", cardcode = "", docnum="";
        //                    DateTime date = new DateTime();
        //                    if (rs.RecordCount > 0)
        //                    {
        //                        docEntry = int.Parse(rs.Fields.Item("DocEntry").Value.ToString());
        //                        docnum = rs.Fields.Item("Docnum").Value.ToString();
        //                        date = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
        //                        cardcode = rs.Fields.Item("U_CARDCODE").Value.ToString();
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
        //                            string abc = "select * from oitw where itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' and whscode='" + rs1.Fields.Item("U_SOWHSCOD").Value.ToString() + "'";
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

        //                        rs.DoQuery("select * from ige1 where docentry =" + oIssueKey);
        //                        cnt = 0;
        //                        rs.MoveFirst();
        //                        rs1.MoveFirst();
        //                        while (!rs1.EoF)
        //                        {
        //                            if (cnt > 0)
        //                            {
        //                                oReceipt.Lines.Add();
        //                                oReceipt.Lines.SetCurrentLine(cnt);
        //                            }
        //                            oReceipt.Lines.ItemCode = rs1.Fields.Item("U_SOITEMCO").Value.ToString();
        //                            oReceipt.Lines.ItemDescription = rs1.Fields.Item("U_SOITEMNA").Value.ToString();
        //                            oReceipt.Lines.Quantity = double.Parse(rs1.Fields.Item("U_QUANTITY").Value.ToString());
        //                            oReceipt.Lines.WarehouseCode = rs1.Fields.Item("U_WHSCODE").Value.ToString();
        //                            oReceipt.Lines.UnitPrice = double.Parse(rs.Fields.Item("STOCKPRICE").Value.ToString());

        //                            cnt++;
        //                            rs1.MoveNext();
        //                            rs.MoveNext();
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
        //                        double different = 0;
        //                        foreach (System.Data.DataRow dr in dt.Rows)
        //                        {
        //                            different += double.Parse(dr["different"].ToString());
        //                        }
        //                        if (different != 0)
        //                        {
        //                            rs.MoveFirst();
        //                            oJE.ReferenceDate = date;
        //                            oJE.Memo = "DO Charge Out";
        //                            if (different > 0)
        //                            {
        //                                oJE.Lines.AccountCode = rs.Fields.Item("AcctCode").Value.ToString();
        //                                oJE.Lines.Debit = different;
        //                                oJE.Lines.Add();
        //                                oJE.Lines.SetCurrentLine(1);
        //                                oJE.Lines.AccountCode = "530900";
        //                                oJE.Lines.Credit = different;
        //                            }
        //                            else if(different <0)
        //                            {
        //                                oJE.Lines.AccountCode = "530900";
        //                                oJE.Lines.Debit = -different;
        //                                oJE.Lines.Add();
        //                                oJE.Lines.SetCurrentLine(1);
        //                                oJE.Lines.AccountCode = rs.Fields.Item("AcctCode").Value.ToString();
        //                                oJE.Lines.Credit = -different;
        //                            }
        //                            retcode = oJE.Add();
        //                            if (retcode != 0)
        //                            {
        //                                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //                                SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                                //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
        //                                return;
        //                            }
        //                            SAP.SBOCompany.GetNewObjectCode(out oJEKey);
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
        //                        rs.MoveFirst();

        //                        SAPbobsCOM.Documents oDel = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
        //                        oDel.CardCode = cardcode;
        //                        oDel.DocDate = date;
        //                        oDel.UserFields.Fields.Item("U_ChargeNo").Value = docnum;
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
        //                    string baseentry = "", RInv_docnum = "", oDeldocnum = "";
        //                    double total = 0, RInv_Total = 0;

        //                    if (rs.RecordCount > 0)
        //                    {
        //                        baseentry = rs.Fields.Item("baseentry").Value.ToString();
        //                        oDeldocnum = rs.Fields.Item("docnum").Value.ToString();
        //                        rs.MoveFirst();

        //                        if (rs.Fields.Item("basetype").Value.ToString() == "13")
        //                        {


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
        //                            oJE.Lines.AccountCode = "500300";
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
        //                                    double different = RInv_Total - total;
        //                                    if (different > 0)
        //                                    {
        //                                        SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

        //                                        oJE1.TaxDate = date;
        //                                        oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;

        //                                        oJE1.Lines.AccountCode = "530900";
        //                                        oJE1.Lines.Debit = different;
        //                                        oJE1.Lines.Add();
        //                                        oJE1.Lines.SetCurrentLine(1);
        //                                        oJE1.Lines.AccountCode = "500300";
        //                                        oJE1.Lines.Credit = different;

        //                                        retcode = oJE1.Add();
        //                                    }
        //                                    else if (different < 0)
        //                                    {
        //                                        SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

        //                                        oJE1.TaxDate = date;
        //                                        oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;

        //                                        oJE1.Lines.AccountCode = "500300";
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
        //                            else
        //                            {
        //                                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
        //                                if (oReceiptKey != "")
        //                                {
        //                                    oReceipt = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
        //                                    oReceipt.GetByKey(int.Parse(oReceiptKey));
        //                                    oReceipt.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                                    oReceipt.Update();
        //                                }
        //                                if (oIssueKey != "")
        //                                {
        //                                    oIssue = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
        //                                    oIssue.GetByKey(int.Parse(oIssueKey));
        //                                    oIssue.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                                    oIssue.Update();
        //                                }
        //                                if (oJEKey != "")
        //                                {
        //                                    oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
        //                                    oJE.GetByKey(int.Parse(oJEKey));
        //                                    oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                                    oJE.Update();
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);

        //                            if (oReceiptKey != "")
        //                            {
        //                                oReceipt = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
        //                                oReceipt.GetByKey(int.Parse(oReceiptKey));
        //                                oReceipt.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                                oReceipt.Update();
        //                            }
        //                            if (oIssueKey != "")
        //                            {
        //                                oIssue = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
        //                                oIssue.GetByKey(int.Parse(oIssueKey));
        //                                oIssue.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                                oIssue.Update();
        //                            }
        //                            if (oJEKey != "")
        //                            {
        //                                oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
        //                                oJE.GetByKey(int.Parse(oJEKey));
        //                                oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
        //                                oJE.Update();
        //                            }
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
        //            case SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD:
        //                string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
        //                string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
        //                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;

        //                SAPbouiCOM.DBDataSource ods = null;
        //                ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname);
        //                if (ods.GetValue("Status", 0).ToUpper().Trim() == "C")
        //                    if (ods.GetValue("Canceled", 0).ToUpper().Trim() == "Y")
        //                        oForm.DataSources.UserDataSources.Item("docstatus").Value = "CANCELLED";
        //                    else
        //                        oForm.DataSources.UserDataSources.Item("docstatus").Value = "CLOSED";
        //                else if (ods.GetValue("Status", 0).ToUpper().Trim() == "O")
        //                    oForm.DataSources.UserDataSources.Item("docstatus").Value = "OPEN";

        //                ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname1);
        //                for (int x = 0; x < ods.Size; x++)
        //                {
        //                    ods.SetValue("U_ORIQTY", x, ods.GetValue("U_QUANTITY", x));
        //                }
        //                oMatrix.LoadFromDataSource();

        //                arrangematrix(oForm, oMatrix, "@" + dsname1);
        //                break;
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
        //    }
        //}
        public static void SPLANRemainingOpen(SAPbouiCOM.Form oForm, bool isDuplicate)
        {
            oForm.Freeze(true);

            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
            string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            SAPbouiCOM.DBDataSource ods1 = oForm.DataSources.DBDataSources.Item("@" + dsname1);
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            string comparecolumn = "";
            switch (dsname)
            {
                case "FT_SPLAN":
                    comparecolumn = "U_TPQTY";
                    break;
                case "FT_TPPLAN":
                    comparecolumn = "U_CMQTY";
                    break;
            }
            double comparevalue = 0;
            for (int x = 0; x < ods1.Size; x++)
            {
                ods1.SetValue("U_ORIQTY", x, ods1.GetValue("U_QUANTITY", x));
                if (ods1.GetValue("U_LSTATUS", x) == "O")
                {
                    if (comparecolumn != "")
                    {
                        comparevalue = double.Parse(ods1.GetValue(comparecolumn, x));
                        if (comparevalue != 0)
                            oMatrix.CommonSetting.SetRowEditable(x + 1, false);
                        else
                            oMatrix.CommonSetting.SetRowEditable(x + 1, true);
                    }
                    else
                        oMatrix.CommonSetting.SetRowEditable(x + 1, true);
                }
                else
                    oMatrix.CommonSetting.SetRowEditable(x + 1, false);

            }
            ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dsname1);

            if (dsname == "FT_SPLAN" && dsname1 == "FT_SPLAN1")
            {

                bool found = false;
                for (int x = 0; x < ods1.Size; x++)
                {
                    if (float.Parse(ods1.GetValue("U_TPQTY", x).ToString()) > 0)
                        found = true;
                }
                if (found)
                {
                    oForm.Freeze(false);
                    return;
                }
                int docentry = 0;
                if (!int.TryParse(oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("DocEntry", 0), out docentry))
                    docentry = 0;

                System.Data.DataTable dt = new System.Data.DataTable();
                dt.Columns.Add("BASEOBJ", typeof(string));
                dt.Columns.Add("BASEENT", typeof(string));
                dt.Columns.Add("BASELINE", typeof(string));
                dt.Columns.Add("SOQTY", typeof(float));
                string sql = "";
                string table = "";
                float ropenqty = 0;
                for (int x = 0; x < ods1.Size; x++)
                {
                    found = false;
                    if (dt.Rows.Count > 0)
                    {
                        System.Data.DataRow[] dr = dt.Select(" BASEOBJ ='" + ods1.GetValue("U_BASEOBJ", x) + "' and BASEENT = '" + ods1.GetValue("U_BASEENT", x) + "' and BASELINE = '" + ods1.GetValue("U_BASELINE", x) + "' ");
                        if (dr.Length > 0)
                        {
                            found = true;
                        }
                    }
                    if (!found)
                    {
                        System.Data.DataRow dr = dt.Rows.Add();
                        dr["BASEOBJ"] = ods1.GetValue("U_BASEOBJ", x);
                        dr["BASEENT"] = ods1.GetValue("U_BASEENT", x);
                        dr["BASELINE"] = ods1.GetValue("U_BASELINE", x);

                        table = dr["BASEOBJ"].ToString() == "17" ? "RDR1" : "INV1";
                        if (table == "RDR1")
                        {
                            sql = "select T0.Quantity - isnull(T0.U_SPLANQTY,0) + isnull(T1.U_QUANTITY,0) - isnull(T9.Quantity,0) from RDR1 T0 left join (select U_BASEOBJ, U_BASEENT, U_BASELINE, sum(U_QUANTITY) as U_QUANTITY from [@FT_SPLAN1] inner join [@FT_SPLAN] on [@FT_SPLAN].DocEntry = [@FT_SPLAN1].DocEntry where [@FT_SPLAN].DocEntry = " + docentry.ToString() + " and [@FT_SPLAN].Status = 'O' and U_TPQTY = 0 group by U_BASEOBJ, U_BASEENT, U_BASELINE) T1 on T0.DocEntry = T1.U_BASEENT and T0.LineNum = T1.U_BASELINE and T0.ObjType = T1.U_BASEOBJ ";
                            sql = sql + " left join (select INV1.BaseType, INV1.BaseEntry, INV1.BaseLine, sum(INV1.Quantity) as Quantity from INV1 inner join OINV on INV1.DocEntry = OINV.DocEntry and OINV.isIns = 'Y' where INV1.BaseType = 17 and OINV.CANCELED = 'N' and INV1.BaseEntry = " + dr["BASEENT"].ToString() + " group by INV1.BaseType, INV1.BaseEntry, INV1.BaseLine) T9 on T0.ObjType = T9.BaseType and T0.DocEntry = T9.BaseEntry and T0.LineNum = T9.BaseLine";
                        }
                        else
                            sql = "select T0.Quantity - isnull(T0.U_SPLANQTY,0) + isnull(T1.U_QUANTITY,0) from INV1 T0 left join (select U_BASEOBJ, U_BASEENT, U_BASELINE, sum(U_QUANTITY) as U_QUANTITY from [@FT_SPLAN1] inner join [@FT_SPLAN] on [@FT_SPLAN].DocEntry = [@FT_SPLAN1].DocEntry where [@FT_SPLAN].DocEntry = " + docentry.ToString() + " and [@FT_SPLAN].Status = 'O' and U_TPQTY = 0 group by U_BASEOBJ, U_BASEENT, U_BASELINE) T1 on T0.DocEntry = T1.U_BASEENT and T0.LineNum = T1.U_BASELINE and T0.ObjType = T1.U_BASEOBJ ";

                        sql = sql + " where T0.DocEntry = " + dr["BASEENT"].ToString() + " and T0.LineNum = " + dr["BASELINE"].ToString();
                        rs.DoQuery(sql);
                        if (rs.RecordCount > 0)
                        {
                            ropenqty = float.Parse(rs.Fields.Item(0).Value.ToString());
                            dr["SOQTY"] = ropenqty;
                        }
                        else
                        {
                            dr["SOQTY"] = 0;
                        }
                        //if (float.TryParse(ods1.GetValue("U_SOQTY", x), out ropenqty))
                        //    dr["SOQTY"] = ropenqty;
                        //else
                        //    dr["SOQTY"] = 0;
                    }
                }
                float quantity = 0;
                foreach (System.Data.DataRow dr in dt.Rows)
                {
                    ropenqty = float.Parse(dr["SOQTY"].ToString());
                    found = false;
                    for (int x = 0; x < ods1.Size; x++)
                    {
                        if (ods1.GetValue("U_BASEOBJ", x) == dr["BASEOBJ"].ToString() && ods1.GetValue("U_BASEENT", x) == dr["BASEENT"].ToString() && ods1.GetValue("U_BASELINE", x) == dr["BASELINE"].ToString())
                        {
                            if (found)
                            {
                                ods1.SetValue("U_SOQTY", x, ropenqty.ToString());
                                //ods1.SetValue("U_QUANTITY", x, ropenqty.ToString());
                            }
                            else
                            {
                                found = true;
                                ods1.SetValue("U_SOQTY", x, ropenqty.ToString());
                                if (isDuplicate)
                                    ods1.SetValue("U_QUANTITY", x, ropenqty.ToString());
                            }
                            if (!float.TryParse(ods1.GetValue("U_QUANTITY", x), out quantity))
                                quantity = 0;
                            ropenqty = ropenqty - quantity;
                        }
                    }
                }
                oMatrix.LoadFromDataSource();

            }
            oForm.Freeze(false);
        }
        public static void SetLTotal(SAPbouiCOM.Form oForm)
        {
            oForm.Freeze(true);
            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
            string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
            if ((dsname == "FT_CHARGE" && dsname1 == "FT_CHARGE1") || (dsname == "FT_TPPLAN" && dsname1 == "FT_TPPLAN1"))
            {

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                oMatrix.FlushToDataSource();
                SAPbouiCOM.DBDataSource ods1 = oForm.DataSources.DBDataSources.Item("@" + dsname1);
                float linetotal = 0;
                float doctotal = 0;

                for (int x = 0; x < ods1.Size; x++)
                {
                    linetotal = float.Parse(ods1.GetValue("U_QUANTITY", x).ToString()) * float.Parse(ods1.GetValue("U_WEIGHT", x).ToString()) / 1000;
                    ods1.SetValue("U_LTOTAL", x, linetotal.ToString());
                    //if (x + 1 == row)
                    //{
                    //    ods.SetValue("U_LTOTAL", x, linetotal.ToString());
                    //    if (oMatrix.Columns.Item("U_LTOTAL").Editable == false)
                    //    {
                    //        oMatrix.Columns.Item("U_LTOTAL").Editable = true;
                    //        try
                    //        {
                    //            ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_LTOTAL").Cells.Item(row).Specific).String = linetotal.ToString();
                    //        }
                    //        catch { }
                    //        oMatrix.Columns.Item("U_LTOTAL").Editable = false;
                    //    }
                    //    else
                    //        ((SAPbouiCOM.EditText)oMatrix.Columns.Item("U_LTOTAL").Cells.Item(row).Specific).String = linetotal.ToString();
                    //}
                    doctotal += linetotal;
                }
                oMatrix.LoadFromDataSource();

                SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                if (ods.GetValue("U_AREA", 0).ToString().Trim() == "")
                {
                    doctotal = 0;
                }
                else if (ods.GetValue("U_AREA", 0).ToString().Substring(0, 1) == "X")
                {
                    doctotal = float.Parse(ods.GetValue("U_APRICE", 0));
                }
                else
                {
                    linetotal = float.Parse(ods.GetValue("U_APRICE", 0).ToString());
                    doctotal = doctotal * linetotal;
                }
                ods.SetValue("U_DOCTOTAL", 0, doctotal.ToString());

                //oMatrix.Columns.Item(coluid).Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            oForm.Freeze(false);
        }
        public static void SetRemarks(SAPbouiCOM.Form oForm)
        {
            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
            string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
            if (dsname == "FT_TPPLAN" || dsname == "FT_SPLAN")
            {
                SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname1);
                if (ods.Size > 0)
                {
                    if (dsname == "FT_TPPLAN")
                    {
                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_ADDRESS", 0).ToString().Trim() == "")
                        {
                            oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("U_ADDRESS", 0, ods.GetValue("U_ADDRESS", 0));
                        }
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_REMARKS", 0).ToString().Trim() == "" && oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_REMARKS", 0).ToString().Trim() == "")
                    {
                        string objtype = ods.GetValue("U_BASEOBJ", 0).ToString();
                        int SPDocEntry = 0;
                        if (int.TryParse(ods.GetValue("U_BASEENT", 0).ToString(), out SPDocEntry))
                        {
                            if (SPDocEntry > 0)
                            {
                                SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                if (dsname == "FT_TPPLAN")
                                    rc.DoQuery("select isnull(U_REMARKS,''), isnull(U_FREMARKS,'') from [@FT_SPLAN] where DocEntry = " + SPDocEntry.ToString());
                                else
                                {
                                    if (objtype == "17" )
                                        rc.DoQuery("select isnull(U_REMARKS,''), isnull(U_FREMARKS,'') from ORDR where DocEntry = " + SPDocEntry.ToString());
                                    else if (objtype == "13")
                                        rc.DoQuery("select isnull(U_REMARKS,''), isnull(U_FREMARKS,'') from OINV where DocEntry = " + SPDocEntry.ToString());
                                }
                                if (rc.RecordCount > 0)
                                {
                                    oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("U_REMARKS", 0, rc.Fields.Item(0).Value.ToString());
                                    oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("U_FREMARKS", 0, rc.Fields.Item(1).Value.ToString());
                                }
                            }
                        }
                    }

                }
            }

        }
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        if (pVal.ItemUID == "grid1")
                        {
                            if (pVal.ColUID == "U_WHSCODE")
                            { }
                            else if (pVal.ColUID == "U_ITEMCODE")
                            {
                                string dsname = oForm.DataSources.UserDataSources.Item("dsname1").ValueEx;
                                SAPbouiCOM.Matrix oMatrx = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                oMatrx.FlushToDataSource();
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_ALLITEM", pVal.Row - 1).ToString() == "Y")
                                { }
                                else
                                {
                                    ObjectFunctions.CopyFromGridColumn(oForm, "grid1", dsname, pVal.ColUID, pVal.Row - 1);
                                    BubbleEvent = false;
                                }
                            }
                            else
                            {
                                string dsname = oForm.DataSources.UserDataSources.Item("dsname1").ValueEx;
                                SAPbouiCOM.Matrix oMatrx = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                oMatrx.FlushToDataSource();
                                ObjectFunctions.CopyFromGridColumn(oForm, "grid1", dsname, pVal.ColUID, pVal.Row - 1);
                                BubbleEvent = false;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            if (pVal.Row >= 0)
                            {
                                string dsname = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                                {
                                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status = 'OPEN' allowed to proceed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "folder1")
                            oForm.PaneLevel = 1;
                        else if (pVal.ItemUID == "folder2")
                            oForm.PaneLevel = 2;
                        else if (pVal.ItemUID == "1")
                        {
                            //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            //{
                            //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                            //    string docdate = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_DOCDATE", 0).ToString();
                            //    docdate = docdate.Trim();
                            //    if (docdate.Length == 8)
                            //    {
                            //        docdate = docdate.Substring(0, 4) + "-" + docdate.Substring(4, 2) + "-" + docdate.Substring(6, 2);
                            //        string sql = "select NNM1.Series from NNM1 inner join OFPR on NNM1.Indicator = OFPR.Indicator where NNM1.ObjectCode = '" + dsname + "' and '" + docdate + "' between OFPR.F_RefDate and OFPR.T_RefDate";
                            //        rs.DoQuery(sql);
                            //        if (rs.RecordCount > 0)
                            //        {
                            //            oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("Series", 0, rs.Fields.Item(0).Value.ToString());
                            //        }

                            //    }
                            //}
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                oMatrix.FlushToDataSource();
                                string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.ToString();
                                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("@" + dsname1);

                                
                                #region Check Sufficient Qty
                                /*
                                System.Data.DataTable dt = new System.Data.DataTable();
                                dt.Columns.Add("ItemCode", typeof(string));
                                dt.Columns.Add("Quantity", typeof(double));
                                dt.Columns.Add("WhsCode", typeof(string));
                                bool found = false;
                                double tempdo = 0;
                                double temprpo = 0;
                                string sql = "";
                                for (int x = 0; x < ds.Size; x++)
                                {
                                    found = false;
                                    if (dt.Rows.Count > 0)
                                    {
                                        System.Data.DataRow[] dr = dt.Select(" ItemCode ='" + ds.GetValue("U_ITEMCODE", x) + "' and whscode = '" + ds.GetValue("U_WHSCODE", x) + "'");
                                        if (dr.Length > 0)
                                        {
                                            if (dsname1 == "FT_SPLAN1")
                                            {
                                                temprpo = double.Parse(ds.GetValue("U_RBPOQTY", x));
                                                tempdo = 0;
                                                if (temprpo == 0)
                                                {
                                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                                    {
                                                        sql = "select sum(isnull(T1.U_CMQTY,0)) from [@FT_TPPLAN] T0 inner join [@FT_TPPLAN1] T1 on T0.DocEntry = T1.DocEntry and T0.Status = 'C' and T0.Canceled = 'N'";
                                                        sql = sql + " where T1.U_BASEOBJ = 'FT_SPLAN' and T1.U_BASEENT = " + ds.GetValue("DocEntry", x).ToString() + " and T1.U_BASELINE = " + ds.GetValue("LineId", x).ToString();
                                                        rs.DoQuery(sql);
                                                        if (rs.RecordCount > 0)
                                                        {
                                                            if (rs.Fields.Item(0).Value != null)
                                                            {
                                                                tempdo = double.Parse(rs.Fields.Item(0).Value.ToString());
                                                            }
                                                        }
                                                    }
                                                }
                                                if (tempdo > temprpo)
                                                    dr[0]["Quantity"] = double.Parse(dr[0]["Quantity"].ToString()) + double.Parse(ds.GetValue("U_QUANTITY", x).ToString()) - tempdo;
                                                else
                                                    dr[0]["Quantity"] = double.Parse(dr[0]["Quantity"].ToString()) + double.Parse(ds.GetValue("U_QUANTITY", x).ToString()) - temprpo;

                                            }
                                            else
                                            {
                                                dr[0]["Quantity"] = double.Parse(dr[0]["Quantity"].ToString()) + double.Parse(ds.GetValue("U_QUANTITY", x).ToString());
                                            }
                                            found = true;
                                        }
                                    }
                                    if (!found)
                                    {
                                        System.Data.DataRow dr = dt.Rows.Add();
                                        dr["ItemCode"] = ds.GetValue("U_ITEMCODE", x);
                                        if (dsname1 == "FT_SPLAN1")
                                        {
                                            temprpo = double.Parse(ds.GetValue("U_RBPOQTY", x));
                                            tempdo = 0;
                                            if (temprpo == 0)
                                            {
                                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                                {
                                                    sql = "select sum(isnull(T1.U_CMQTY,0)) from [@FT_TPPLAN] T0 inner join [@FT_TPPLAN1] T1 on T0.DocEntry = T1.DocEntry and T0.Status = 'C' and T0.Canceled = 'N'";
                                                    sql = sql + " where T1.U_BASEOBJ = 'FT_SPLAN' and T1.U_BASEENT = " + ds.GetValue("DocEntry", x).ToString() + " and T1.U_BASELINE = " + ds.GetValue("LineId", x).ToString();
                                                    rs.DoQuery(sql);
                                                    if (rs.RecordCount > 0)
                                                    {
                                                        if (rs.Fields.Item(0).Value != null)
                                                        {
                                                            tempdo = double.Parse(rs.Fields.Item(0).Value.ToString());
                                                        }
                                                    }
                                                }
                                            }
                                            if (tempdo > temprpo)
                                                dr["Quantity"] = double.Parse(ds.GetValue("U_Quantity", x)) - tempdo;
                                            else
                                                dr["Quantity"] = double.Parse(ds.GetValue("U_Quantity", x)) - temprpo;

                                        }
                                        else
                                            dr["Quantity"] = double.Parse(ds.GetValue("U_Quantity", x));

                                        dr["WhsCode"] = ds.GetValue("U_WHSCODE", x);
                                    }
                                }
                                for (int i = 0; i < dt.Rows.Count; i++)
                                {
                                    rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                                    rs.DoQuery("select OITW.* from OITW inner join OITM on OITW.itemcode = OITM.itemcode and OITM.InvntItem = 'Y' where OITW.itemcode ='" + dt.Rows[i]["ItemCode"].ToString() + "' and OITW.whscode='" + dt.Rows[i]["WhsCode"].ToString() + "'");
                                    if (rs.RecordCount > 0)
                                    {
                                        if (double.Parse(rs.Fields.Item("OnHand").Value.ToString()) < double.Parse(dt.Rows[i]["Quantity"].ToString()))
                                        {
                                            SAP.SBOApplication.MessageBox("Insufficient quantity for item code - " + dt.Rows[i]["ItemCode"].ToString(), 1, "Ok", "", "");
                                            BubbleEvent = false;
                                            return;
                                            //throw new Exception("Insufficient quantity for item code - " + dt.Rows[i]["ItemCode"].ToString());
                                        }
                                    }
                                }
                                */
                                #endregion
                                
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
                                if (validateBatches(oForm))
                                { }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }

                                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                                {
                                    //string DocNum = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocNum", 0).Trim();
                                    //string BookNo = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).Trim();
                                    //string SODOCNUM = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_SODOCNUM", 0).Trim();
                                    //string SOENTRY = oForm.DataSources.DBDataSources.Item("@FT_APSA1").GetValue("U_SOENTRY", 0).Trim();
                                    //sendMsg("U", DocNum, BookNo, SODOCNUM, SOENTRY);
                                }

                                string dsname = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                                string status = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0).Trim();
                                if (status == "O")
                                {
                                    if (dsname == "FT_TPPLAN" || dsname == "FT_SPLAN")
                                    {
                                        SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);

                                        if (ods.GetValue("U_RELEASE", 0) == "N")
                                        {
                                            string app = ods.GetValue("U_APP", 0);
                                            if (string.IsNullOrEmpty(app)) app = "O";
                                            if (app == "N")
                                            {
                                                ods.SetValue("U_APP", 0, "W");
                                            }
                                            if (app == "W")
                                            {
                                                SAP.SBOApplication.MessageBox("Document is pending for approval.");
                                            }
                                            else
                                            {
                                                int rtn = SAP.SBOApplication.MessageBox("Do You want to RELEASE this " + oForm.Title + ".", 1, "Yes", "No", "Cancel");
                                                if (rtn == 1)
                                                {
                                                    ods.SetValue("U_RELEASE", 0, "Y");
                                                }
                                                else if (rtn == 2)
                                                {
                                                    ods.SetValue("U_RELEASE", 0, "N");
                                                }
                                                else
                                                {
                                                    BubbleEvent = false;
                                                    return;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                            }
                        }
                        //if (pVal.ItemUID == "cb_gendo")
                        //{

                        //    //if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //    {
                        //        if (oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("Status", 0) != "O")
                        //        {
                        //            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status = 'O' allowed Generate DO.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //            BubbleEvent = false;
                        //        }

                        //    }
                        //    else
                        //    {
                        //        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only OK mode allowed Generate DO.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //        BubbleEvent = false;
                        //    }
                        //}
                        if (pVal.ItemUID == "cb_trans")
                        {
                            SAP.SBOApplication.ActivateMenuItem("3088");
                        }
                        if (pVal.ItemUID == "cb_batch")
                        {
                            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.ToString();
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                                {
                                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document Status is not OPEN.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
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
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT:
                        //if (pVal.ItemUID == "U_DRIVER")
                        //{
                        //    SAPbouiCOM.ComboBox oCBDriver = (SAPbouiCOM.ComboBox)oForm.Items.Item(pVal.ItemUID).Specific;
                        //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                        //    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                        //    string code = oCBDriver.Selected.Value.ToString();
                        //    rs.DoQuery("select U_ICNo, name from [@Driver] where code='" + code + "'");
                        //    if (rs.RecordCount > 0)
                        //    {
                        //        ods.SetValue("U_DRIVERIC", 0, rs.Fields.Item("U_ICNo").Value.ToString());
                        //        ods.SetValue("U_DRIVERNA", 0, rs.Fields.Item("name").Value.ToString());
                        //    }
                        //}

                        //if (pVal.ItemUID == "U_AREA")
                        //{
                        //    SAPbouiCOM.ComboBox oCBDriver = (SAPbouiCOM.ComboBox)oForm.Items.Item(pVal.ItemUID).Specific;
                        //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                        //    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                        //    string code = oCBDriver.Selected.Value.ToString();
                        //    rs.DoQuery("select U_Price_MT from [@TRANSPORT_CHARGES] where code='" + code + "'");
                        //    if (rs.RecordCount > 0)
                        //    {
                        //        ods.SetValue("U_APRICE", 0, rs.Fields.Item("U_Price_MT").Value.ToString());
                        //    }
                        //    SetLTotal(oForm);
                        //}
                        if (pVal.ItemUID == "cb_Copy")
                        {
                            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                            string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
                            bool cancon = true;
                            if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                                cancon = false;

                            if (dsname == "FT_TPPLAN" || dsname == "FT_SPLAN")
                            {
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_RELEASE", 0) == "Y"
                                    || oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APP", 0) == "W")
                                {
                                    cancon = false;
                                }
                            }
                            if (cancon)
                            {
                                ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).FlushToDataSource();
                                SAPbouiCOM.ButtonCombo oButtonCombo = (SAPbouiCOM.ButtonCombo)oForm.Items.Item(pVal.ItemUID).Specific;
                                string code = oButtonCombo.Selected.Value.ToString();
                                ObjectFunctions.CopyFromGrid1(oForm, oButtonCombo);
                            }
                            else
                            {
                                SAP.SBOApplication.SetStatusBarMessage("Document Status is not available for Copy From.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                            }

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        //if (pVal.ItemUID == "grid1")
                        //{
                        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        //    {
                        //        if (pVal.Row >= 0)
                        //        {
                        //            int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("@FT_APSA1").GetValue("DocEntry", pVal.Row - 1).ToString());
                        //            int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("@FT_APSA1").GetValue("LineId", pVal.Row - 1).ToString());
                        //            string bookno = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).Trim();
                        //            if (pVal.ColUID == "U_BookNo")
                        //            {
                        //                InitForm.CONM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APSAC", "U_CONNO,U_BookNo", bookno);
                        //            }
                        //        }
                        //    }
                        //}
                        if (pVal.ItemUID == "U_ADDRESS" || pVal.ItemUID == "U_REMARKS" || pVal.ItemUID == "U_FREMARKS")
                        {
                            SAPbouiCOM.Item oItem = oForm.Items.Item(pVal.ItemUID);
                            if (oItem.Type == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            {
                                string dsname = oForm.DataSources.UserDataSources.Item("dsname").ValueEx;
                                int docentry = 0;
                                if (!int.TryParse(oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("DocEntry", 0).ToString(), out docentry))
                                    docentry = 0;
                                int row = 0;
                                string value = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(pVal.ItemUID, 0).ToString();
                                CFLInit.CFLText(oForm.UniqueID, docentry, row, "@" + dsname, pVal.ItemUID, "", value);
                            }
                        }
                        else if (pVal.ItemUID == "grid1")
                        {
                            if ((pVal.ColUID == "U_LREMARKS" || pVal.ColUID == "U_EREMARKS") && pVal.Row > 0)
                            {
                                string dsname = oForm.DataSources.UserDataSources.Item("dsname1").ValueEx;
                                int docentry = 0;
                                if (!int.TryParse(oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("DocEntry", pVal.Row - 1).ToString(), out docentry))
                                    docentry = 0;
                                int row = pVal.Row - 1;
                                string value = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(pVal.ColUID, row).ToString();
                                CFLInit.CFLText(oForm.UniqueID, docentry, row, "@" + dsname, pVal.ColUID, "grid1", value);
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST:
                        SAPbouiCOM.IChooseFromListEvent oCFLEvento = (SAPbouiCOM.IChooseFromListEvent)pVal;
                        SAPbouiCOM.DataTable oDataTable = oCFLEvento.SelectedObjects;
                        if (oDataTable != null && oDataTable.Rows.Count > 0)
                        {
                            string val = oDataTable.GetValue(0, 0).ToString();
                            bool changed = false;

                            if (oCFLEvento.SelectedObjects != null)
                            {
                                if (pVal.ItemUID == "U_CARDCODE")
                                {

                                    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                                    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                                    changed = true;
                                    //ods.SetValue(pVal.ItemUID, 0, val);
                                    //ods.SetValue("U_CARDNAME", 0, oDataTable.GetValue(1, 0).ToString());
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific).String = val;
                                    }
                                    catch { }
                                    try
                                    {
                                        ods.SetValue("U_CARDNAME", 0, oDataTable.GetValue(1, 0).ToString());
                                    }
                                    catch { }
                                }
                                else if (pVal.ItemUID == "U_WHSCODE")
                                {
                                    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                                    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                                    changed = true;
                                    //ods.SetValue(pVal.ItemUID, 0, val);
                                    //ods.SetValue("U_CARDNAME", 0, oDataTable.GetValue(1, 0).ToString());
                                    try
                                    {
                                        ((SAPbouiCOM.EditText)oForm.Items.Item(pVal.ItemUID).Specific).String = val;
                                    }
                                    catch (Exception ex)
                                    {
                                        //FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                                        //SAPbouiCOM.Form oCFL = SAP.SBOApplication.Forms.Item(oCFLEvento.FormUID);
                                        //oCFL.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                    }
                                }
                                else if (pVal.ItemUID == "grid1")
                                {
                                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                                    string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
                                    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname1);
                                    
                                    if (pVal.ColUID == "U_WHSCODE")
                                    {
                                        changed = true;
                                        //ods.SetValue(pVal.ColUID, pVal.Row - 1, val);
                                        deletebatch(oForm, pVal.Row - 1);
                                        try
                                        {
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String = val.Trim();
                                        }
                                        catch (Exception ex)
                                        {
                                            //SAPbouiCOM.Form oCFL = SAP.SBOApplication.Forms.Item(oCFLEvento.FormUID);
                                            //oCFL.Items.Item("2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        }
                                        oMatrix.FlushToDataSource();
                                    }
                                    else if (pVal.ColUID == "U_BIN")
                                    {
                                        changed = true;
                                        //ods.SetValue(pVal.ColUID, pVal.Row - 1, val);
                                        deletebatch(oForm, pVal.Row - 1);
                                        try
                                        {
                                            ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String = val;
                                        }
                                        catch { }
                                        oMatrix.FlushToDataSource();
                                    }
                                    else if (pVal.ColUID == "U_ITEMCODE")
                                    {
                                        changed = true;
                                        ods.SetValue(pVal.ColUID, pVal.Row - 1, val);
                                        ods.SetValue("U_ITEMNAME", pVal.Row - 1, oDataTable.GetValue(1, 0).ToString());
                                        deletebatch(oForm, pVal.Row - 1);
                                        //try
                                        //{
                                        //    ((SAPbouiCOM.EditText)oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Specific).String = val;
                                        //}
                                        //catch
                                        //{
                                        //    //FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                                        //}
                                        //oMatrix.FlushToDataSource();
                                        //ods.SetValue("U_ITEMNAME", pVal.Row - 1, oDataTable.GetValue(1, 0).ToString());
                                        oMatrix.LoadFromDataSource();
                                    }
                                    if (changed)
                                    {
                                        //oMatrix.LoadFromDataSource();
                                        oMatrix.Columns.Item(pVal.ColUID).Cells.Item(pVal.Row).Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                                        //oMatrix.FlushToDataSource();
                                    }
                                }
                                if (changed && oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE:
                        if (oForm.TypeEx == "FT_SPLAN")
                        {
                            oForm.EnableMenu("1284", true);//cancel
                            oForm.EnableMenu("1286", false);//close
                            oForm.EnableMenu("1287", false);//duplicate
                            oForm.EnableMenu("1285", true);//restore
                        }
                        else if (oForm.TypeEx == "FT_TPPLAN")
                        {
                            oForm.EnableMenu("1284", true);//cancel
                            oForm.EnableMenu("1286", false);//close
                            oForm.EnableMenu("1287", false);//duplicate
                            oForm.EnableMenu("1285", true);//restore
                        }
                        else if (oForm.TypeEx == "FT_CHARGE")
                        {
                            oForm.EnableMenu("1284", true);//cancel
                            oForm.EnableMenu("1286", false);//close
                            oForm.EnableMenu("1287", false);//duplicate
                            oForm.EnableMenu("1285", false);//restore
                        }
                        else
                        {
                            oForm.EnableMenu("1284", false);//cancel
                            oForm.EnableMenu("1286", false);//close
                            oForm.EnableMenu("1287", false);//duplicate
                            oForm.EnableMenu("1285", false);//restore
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_DEACTIVATE:
                        oForm.EnableMenu("1284", false);//cancel
                        oForm.EnableMenu("1286", false);//close
                        oForm.EnableMenu("1285", false);//restore
                        oForm.EnableMenu("1287", false);//duplicate
                        break;
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "cb_gendo")
                        {

                            //int docentry = genDO(oForm);
                            //if (docentry < 0)
                            //{
                            //    return;
                            //}

                            //SAPbobsCOM.Documents oDoc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                            //if (!oDoc.GetByKey(docentry))
                            //{
                            //    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Cannot find Deliery Order from DocEntry " + docentry.ToString(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            //    return;
                            //}

                            //int docnum = oDoc.DocNum;

                            //string sadocentry = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocEntry", 0).Trim();
                            //string sono = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_SODOCNUM", 0).Trim();
                            //string bookno = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("U_BookNo", 0).Trim();
                            //string sadocnum = oForm.DataSources.DBDataSources.Item("@FT_APSA").GetValue("DocNum", 0).Trim();

                            //string emailmsg = "Base on" + Environment.NewLine + "Booking No " + bookno + Environment.NewLine + "Shipping Advise No " + sadocnum + ".";

                            //EmailClass email = new EmailClass();
                            //email.EmailMsg = emailmsg;
                            //email.EmailSubject = "Delivery Order No " + docnum.ToString() + " from SC No " + sono + " Generated.";
                            //email.ObjType = "FT_APSA";

                            //email.SendEmail();
                            ////oForm.DataSources.DBDataSources.Item("@FT_APSA").SetValue("Status", 0, "C");
                            ////oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;

                            ////oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText("DO Generated please check from Delivery Order Screen.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                            //oForm.Close();
                            //return;
                        }
                        else if (pVal.ItemUID == "cb_batch")
                        {

                            InitForm.batchno(oForm.UniqueID, "@FT_CHARGE1", "@FT_CHARGE2");
                            oForm.Freeze(true);
                        }
                        else if (pVal.ItemUID == "1")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("DocNum", 0) == "0")
                                { }
                                else
                                {
                                    SAP.SBOApplication.ActivateMenuItem("1289");
                                    SAP.SBOApplication.ActivateMenuItem("1288");
                                }
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.ActionSuccess)
                            {
                                AddNew(oForm);
                            }
                        }
                        else if (pVal.ItemUID == "2")
                        { }
                        else if (pVal.ItemUID == "cb_Copy")
                        {

                        }
                        else
                        {
                            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                            string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.Trim();
                            if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) == "O")
                            {
                                if (pVal.ItemUID.Substring(0, 2) == "C_")
                                    ObjectFunctions.CopyFromItem(oForm, pVal.ItemUID);
                            }

                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_GOT_FOCUS:
                        if (pVal.ItemUID == "grid1")
                        {
                            SAP.currentMatrixRow = pVal.Row;
                            if (pVal.Row > 0)
                            {
                            }
                            else
                            {
                                oForm.EnableMenu("1293", false);//delete row
                                if (oForm.TypeEx == "FT_SPLAN")
                                {
                                    oForm.EnableMenu("1299", false);//close row
                                    oForm.EnableMenu("1312", false);//open row
                                }
                                //oForm.EnableMenu("1294", false);//duplicate row
                            }
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_LOST_FOCUS:
                        if (pVal.ItemUID == "grid1")
                        {
                            oForm.EnableMenu("1293", false);//delete row
                            //oForm.EnableMenu("1294", false);//duplicate row
                            if (pVal.ColUID == "U_QUANTITY")
                            {
                                SetLTotal(oForm);
                            }
                            else if (pVal.ColUID == "U_WEIGHT")
                            {
                                SetLTotal(oForm);
                            }
                            break;
                        }
                        else if (pVal.ItemUID == "U_DOCDATE")
                        {
                            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                            string docdate = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue(pVal.ItemUID, 0).ToString();
                            docdate = docdate.Trim();
                            if (docdate.Length == 8)
                            {
                                docdate = docdate.Substring(0, 4) + "-" + docdate.Substring(4, 2) + "-" + docdate.Substring(6, 2);
                                string sql = "select NNM1.Series from NNM1 inner join OFPR on NNM1.Indicator = OFPR.Indicator where NNM1.ObjectCode = '" + dsname + "' and '" + docdate + "' between OFPR.F_RefDate and OFPR.T_RefDate";
                                rs.DoQuery(sql);
                                if (rs.RecordCount > 0)
                                {
                                    oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("Series", 0, rs.Fields.Item(0).Value.ToString());
                                }

                            }
                        }
                        else if (pVal.ItemUID == "U_APRICE")
                        {
                            SetLTotal(oForm);
                        }
                        //else if (pVal.ItemUID == "U_AREA")
                        //{
                        //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                        //    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                        //    string tpcode = ods.GetValue("U_TPCODE", 0);
                        //    string area = ods.GetValue("U_AREA", 0);
                        //    rs.DoQuery("select isnull(T0.U_Price_MT,0) as U_APRICE from [@TRANSPORTER_AREA_D] T0 inner join [@TRANSPORTER_AREA] T1 on T0.Code = T1.Code and T0.U_Price_MT > 0 and isnull(T0.U_EXPIRED,'N') = 'N' inner join [@TRANSPORTER] T2 on T1.U_Transporter = T2.Code and isnull(T2.U_Blacklist,'N') = 'N' where T2.Code = '" + tpcode + "' and T0.U_Area = '" + area + "'");
                        //    if (rs.RecordCount > 0)
                        //    {
                        //        ods.SetValue("U_APRICE", 0, rs.Fields.Item("U_APRICE").Value.ToString());
                        //    }



                        //}
                        //else if (pVal.ItemUID == "U_DRIVERIC")
                        //{
                        //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                        //    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                        //    string ic = ods.GetValue(pVal.ItemUID, 0);
                        //    rs.DoQuery("select code, name from [@Driver] where U_ICNo='" + ic + "'");
                        //    if (rs.RecordCount > 0)
                        //    {
                        //        ods.SetValue("U_DRIVER", 0, rs.Fields.Item("code").Value.ToString());
                        //        ods.SetValue("U_DRIVERNA", 0, rs.Fields.Item("name").Value.ToString());
                        //    }

                        //}
                        //else if (pVal.ItemUID == "U_DRIVER")
                        //{
                        //    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.Trim();
                        //    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname);
                        //    string code = ods.GetValue(pVal.ItemUID, 0);
                        //    rs.DoQuery("select U_ICNo, name from [@Driver] where code='" + code + "'");
                        //    if (rs.RecordCount > 0)
                        //    {
                        //        ods.SetValue("U_DRIVERIC", 0, rs.Fields.Item("U_ICNo").Value.ToString());
                        //        ods.SetValue("U_DRIVERNA", 0, rs.Fields.Item("name").Value.ToString());
                        //    }
                        //}
                        break;
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "grid1")
                        {
                            if (pVal.Row > 0 && pVal.ColUID == "VisOrder")
                            {
                                SAPbouiCOM.Matrix oGrid = (SAPbouiCOM.Matrix)oForm.Items.Item(pVal.ItemUID).Specific;
                                oGrid.SelectRow(pVal.Row, true, false);

                                oForm.EnableMenu("1293", true);//delete row
                                if (oForm.TypeEx == "FT_SPLAN")
                                {
                                    oForm.EnableMenu("1299", true);//close row
                                    oForm.EnableMenu("1312", false);//open row
                                }
                                //oForm.EnableMenu("1294", true);//duplicate row
                            }
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
                string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value;
                SAPbouiCOM.DBDataSource oDS = oForm.DataSources.DBDataSources.Item("@" + dsname);
                SAPbouiCOM.Matrix oMatrix = null;
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //long docentry = 0;
                //string sql = "";
                //int value = 0;
                int rtn = 0;
                int docentry = 0;
                string sql = "";

                if (pVal.MenuUID == "1286") // Close doc
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only OK mode allowed Cancel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status Open allowed Close.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                        return;
                    }
                    //if (oForm.TypeEx == "FT_TPPLAN")
                    //{
                    //    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APP", 0) == "W")
                    //    {
                    //        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("This Document is Pending for Approval.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    //        BubbleEvent = false;
                    //        return;
                    //    }
                    //}
                    docentry = int.Parse(oDS.GetValue("DocEntry", 0));
                    sql = "";
                    if (dsname == "FT_SPLAN")
                        sql = "select count(*) from [@FT_TPPLAN1] T1 inner join [@FT_TPPLAN] T0 on T0.DocEntry = T1.DocEntry where T0.Status = 'O' and T1.BASEENT = " + docentry;

                    if (sql != "")
                    {
                        rs.DoQuery(sql);
                        rs.MoveFirst();
                        if (rs.Fields.Item(0).Value == null)
                            rtn = 0;
                        else
                            rtn = int.Parse(rs.Fields.Item(0).Value.ToString());

                        if (rtn > 0)
                        {
                            if (dsname == "FT_SPLAN")
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Outstanding Transport Planning found, Closing is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            BubbleEvent = false;
                            return;
                        }

                    }

                    rtn = SAP.SBOApplication.MessageBox("Close Document is irreversable." + System.Environment.NewLine + "Are you sure you want to continue?", 1, "OK", "Cancel");
                    if (rtn == 2)
                    {
                        BubbleEvent = false;
                        return;
                    }
                }
                else if (pVal.MenuUID == "1284") // cancel doc
                {
                    if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only OK mode allowed Cancel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Only Status Open allowed Cancel.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN")
                    {
                        //if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_RELEASE", 0) == "Y")
                        //{
                        //    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("This Document is RELEASED.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //    BubbleEvent = false;
                        //    return;
                        //}
                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APP", 0) == "W")
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("This Document is Pending for Approval.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                            return;
                        }
                    }
                    docentry = int.Parse(oDS.GetValue("DocEntry", 0));
                    sql = "";
                    if (dsname == "FT_SPLAN")
                        sql = "select sum(isnull(U_TPQTY,0)) from [@FT_SPLAN1] where DocEntry = " + docentry;
                    else if (dsname == "FT_TPPLAN")
                        sql = "select sum(isnull(U_CMQTY,0)) from [@FT_TPPLAN1] where DocEntry = " + docentry;
                    if (sql != "")
                    {
                        rs.DoQuery(sql);
                        rs.MoveFirst();
                        int value = int.Parse(rs.Fields.Item(0).Value.ToString());
                        if (value > 0)
                        {
                            if (dsname == "FT_SPLAN")
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Transport Planning found, Cancel is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            else if (dsname == "FT_TPPLAN")
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Charge Module Order found, Cancel is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);

                            BubbleEvent = false;
                            return;
                        }
                    }

                    rtn = SAP.SBOApplication.MessageBox("Cancel Document is irreversable." + System.Environment.NewLine + "Are you sure you want to continue?", 1, "OK", "Cancel");
                    if (rtn == 2)
                    {
                        BubbleEvent = false;
                        return;
                    }

                }
                else if (pVal.MenuUID == "1299") // close row
                {
                    BubbleEvent = false;
                    bool deleterow = false;
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                    {
                        FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document Status is not OPEN.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_RELEASE", 0) == "Y")
                    {
                        FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document is Released.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APP", 0) == "W")
                    {
                        FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document is pending for approval.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false;
                        return;
                    }
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                    string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value;
                    string checkcol = "";
                    if (dsname == "FT_SPLAN") checkcol = "U_TPQTY";
                    if (dsname == "FT_TPPLAN") checkcol = "U_CMQTY";
                    double qty = 0;
                    for (int x = 1; x <= oMatrix.RowCount; x++)
                    {
                        if (oMatrix.IsRowSelected(x))
                        {
                            qty = 0;
                            if (checkcol != "")
                            {
                                if (double.TryParse(oForm.DataSources.DBDataSources.Item("@" + dsname1).GetValue(checkcol, x - 1), out qty))
                                {

                                }
                                else
                                    qty = 0;
                            }
                            //oMatrix.Columns.Item("U_ITEMCODE").Cells.Item(x).Specific;

                            if (qty == 0)
                            {
                                FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Please delete the Item instead of Close.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                            else
                            {
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname1).GetValue("U_LSTATUS", x - 1) == "O")
                                {
                                    deleterow = true;
                                    oForm.DataSources.DBDataSources.Item("@" + dsname1).SetValue("U_LSTATUS", x - 1, "C");
                                    oForm.DataSources.DBDataSources.Item("@" + dsname1).SetValue("U_QUANTITY", x - 1, qty.ToString());
                                    oMatrix.LoadFromDataSource();
                                    oMatrix.CommonSetting.SetRowEditable(x, false);
                                }
                            }
                        }
                    }

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deleterow) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                else if (pVal.MenuUID == "1312") // open row
                {
                    bool deleterow = false;
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                    {
                        FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document Status is not OPEN.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_RELEASE", 0) == "Y")
                    {
                        FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document is Released.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false;
                        return;
                    }
                    if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APP", 0) == "W")
                    {
                        FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document is pending for approval.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                        BubbleEvent = false;
                        return;
                    }
                    oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                    string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value;
                    string checkcol = "";
                    if (dsname == "FT_SPLAN") checkcol = "U_TPQTY";
                    if (dsname == "FT_TPPLAN") checkcol = "U_CMQTY";
                    double qty = 0;
                    for (int x = 1; x <= oMatrix.RowCount; x++)
                    {
                        if (oMatrix.IsRowSelected(x))
                        {
                            qty = 0;
                            if (checkcol != "")
                            {
                                if (double.TryParse(oForm.DataSources.DBDataSources.Item("@" + dsname1).GetValue(checkcol, x - 1), out qty))
                                {

                                }
                                else
                                    qty = 0;
                            }
                            //oMatrix.Columns.Item("U_ITEMCODE").Cells.Item(x).Specific;

                            if (qty == 0)
                            {
                                FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Please delete the Item instead of Close.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            }
                            else
                            {
                                if (oForm.DataSources.DBDataSources.Item("@" + dsname1).GetValue("U_LSTATUS", x - 1) == "C")
                                {
                                    deleterow = true;
                                    oForm.DataSources.DBDataSources.Item("@" + dsname1).SetValue("U_LSTATUS", x - 1, "O");
                                    //oForm.DataSources.DBDataSources.Item("@" + dsname1).SetValue("U_QUANTITY", x - 1, qty.ToString());
                                    oMatrix.LoadFromDataSource();
                                    //oMatrix.CommonSetting.SetRowEditable(x, false);
                                }
                            }
                        }
                    }
                    BubbleEvent = false;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deleterow) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                else if (pVal.MenuUID == "1293") // delete row
                {
                    bool deleterow = false;
                    //if (oForm.ActiveItem == "grid1")
                    {
                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0) != "O")
                        {
                            FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document Status is not OPEN.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                            return;
                        }
                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_RELEASE", 0) == "Y")
                        {
                            FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document is Released.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                            return;
                        }
                        if (oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APP", 0) == "W")
                        {
                            FT_ADDON.SAP.SBOApplication.SetStatusBarMessage("Document is pending for approval.", SAPbouiCOM.BoMessageTime.bmt_Short, true);
                            BubbleEvent = false;
                            return;
                        }

                        oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                        //if (oMatrix.RowCount == 1)
                        //{
                        //    BubbleEvent = false;
                        //}
                        //else
                        //{
                        string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value;
                        string dsnameb = oForm.DataSources.UserDataSources.Item("dsnameb").Value;

                        if (dsname == "FT_TPPLAN")
                        {
                            if (oMatrix.RowCount == 1)
                            {
                                oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("U_ADDRESS", 0, "");
                                oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("U_REMARKS", 0, "");
                                oForm.DataSources.DBDataSources.Item("@" + dsname).SetValue("U_FREMARKS", 0, "");
                            }
                        }
                        double qty = 0;
                        string checkcol = "";
                        if (dsname == "FT_SPLAN") checkcol = "U_TPQTY";
                        if (dsname == "FT_TPPLAN") checkcol = "U_CMQTY";
                        for (int x = 1; x <= oMatrix.RowCount; x++)
                        {
                            if (oMatrix.IsRowSelected(x))
                            {
                                qty = 0;
                                if (checkcol != "")
                                {
                                    if (double.TryParse(oForm.DataSources.DBDataSources.Item("@" + dsname1).GetValue(checkcol, x - 1), out qty))
                                    {

                                    }
                                    else
                                        qty = 0;
                                }
                                //oMatrix.Columns.Item("U_ITEMCODE").Cells.Item(x).Specific;
                                if (qty > 0)
                                {
                                    BubbleEvent = false;
                                    return;
                                }
                                else
                                {
                                    oMatrix.DeleteRow(x);
                                    deleterow = true;
                                    arrangematrix(oForm, oMatrix, "@" + dsname1);
                                    SPLANRemainingOpen(oForm, false);

                                    if (dsnameb.Trim() != "")
                                    {
                                        oDS = oForm.DataSources.DBDataSources.Item("@" + dsnameb);
                                        int row = oDS.Size;
                                        for (int y = row - 1; y >= 0; y--)
                                        {
                                            oDS.RemoveRecord(oDS.Size - 1);
                                        }

                                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Delete Row will reset Batch information. Please reassign Batch No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                                    }
                                }
                            }
                        }


                        //}
                    }
                    BubbleEvent = false;

                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && deleterow) oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
                }
                else if (pVal.MenuUID == "1285") // restore
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        if (oDS.GetValue("Status", 0) == "O")
                        {
                            if (oDS.GetValue("U_RELEASE", 0) == "N")
                            {
                                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document is Not RELEASE, Restore is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                BubbleEvent = false;
                            }
                            else
                            {
                                if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN")
                                {
                                    if (oDS.GetValue("U_APP", 0) == "W")
                                    {
                                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document is pending For approval, Restore is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        BubbleEvent = false;
                                    }
                                    else
                                    {
                                        //oDS.SetValue("U_APP", 0, "O");
                                        oDS.SetValue("U_RELEASE", 0, "N");

                                        docentry = int.Parse(oDS.GetValue("DocEntry", 0));
                                        //sql = "update [@" + oForm.TypeEx + "] set U_APP = 'O', U_RELEASE = 'N' where docentry = " + docentry;
                                        sql = "update [@" + oForm.TypeEx + "] set U_RELEASE = 'N' where docentry = " + docentry;
                                        rs.DoQuery(sql);

                                        SAP.SBOApplication.ActivateMenuItem("1289");
                                        SAP.SBOApplication.ActivateMenuItem("1288");
                                        BubbleEvent = false;
                                    }
                                }
                                else if (oForm.TypeEx == "FT_SPLAN")
                                {
                                    oDS.SetValue("U_RELEASE", 0, "N");

                                    docentry = int.Parse(oDS.GetValue("DocEntry", 0));
                                    sql = "update [@" + oForm.TypeEx + "] set U_RELEASE = 'N' where docentry = " + docentry;
                                    rs.DoQuery(sql);
                                    sql = "update [@" + oForm.TypeEx + "1] set U_LSTATUS = 'O' where U_LSTATUS = 'C' and docentry = " + docentry;
                                    rs.DoQuery(sql);

                                    SAP.SBOApplication.ActivateMenuItem("1289");
                                    SAP.SBOApplication.ActivateMenuItem("1288");
                                    BubbleEvent = false;
                                }
                            }
                        }
                        else
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document Closed, Restore is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document Modified, Restore is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                    BubbleEvent = false;
                }
                else if (pVal.MenuUID == "1294") // duplicate row
                {
                    if (oForm.DataSources.UserDataSources.Item("dsname").ValueEx == "FT_CHARGE")
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Duplicate Row in Charge Module is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                    else
                    {
                        if (oDS.GetValue("Status", 0) == "O")
                        {

                        }
                        else
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document is not Open. Duplicate Row is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                    }
                }
                else if (pVal.MenuUID == "1287") // duplicate
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                    {
                        if (oDS.GetValue("Status", 0) == "C")
                        {
                            string docentrys = oForm.DataSources.DBDataSources.Item("@" + oForm.DataSources.UserDataSources.Item("dsname").ValueEx).GetValue("DocEntry", 0);
                            SAP.SBOApplication.ActivateMenuItem("1282");
                            duplicateDoc(oForm, docentrys);
                            BubbleEvent = false;
                        }
                        else
                        {
                            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document is not Close/Cancel. Duplicate is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            BubbleEvent = false;
                        }
                    }
                    else
                    {
                        FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Document must be OK mode. Duplicate is not allowed.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        BubbleEvent = false;
                    }
                }
                else if (pVal.MenuUID == "1282") // New
                {

                    //string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").ValueEx;
                    //oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                    //for (int x = 1; x <= oMatrix.RowCount; x++)
                    //{
                    //    oMatrix.CommonSetting.SetRowEditable(x, true);
                    //}
                    //ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dsname1);
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
                    //((SAPbouiCOM.EditText)oForm.Items.Item("sodocnum").Specific).Value = "";
                    //((SAPbouiCOM.EditText)oForm.Items.Item("dodocnum").Specific).Value = "";
                    //((SAPbouiCOM.EditText)oForm.Items.Item("status").Specific).Value = "";

                }
                else if (pVal.MenuUID == "1284") // cancel doc
                {
                }
                else if (pVal.MenuUID == "1282") // add doc
                {
                    AddNew(oForm);
                }
                else if (pVal.MenuUID == "1294") // duplicate row
                {
                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                    string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value;
                    string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value;
                    SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("@" + dsname1);

                    if (dsname == "FT_SPLAN")
                        ods.SetValue("U_TPQTY", ods.Size - 1, "0");
                    else if (dsname == "FT_TPPLAN")
                        ods.SetValue("U_CMQTY", ods.Size - 1, "0");

                    arrangematrix(oForm, oMatrix, "@" + oForm.DataSources.UserDataSources.Item("dsname1").ValueEx.Trim());
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_UPDATE_MODE;
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
            SAPbouiCOM.UserDataSource uds = oForm.DataSources.UserDataSources.Item("dsname");
            string dsname = uds.ValueEx;
            uds = oForm.DataSources.UserDataSources.Item("dsname1");
            string dsname1 = uds.ValueEx;
            uds = oForm.DataSources.UserDataSources.Item("dsnameb");
            string dsnameb = uds.ValueEx;
            string linkto = "";
            string soinvnitem = "";
            string invnitem = "";
            string temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("Status", 0).ToString();
            if (temp != "O")
            {
                SAP.SBOApplication.StatusBar.SetText("Cannot proceed, Document is not OPEN.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_CARDCODE", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                linkto = oForm.Items.Item("U_CARDCODE").LinkTo;
                SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_DOCDATE", 0).ToString();
            if (temp == null || temp.Trim() == "")
            {
                linkto = oForm.Items.Item("U_DOCDATE").LinkTo;
                SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            if (dsname != "FT_SPLAN")
            {
                temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_AREA", 0).ToString();
                if (temp == null || temp.Trim() == "")
                {
                    linkto = oForm.Items.Item("U_AREA").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
            }
            if (dsname == "FT_TPPLAN")
            {
                temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_WHSCODE", 0).ToString();
                if (temp == null || temp.Trim() == "")
                {
                    linkto = oForm.Items.Item("U_WHSCODE").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                temp = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_REAPRICE", 0).ToString();
                if (temp == null || temp.Trim() == "")
                {
                    linkto = oForm.Items.Item("U_REAPRICE").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                double price = 0;
                string sprice = "";
                sprice = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_APRICE", 0).ToString();
                if (sprice == null)
                {
                    linkto = oForm.Items.Item("U_APRICE").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(sprice, out price))
                {
                    linkto = oForm.Items.Item("U_APRICE").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                double doctotal = 0;
                string sdoctotal = "";
                sdoctotal = oForm.DataSources.DBDataSources.Item("@" + dsname).GetValue("U_DOCTOTAL", 0).ToString();
                if (sdoctotal == null)
                {
                    linkto = oForm.Items.Item("U_DOCTOTAL").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(sdoctotal, out doctotal))
                {
                    linkto = oForm.Items.Item("U_DOCTOTAL").LinkTo;
                    SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (temp == "Y")
                {
                    if (price == 0)
                    {
                        linkto = oForm.Items.Item("U_APRICE").LinkTo;
                        SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is required.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    if (doctotal == 0)
                    {
                        linkto = oForm.Items.Item("U_DOCTOTAL").LinkTo;
                        SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is required.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }
                else
                {
                    if (price != 0)
                    {
                        linkto = oForm.Items.Item("U_APRICE").LinkTo;
                        SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is not required.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    if (doctotal != 0)
                    {
                        linkto = oForm.Items.Item("U_DOCTOTAL").LinkTo;
                        SAP.SBOApplication.StatusBar.SetText(((SAPbouiCOM.StaticText)oForm.Items.Item(linkto).Specific).Caption + " is not required.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }
            }

            SAPbouiCOM.DBDataSource ods = null;
            SAPbouiCOM.Matrix oMatrix = null;

            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname1);
            SAPbouiCOM.DBDataSource oDS2 = null;
            if (dsnameb.Trim() != "")
                oDS2 = oForm.DataSources.DBDataSources.Item("@" + dsnameb);

            oMatrix.FlushToDataSource();
            double qty = 0;
            double brpoqty = 0;
            double openqty = 0;
            bool found = false;
            bool wrongbatch = false;
            string whscode = "";
            string itemcode = "";
            string visorder = "";
            string columnname = "";
            int cnt = 0;
            SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            for (int x = 0; x < ods.Size; x++)
            {
                visorder = ods.GetValue("VisOrder", x).Trim();

                columnname = "U_SODOCNUM";
                temp = ods.GetValue(columnname, x);
                if (temp == null || temp.Trim() == "")
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (int.Parse(temp) <= 0)
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " is missing.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                columnname = "U_QUANTITY";
                temp = ods.GetValue(columnname, x);
                if (temp == null)
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (!double.TryParse(temp, out qty))
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (qty == 0)
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is cannot be zero.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
                if (dsname == "FT_SPLAN")
                {
                    columnname = "U_RBPOQTY";
                    temp = ods.GetValue(columnname, x);
                    if (temp == null)
                    {
                        SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    if (!double.TryParse(temp, out brpoqty))
                    {
                        SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    if (brpoqty > 0)
                    {
                        if (brpoqty != qty)
                        {
                            SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " Reserve and Reserve Blanket PO is not tally.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                    }

                    columnname = "U_SOQTY";
                    temp = ods.GetValue(columnname, x);
                    if (temp == null)
                    {
                        SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is null.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    if (!double.TryParse(temp, out openqty))
                    {
                        SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    if (openqty == 0)
                    {
                        SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is cannot be zero.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                    columnname = "U_QUANTITY";
                    if (openqty > 0)
                    {
                        if (qty > openqty)
                        {
                            SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " Reserve cannot over SO Open QTY.(1)", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        else if (qty < 0)
                        {
                            SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " Reserve cannot less than zero QTY.(1)", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                    }
                    else if (openqty < 0)
                    {
                        if (qty < openqty)
                        {
                            SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " Reserve cannot less than SO Open QTY.(2)", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                        else if (qty > 0)
                        {
                            SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " Reserve cannot over zero QTY.(1)", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;
                        }
                    }
                }
                else
                    openqty = double.MaxValue;

                whscode = ods.GetValue("U_WHSCODE", x).Trim();
                itemcode = ods.GetValue("U_ITEMCODE", x).Trim();

                if (itemcode != "")
                    cnt++;

                //if (openqty == 0 && qty == 0)
                //{ }
                //else if (openqty < qty)
                //{
                //    SAP.SBOApplication.StatusBar.SetText("Open Quantity for #" + visorder + " is not enough.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //    return false;
                //}
                //if (qty == 0)
                //{ }
                //if (qty < 0)
                //{
                //    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item("U_ITEMCODE").TitleObject.Caption + " for #" + visorder + " is less than 0.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                //    return false;
                //}
                //else if (qty > 0)
                invnitem = "";
                soinvnitem = "";

                if (whscode == "")
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item("U_WHSCODE").TitleObject.Caption + " for #" + visorder + " is emtpy.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                rc.DoQuery("select isnull(InvntItem,'N') from OITM where ItemCode = '" + itemcode + "'");
                if (rc.RecordCount > 0)
                {
                    invnitem = rc.Fields.Item(0).Value.ToString();
                }
                rc.DoQuery("select isnull(InvntItem,'N') from OITM where ItemCode = '" + ods.GetValue("U_SOITEMCO", x).Trim() + "'");
                if (rc.RecordCount > 0)
                {
                    soinvnitem = rc.Fields.Item(0).Value.ToString();
                }
                if (invnitem.ToUpper() != soinvnitem.ToUpper())
                {
                    SAP.SBOApplication.StatusBar.SetText("Stock Item and Non-stock Item cannot interchange.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                if (invnitem.ToUpper() == "Y")
                {
                    if (qty > 0)
                    {
                        found = true;
                        if (oDS2 != null)
                        {
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
                                        oDS2.RemoveRecord(y);
                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        SAP.SBOApplication.StatusBar.SetText("Quantity for #" + visorder + " cannot less than or equal zero.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        return false;
                    }
                }
            }
            if (cnt == 0)
            {
                SAP.SBOApplication.StatusBar.SetText("No Item found.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                return false;
            }

            bool ManBtchNum = false;
            bool BinActivat = false;
            if (wrongbatch)
            {
                rc.DoQuery("select 1 from OITM where ManBtchNum = 'Y' and ItemCode = '" + itemcode + "'");
                if (rc.RecordCount > 0)
                    ManBtchNum = true;
                rc.DoQuery("select 1 from OWHS where BinActivat = 'Y' and WhsCode = '" + whscode + "'");
                if (rc.RecordCount > 0)
                    BinActivat = true;

                if (BinActivat || ManBtchNum)
                {
                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Batch No or Bin information Missing. Please assign.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }
            }

            //if (!found)
            //{
            //    SAP.SBOApplication.StatusBar.SetText("Cannot proceed, no item found.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            //    return false;
            //}

            oMatrix.LoadFromDataSource();

            return true;
        }
        private static bool validateBatches(SAPbouiCOM.Form oForm)
        {
            SAPbouiCOM.UserDataSource uds = oForm.DataSources.UserDataSources.Item("dsname");
            string dsname = uds.ValueEx;
            uds = oForm.DataSources.UserDataSources.Item("dsname1");
            string dsname1 = uds.ValueEx;
            uds = oForm.DataSources.UserDataSources.Item("dsnameb");
            string dsnameb = uds.ValueEx;

            if (dsnameb == null || dsnameb.Trim() == "")
                return true;

            SAPbouiCOM.Matrix oMatrix = oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;

            SAPbouiCOM.DBDataSource ods = null;
            SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

            ods = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Item("@" + dsname1);
            SAPbouiCOM.DBDataSource oDS2 = oForm.DataSources.DBDataSources.Item("@" + dsnameb);

            decimal qty = 0;
            decimal batchqty = 0;
            decimal tempqty = 0;
            string whscode = "";
            string itemcode = "";
            string visorder = "";
            string columnname = "";

            bool ManBtchNum = false;
            bool BinActivat = false;

            for (int x = 0; x < ods.Size; x++)
            {
                visorder = ods.GetValue("VisOrder", x).Trim();

                columnname = "U_QUANTITY";
                if (!decimal.TryParse(ods.GetValue(columnname, x), out qty))
                {
                    SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " is invalid.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                whscode = ods.GetValue("U_WHSCODE", x).Trim();
                itemcode = ods.GetValue("U_ITEMCODE", x).Trim();


                if (qty == 0)
                { }
                else
                {
                    oRS.DoQuery("select 1 from OITM where ManBtchNum = 'Y' and ItemCode = '" + itemcode + "'");
                    if (oRS.RecordCount > 0)
                        ManBtchNum = true;
                    oRS.DoQuery("select 1 from OWHS where BinActivat = 'Y' and WhsCode = '" + whscode + "'");
                    if (oRS.RecordCount > 0)
                        BinActivat = true;

                    oRS.DoQuery("select 1 from OITM where InvntItem = 'Y' and ItemCode = '" + itemcode + "'");
                    if (oRS.RecordCount > 0)
                    {
                        if (qty > 0)
                        {
                            if (ManBtchNum || BinActivat)
                            {
                                batchqty = 0;
                                for (int y = 0; y < oDS2.Size; y++)
                                {
                                    if (oDS2.GetValue("U_BASEVIS", y).Trim() == visorder)
                                    {
                                        if (oDS2.GetValue("U_WHSCODE", y).Trim() == whscode && oDS2.GetValue("U_ITEMCODE", y).Trim() == itemcode)
                                        {
                                            if (!decimal.TryParse(oDS2.GetValue("U_QUANTITY", y), out tempqty))
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
                                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("#" + visorder + " Quantity and Batch Quantity is not tally. Please assign Batch No.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return false;
                                }

                            }

                        }
                        else
                        {
                            SAP.SBOApplication.StatusBar.SetText(oMatrix.Columns.Item(columnname).TitleObject.Caption + " for #" + visorder + " cannot less than zero (Batch or Bin).", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return false;

                        }
                    }
                }

            }

            return true;
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
        private static bool genDO (SAPbouiCOM.Form oForm, int docEntry)
        {
            try
            {
                bool rtn = false;
                if (oForm.TypeEx == "FT_CHARGE")
                {
                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                    SAPbobsCOM.Recordset rs2 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    int owner = 0;
                    rs.DoQuery("select empID from OHEM where userid = " + SAP.SBOCompany.UserSignature.ToString());
                    if (rs.RecordCount > 0)
                    {
                        if (!int.TryParse(rs.Fields.Item(0).Value.ToString(), out owner))
                            owner = 0;
                    }
                    int cnt = 0, retcode = 0;
                    string oIssueKey = "", oReceiptKey = "", oJEKey = "", oDelKey = "", cardcode = "", docnum = "", baseType = "", oDeldocnum = "", bpType = "", cogsAcct = "", ManBtchNum = "", BinActivat = ""; 
                    DateTime date = new DateTime();
                    double different = 0, quantity =0, total_diff=0;
                    string status = "";
                    int currentline = 0;
                    double linetotal = 0;
                    string productgroup = "";
                    double temp = 0;
                    if (docEntry == 0)
                    {
                        rs.DoQuery("Select top 1 T0.*,T3.U_BASEOBJ as [BASETYPE],T4.U_Type, T0.Status from [@FT_CHARGE] T0  " +
                        " INNER JOIN[@FT_CHARGE1] T1 ON T0.DOCENTRY = T1.DOCENTRY " +
                        " INNER JOIN[@FT_TPPLAN1] T2 ON T1.U_BASEENT = T2.DocEntry AND T1.U_BASEOBJ = T2.Object " +
                        " INNER JOIN[@FT_SPLAN1] T3 ON T2.U_BASEENT = T3.DocEntry AND T2.U_BASEOBJ = T3.Object " +
                        " INNER JOIN OCRD T4 on  T0.U_CARDCODE = T4.CARDCODE ORDER BY T0.DOCENTRY DESC");
                    }
                    else if (docEntry > 0)
                    {
                        rs.DoQuery("Select T0.*,T3.U_BASEOBJ as [BASETYPE],T4.U_Type, T0.Status from [@FT_CHARGE] T0  " +
                        " INNER JOIN[@FT_CHARGE1] T1 ON T0.DOCENTRY = T1.DOCENTRY " +
                        " INNER JOIN[@FT_TPPLAN1] T2 ON T1.U_BASEENT = T2.DocEntry AND T1.U_BASEOBJ = T2.Object " +
                        " INNER JOIN[@FT_SPLAN1] T3 ON T2.U_BASEENT = T3.DocEntry AND T2.U_BASEOBJ = T3.Object " +
                        " INNER JOIN OCRD T4 on  T0.U_CARDCODE = T4.CARDCODE " +
                        " WHERE T0.DOCENTRY = " + docEntry.ToString() );
                    }
                    if (rs.RecordCount > 0)
                    {
                        docEntry = int.Parse(rs.Fields.Item("DocEntry").Value.ToString());
                        docnum = rs.Fields.Item("Docnum").Value.ToString();
                        date = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
                        cardcode = rs.Fields.Item("U_CARDCODE").Value.ToString();
                        baseType = rs.Fields.Item("BASETYPE").Value.ToString();
                        bpType = rs.Fields.Item("U_Type").Value.ToString();
                        status = rs.Fields.Item("Status").Value.ToString();
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
                    }
                    if (status != "O")
                    {
                        SAP.SBOApplication.MessageBox("Document is not OPEN, no Delivery generated.", 1, "Ok", "", "");
                        return rtn;
                    }


                    SAPbobsCOM.Documents oReceipt = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                    SAPbobsCOM.Documents oIssue = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                    SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                    int cntbatch = 0;
                    int cntbin = 0;

                    //rs1.DoQuery("select * from [@FT_CHARGE1] T0 inner join OITM  T1 on T0.U_SOITEMCO = T1.itemcode " +
                    //    " inner join OITM T2  on T0.U_itemcode = T2.itemcode where T0.docentry = " + docEntry + " and T0.U_SOITEMCO <> T0.U_ITEMCODE and isnull(T1.invntitem,'N') = 'Y' and isnull(T2.invntitem,'N') = 'Y' ");
                    rs1.DoQuery("select * from [@FT_CHARGE1] T0 inner join OITM  T1 on T0.U_SOITEMCO = T1.itemcode " +
                        " inner join OITM T2  on T0.U_itemcode = T2.itemcode " +
                        " where T0.docentry = " + docEntry + " and T0.U_SOITEMCO <> T0.U_ITEMCODE and isnull(T1.invntitem,'N') = 'Y' and isnull(T2.invntitem,'N') = 'Y' ");
                    if (rs1.RecordCount > 0)
                    {
                        if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();

                        System.Data.DataTable dt = new System.Data.DataTable();
                        dt.Columns.Add("OriItemCode", typeof(string));
                        dt.Columns.Add("RevItemCode", typeof(string));
                        dt.Columns.Add("LineId", typeof(int));
                        dt.Columns.Add("OriItemCost", typeof(double));
                        dt.Columns.Add("RevItemCost", typeof(double));
                        dt.Columns.Add("Quantity", typeof(double));
                        dt.Columns.Add("Different", typeof(double));
                        dt.Columns.Add("ProductGroup", typeof(string));

                        rs1.MoveFirst();

                        oReceipt.DocDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
                        oReceipt.TaxDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
                        oReceipt.UserFields.Fields.Item("U_IsType").Value = "CHRG_OUT";

                        oIssue.DocDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
                        oIssue.TaxDate = DateTime.Parse(rs.Fields.Item("U_DOCDATE").Value.ToString());
                        oIssue.UserFields.Fields.Item("U_IsType").Value = "CHRG_OUT";
                        while (!rs1.EoF)
                        {
                            if (cnt > 0)
                            {
                                oIssue.Lines.Add();
                                oIssue.Lines.SetCurrentLine(cnt);
                            }
                            rs.DoQuery("select T0.*, isnull(T2.U_CostCenter,'') as productgroup from oitm T0 inner join oitb T2 on T0.itmsgrpcod = T2.itmsgrpcod where T0.itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' ");
                            if (rs.RecordCount > 0)
                            {
                                System.Data.DataRow dr = dt.Rows.Add();
                                dr["OriItemCode"] = rs1.Fields.Item("U_SOITEMCO").Value.ToString();
                                dr["LineId"] = rs1.Fields.Item("LineId").Value.ToString();
                                dr["ProductGroup"] = rs.Fields.Item("productgroup").Value.ToString();
                                dr["Quantity"] = rs1.Fields.Item("U_QUANTITY").Value.ToString();
                                //rs.DoQuery("select * from oitm where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' ");
                                //ykw 20180421
                                //if (rs.Fields.Item("AvgPrice").Value.ToString() == "0")
                                //{
                                //    rs.DoQuery("select * from oitm where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' ");
                                //    if (rs.RecordCount > 0)
                                //    {
                                //        dr["OriItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
                                //    }
                                //}
                                //else
                                //    dr["OriItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
                                //rs.DoQuery("select * from OITM where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' ");
                                //ykw 20180421
                                //ykw 20230316
                                rs.DoQuery("select * from OITW where itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' and whscode = '" + rs1.Fields.Item("U_WHSCODE").Value.ToString()  + "'");
                                if (rs.RecordCount > 0)
                                {
                                    if (rs.Fields.Item("AvgPrice").Value.ToString() == "0")
                                    {
                                        rs.DoQuery("select * from OITW where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' and whscode = '" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                        if (rs.RecordCount > 0)
                                        {
                                            dr["OriItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
                                        }
                                    }
                                    else
                                        dr["OriItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
                                }
                                rs.DoQuery("select * from OITW where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' and whscode = '" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                //ykw 20230316
                                if (rs.RecordCount > 0)
                                {
                                    dr["RevItemCode"] = rs1.Fields.Item("U_ITEMCODE").Value.ToString();
                                    dr["RevItemCost"] = rs.Fields.Item("AvgPrice").Value.ToString();
                                    dr["Different"] = double.Parse(dr["OriItemCost"].ToString()) - double.Parse(rs.Fields.Item("AvgPrice").Value.ToString());
                                }
                            }
                            oIssue.Lines.ItemCode = rs1.Fields.Item("U_ITEMCODE").Value.ToString();
                            oIssue.Lines.ItemDescription = rs1.Fields.Item("U_ITEMNAME").Value.ToString();
                            oIssue.Lines.Quantity = double.Parse(rs1.Fields.Item("U_QUANTITY").Value.ToString());
                            oIssue.Lines.WarehouseCode = rs1.Fields.Item("U_WHSCODE").Value.ToString();

                            rs.DoQuery("select isnull(T2.U_CostCenter,'') as productgroup from oitm T0 inner join oitb T2 on T0.itmsgrpcod = T2.itmsgrpcod where T0.itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' ");
                            if (rs.RecordCount > 0)
                            {
                                oIssue.Lines.CostingCode = rs.Fields.Item("productgroup").Value.ToString();
                            }

                            rs.DoQuery("select ManBtchNum from OITM where itemcode ='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "'");

                            cntbatch = 0;
                            cntbin = 0;

                            if (rs.RecordCount > 0)
                            {
                                ManBtchNum = rs.Fields.Item("ManBtchNum").Value.ToString();
                                if (ManBtchNum == "Y")
                                {
                                    rs.DoQuery("select sum(U_Quantity)  as [Quantity], U_BATCHNUM as [BatchNum] from [@FT_Charge2] where docentry =" + docEntry +
                                        " and U_BASEVIS =" + rs1.Fields.Item("VisOrder").Value.ToString() + " group by U_BATCHNUM ");
                                    while (!rs.EoF)
                                    {
                                        if (cntbatch > 0)
                                        {
                                            oIssue.Lines.BatchNumbers.Add();
                                            oIssue.Lines.BatchNumbers.SetCurrentLine(cntbatch);
                                        }
                                        oIssue.Lines.BatchNumbers.BatchNumber = rs.Fields.Item("BatchNum").Value.ToString();
                                        oIssue.Lines.BatchNumbers.Quantity = double.Parse(rs.Fields.Item("Quantity").Value.ToString());

                                        #region bin
                                        rs2.DoQuery("select BinActivat from OWHS where WhsCode ='" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                        if (rs2.RecordCount > 0)
                                        {
                                            BinActivat = rs2.Fields.Item("BinActivat").Value.ToString();
                                            if (BinActivat == "Y")
                                            {
                                                rs2.DoQuery("select sum(U_Quantity)  as [Quantity], U_BINABS as [Bin] from [@FT_Charge2] where docentry =" + docEntry +
                                                    " and U_BASEVIS =" + rs1.Fields.Item("VisOrder").Value.ToString() + " and U_BATCHNUM = '" + rs.Fields.Item("BatchNum").Value.ToString() + "' group by U_BINABS ");
                                                while (!rs2.EoF)
                                                {
                                                    if (cntbin > 0)
                                                    {
                                                        oIssue.Lines.BinAllocations.Add();
                                                        oIssue.Lines.BinAllocations.SetCurrentLine(cntbin);
                                                    }
                                                    oIssue.Lines.BinAllocations.BinAbsEntry = int.Parse(rs2.Fields.Item("Bin").Value.ToString());
                                                    oIssue.Lines.BinAllocations.Quantity = double.Parse(rs2.Fields.Item("Quantity").Value.ToString());
                                                    oIssue.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = cntbatch;

                                                    rs2.MoveNext();
                                                    cntbin++;
                                                }
                                            }
                                        }
                                        #endregion

                                        rs.MoveNext();
                                        cntbatch++;
                                    }
                                }
                                else
                                {
                                    #region bin
                                    rs2.DoQuery("select BinActivat from OWHS where WhsCode ='" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                    if (rs2.RecordCount > 0)
                                    {
                                        BinActivat = rs2.Fields.Item("BinActivat").Value.ToString();
                                        if (BinActivat == "Y")
                                        {
                                            rs2.DoQuery("select sum(U_Quantity)  as [Quantity], U_BINABS as [Bin] from [@FT_Charge2] where docentry =" + docEntry +
                                                " and U_BASEVIS =" + rs1.Fields.Item("VisOrder").Value.ToString() + " group by U_BINABS ");
                                            while (!rs2.EoF)
                                            {
                                                if (cntbin > 0)
                                                {
                                                    oIssue.Lines.BinAllocations.Add();
                                                    oIssue.Lines.BinAllocations.SetCurrentLine(cntbin);
                                                }
                                                oIssue.Lines.BinAllocations.BinAbsEntry = int.Parse(rs2.Fields.Item("Bin").Value.ToString());
                                                oIssue.Lines.BinAllocations.Quantity = double.Parse(rs2.Fields.Item("Quantity").Value.ToString());
                                                //oIssue.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = cntbatch;

                                                rs2.MoveNext();
                                                cntbin++;
                                            }
                                        }
                                    }
                                    #endregion

                                }
                            }
                            cnt++;
                            rs1.MoveNext();
                        }
                        retcode = oIssue.Add();
                        if (retcode != 0)
                        {
                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            //SAP.SBOApplication.MessageBox("Document is not OPEN, no Delivery generated.", 1, "Ok", "", "");
                            SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return rtn;
                        }
                        SAP.SBOCompany.GetNewObjectCode(out oIssueKey);

                        //rs.DoQuery("select * from ige1 where docentry =" + oIssueKey);
                        cnt = 0;
                        rs1.MoveFirst();
                        while (!rs1.EoF)
                        {
                            if (cnt > 0)
                            {
                                oReceipt.Lines.Add();
                                oReceipt.Lines.SetCurrentLine(cnt);
                            }

                            oReceipt.Lines.ItemCode = rs1.Fields.Item("U_SOITEMCO").Value.ToString();
                            oReceipt.Lines.ItemDescription = rs1.Fields.Item("U_SOITEMNA").Value.ToString();
                            oReceipt.Lines.Quantity = double.Parse(rs1.Fields.Item("U_QUANTITY").Value.ToString());
                            oReceipt.Lines.WarehouseCode = rs1.Fields.Item("U_WHSCODE").Value.ToString();
                            //rs.DoQuery("select * from oitm where itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' ");
                            //ykw 20180421
                            //if (rs.Fields.Item("AvgPrice").Value.ToString() == "0")
                            //{
                            //    rs.DoQuery("select * from oitm where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' ");
                            //    if (rs.RecordCount > 0)
                            //    {
                            //        oReceipt.Lines.UnitPrice = double.Parse(rs.Fields.Item("AvgPrice").Value.ToString());
                            //    }
                            //}
                            //else
                            //    oReceipt.Lines.UnitPrice = double.Parse(rs.Fields.Item("AvgPrice").Value.ToString());
                            //ykw 20180421

                            //ykw 20230316
                            rs.DoQuery("select * from OITW where itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' and whscode = '" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                            if (rs.RecordCount > 0)
                            {
                                if (rs.Fields.Item("AvgPrice").Value.ToString() == "0")
                                {
                                    rs.DoQuery("select * from OITW where itemcode='" + rs1.Fields.Item("U_ITEMCODE").Value.ToString() + "' and whscode = '" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                    if (rs.RecordCount > 0)
                                    {
                                        oReceipt.Lines.UnitPrice = double.Parse(rs.Fields.Item("AvgPrice").Value.ToString());
                                    }
                                }
                                else
                                    oReceipt.Lines.UnitPrice = double.Parse(rs.Fields.Item("AvgPrice").Value.ToString());
                            }
                            //ykw 20230316

                            rs.DoQuery("select isnull(T2.U_CostCenter,'') as productgroup from oitm T0 inner join oitb T2 on T0.itmsgrpcod = T2.itmsgrpcod where T0.itemcode='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "' ");
                            if (rs.RecordCount > 0)
                            {
                                oReceipt.Lines.CostingCode = rs.Fields.Item("productgroup").Value.ToString();
                            }

                            rs.DoQuery("select ManBtchNum from OITM where itemcode ='" + rs1.Fields.Item("U_SOITEMCO").Value.ToString() + "'");

                            cntbatch = 0;
                            cntbin = 0;

                            if (rs.RecordCount > 0)
                            {
                                ManBtchNum = rs.Fields.Item("ManBtchNum").Value.ToString();
                                if (ManBtchNum == "Y")
                                {
                                    rs.DoQuery("select sum(U_Quantity)  as [Quantity], U_BATCHNUM as [BatchNum] from [@FT_Charge2] where docentry =" + docEntry +
                                        " and U_BASEVIS =" + rs1.Fields.Item("VisOrder").Value.ToString() + " group by U_BATCHNUM ");
                                    while (!rs.EoF)
                                    {
                                        if (cntbatch > 0)
                                        {
                                            oReceipt.Lines.BatchNumbers.Add();
                                            oReceipt.Lines.BatchNumbers.SetCurrentLine(cntbatch);
                                        }
                                        oReceipt.Lines.BatchNumbers.BatchNumber = rs.Fields.Item("BatchNum").Value.ToString();
                                        oReceipt.Lines.BatchNumbers.Quantity = double.Parse(rs.Fields.Item("Quantity").Value.ToString());

                                        #region bin
                                        rs2.DoQuery("select BinActivat from OWHS where WhsCode ='" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                        if (rs2.RecordCount > 0)
                                        {
                                            BinActivat = rs2.Fields.Item("BinActivat").Value.ToString();
                                            if (BinActivat == "Y")
                                            {
                                                rs2.DoQuery("select sum(U_Quantity)  as [Quantity], U_BINABS as [Bin] from [@FT_Charge2] where docentry =" + docEntry +
                                                    " and U_BASEVIS =" + rs1.Fields.Item("VisOrder").Value.ToString() + " and U_BATCHNUM = '" + rs.Fields.Item("BatchNum").Value.ToString() + "' group by U_BINABS ");
                                                while (!rs2.EoF)
                                                {
                                                    if (cntbin > 0)
                                                    {
                                                        oReceipt.Lines.BinAllocations.Add();
                                                        oReceipt.Lines.BinAllocations.SetCurrentLine(cntbin);
                                                    }
                                                    oReceipt.Lines.BinAllocations.BinAbsEntry = int.Parse(rs2.Fields.Item("Bin").Value.ToString());
                                                    oReceipt.Lines.BinAllocations.Quantity = double.Parse(rs2.Fields.Item("Quantity").Value.ToString());
                                                    oReceipt.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = cntbatch;

                                                    rs2.MoveNext();
                                                    cntbin++;
                                                }
                                            }
                                        }
                                        #endregion

                                        rs.MoveNext();
                                        cntbatch++;
                                    }
                                }
                                else
                                {
                                    #region bin
                                    rs2.DoQuery("select BinActivat from OWHS where WhsCode ='" + rs1.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                    if (rs2.RecordCount > 0)
                                    {
                                        BinActivat = rs2.Fields.Item("BinActivat").Value.ToString();
                                        if (BinActivat == "Y")
                                        {
                                            rs2.DoQuery("select sum(U_Quantity)  as [Quantity], U_BINABS as [Bin] from [@FT_Charge2] where docentry =" + docEntry +
                                                " and U_BASEVIS =" + rs1.Fields.Item("VisOrder").Value.ToString() + " group by U_BINABS ");
                                            while (!rs2.EoF)
                                            {
                                                if (cntbin > 0)
                                                {
                                                    oReceipt.Lines.BinAllocations.Add();
                                                    oReceipt.Lines.BinAllocations.SetCurrentLine(cntbin);
                                                }
                                                oReceipt.Lines.BinAllocations.BinAbsEntry = int.Parse(rs2.Fields.Item("Bin").Value.ToString());
                                                oReceipt.Lines.BinAllocations.Quantity = double.Parse(rs2.Fields.Item("Quantity").Value.ToString());
                                                //oReceipt.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = cntbatch;

                                                rs2.MoveNext();
                                                cntbin++;
                                            }
                                        }
                                    }
                                    #endregion

                                }
                            }
                            cnt++;
                            rs1.MoveNext();
                        }
                        retcode = oReceipt.Add();
                        if (retcode != 0)
                        {
                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return rtn;
                        }
                        SAP.SBOCompany.GetNewObjectCode(out oReceiptKey);

                        total_diff = 0;

                        foreach (System.Data.DataRow dr in dt.Rows)
                        {
                            different = double.Parse(dr["different"].ToString());
                            quantity = double.Parse(dr["quantity"].ToString());
                            total_diff = total_diff + Math.Round(quantity * different, 2, MidpointRounding.AwayFromZero);
                        }
                        
                        total_diff = double.Parse(total_diff.ToString("#########.000000"));
                        total_diff = Math.Round(total_diff, 2, MidpointRounding.AwayFromZero);
                        if (baseType == "13")
                        {
                            if (total_diff != 0)
                            {
                                rs.MoveFirst();
                                oJE.ReferenceDate = date;
                                oJE.Memo = "DO Charge Out";
                                //rs1.DoQuery("select * from oinv where docentry =" + rs.Fields.Item("SODocEntry").Value.ToString());
                                //oJE.UserFields.Fields.Item("U_RInvNo").Value = rs1.Fields.Item("docnum").Value.ToString();
                                currentline = 0;
                                if (total_diff != 0)
                                {
                                    currentline++;
                                    oJE.Lines.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                    if (total_diff > 0)
                                        oJE.Lines.Debit = total_diff;
                                    else if (total_diff < 0)
                                        oJE.Lines.Credit = -total_diff;
                                    foreach (System.Data.DataRow dr in dt.Rows)
                                    {
                                        currentline++;
                                        if (currentline > 1)
                                        {
                                            oJE.Lines.Add();
                                            oJE.Lines.SetCurrentLine(currentline - 1);
                                        }
                                        different = double.Parse(dr["different"].ToString());
                                        quantity = double.Parse(dr["quantity"].ToString());
                                        total_diff = Math.Round(quantity * different, 2, MidpointRounding.AwayFromZero);
                                        oJE.Lines.AccountCode = cogsAcct;
                                        if (total_diff > 0)
                                            oJE.Lines.Credit = total_diff;
                                        else if (total_diff < 0)
                                            oJE.Lines.Debit = -total_diff;
                                        if (!string.IsNullOrEmpty(dr["productgroup"].ToString()))
                                            oJE.Lines.CostingCode = dr["productgroup"].ToString();

                                    }
                                }
                                retcode = oJE.Add();
                                if (retcode != 0)
                                {
                                    if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                    //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return rtn;
                                }
                                SAP.SBOCompany.GetNewObjectCode(out oJEKey);
                                total_diff = 0;
                            }
                        }
                    }
                    string sql = "";
                    bool orderfound = false;
                    //sql ="select T1.VisOrder, T0.U_TPCODE, T0.U_LORRY, T0.U_DRIVER, T0.U_DRIVERIC,U_AREA,T1.LineId, T1.*, T3.U_BASEOBJ as [ObjType], T3.U_BASEENT AS [SODocEntry], T3.U_BASELINE AS [SOBaseLine] " +
                    //    " from  [@FT_CHARGE] T0 inner join [@FT_CHARGE1] T1 on T0.docentry = T1.docentry " +
                    //    " inner join[@FT_TPPLAN1] T2 on T1.U_BASEENT = T2.DocEntry and T1.U_BASELINE = T2.LineId " +
                    //    " inner join[@FT_SPLAN1] T3 on T2.U_BASEENT = T3.DocEntry and T2.U_BASELINE = T3.LineId WHERE T1.DOCENTRY =" + docEntry + " order by T3.U_BASEOBJ ";
                    //sql = "select T0.U_TPCODE, T0.U_LORRY, T0.U_DRIVER, T0.U_DRIVERIC,U_AREA, sum(T1.U_quantity) as [U_Quantity],T1.U_SOITEMCO, T1.U_WhsCode, T3.U_BASEOBJ as [ObjType], T3.U_BASEENT AS [SODocEntry], T3.U_BASELINE AS [SOBaseLine] " +
                    //   " from  [@FT_CHARGE] T0 inner join [@FT_CHARGE1] T1 on T0.docentry = T1.docentry " +
                    //   " inner join[@FT_TPPLAN1] T2 on T1.U_BASEENT = T2.DocEntry and T1.U_BASELINE = T2.LineId " +
                    //   " inner join[@FT_SPLAN1] T3 on T2.U_BASEENT = T3.DocEntry and T2.U_BASELINE = T3.LineId WHERE T1.DOCENTRY =" + docEntry +
                    //   " group by T0.U_TPCODE, T0.U_LORRY, T0.U_DRIVER, T0.U_DRIVERIC,U_AREA,T1.U_SOITEMCO, T1.U_WhsCode, T3.U_BASEOBJ, T3.U_BASEENT , T3.U_BASELINE order by T3.U_BASEOBJ ";
                    sql = "select T0.U_APRICE, T0.U_DOCTOTAL, T0.U_REMARKS, T0.U_FREMARKS, T0.U_TPCODE, T0.U_LORRY, T0.U_DRIVER, T0.U_DRIVERIC,U_AREA, sum(T1.U_quantity) as [U_Quantity],T1.U_SOITEMCO, T1.U_WhsCode, T1.U_SOBASEOB as [ObjType], T1.U_SOENTRY AS [SODocEntry], T1.U_SOLINE AS [SOBaseLine], T0.U_Address as [Address] " +
                        " from  [@FT_CHARGE] T0 inner join [@FT_CHARGE1] T1 on T0.docentry = T1.docentry " +
                        " WHERE T1.DOCENTRY =" + docEntry +
                        " group by T0.U_APRICE, T0.U_DOCTOTAL, T0.U_REMARKS, T0.U_FREMARKS, T0.U_TPCODE, T0.U_LORRY, T0.U_DRIVER, T0.U_DRIVERIC,U_AREA,T1.U_SOITEMCO, T1.U_WhsCode, T1.U_SOBASEOB, T1.U_SOENTRY, T1.U_SOLINE, T0.U_Address order by T1.U_SOBASEOB ";
                    rs.DoQuery(sql);
                    
                    if (rs.RecordCount > 0)
                    {
                        cnt = 0;
                        if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();
                        SAPbobsCOM.Documents oOrder = null;
                        SAPbobsCOM.Documents oDel = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDeliveryNotes);
                        oDel.CardCode = cardcode;
                        oDel.DocDate = date;
                        oDel.UserFields.Fields.Item("U_ChargeNo").Value = docnum;
                        oDel.UserFields.Fields.Item("U_Diff").Value = total_diff;
                        if (owner > 0)
                            oDel.DocumentsOwner = owner;
                        //if (rs.Fields.Item("U_REMARKS").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_REMARKS").Value = rs.Fields.Item("U_REMARKS").Value.ToString();
                        //}
                        //if (rs.Fields.Item("U_FREMARKS").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_FREMARKS").Value = rs.Fields.Item("U_FREMARKS").Value.ToString();
                        //}
                        //if (rs.Fields.Item("U_TPCODE").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_Transporter").Value = rs.Fields.Item("U_TPCODE").Value.ToString();
                        //}
                        //if (rs.Fields.Item("U_LORRY").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_LorryNo").Value = rs.Fields.Item("U_LORRY").Value.ToString();
                        //}
                        //if (rs.Fields.Item("U_DRIVER").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_Driver").Value = rs.Fields.Item("U_DRIVER").Value.ToString();
                        //}
                        //if (rs.Fields.Item("U_DRIVERIC").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_ICNo").Value = rs.Fields.Item("U_DRIVERIC").Value.ToString();
                        //}
                        //if (rs.Fields.Item("U_AREA").Value.ToString() != "")
                        //{
                        //    oDel.UserFields.Fields.Item("U_Area").Value = rs.Fields.Item("U_AREA").Value.ToString();
                        //}
                        //float temp = 0;
                        //if (float.TryParse(rs.Fields.Item("U_DOCTOTAL").Value.ToString(), out temp))
                        //{
                        //    oDel.UserFields.Fields.Item("U_Total_TC").Value = temp;
                        //}
                        //temp = 0;
                        //if (float.TryParse(rs.Fields.Item("U_APRICE").Value.ToString(), out temp))
                        //{
                        //    oDel.UserFields.Fields.Item("U_UnitCost_MT").Value = temp;                            
                        //}
                        
                        oDel.Address2 = rs.Fields.Item("Address").Value.ToString();
                        rs.MoveFirst();
                        while (!rs.EoF)
                        {
                            if (cnt > 0)
                            {
                                oDel.Lines.Add();
                                oDel.Lines.SetCurrentLine(cnt);
                            }
                            else
                            {
                                if (rs.Fields.Item("ObjType").Value.ToString() == "17")
                                {
                                    oOrder = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                                    orderfound = oOrder.GetByKey(int.Parse(rs.Fields.Item("SODocEntry").Value.ToString()));
                                }
                                else if (rs.Fields.Item("ObjType").Value.ToString() == "13")
                                {
                                    oOrder = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);
                                    orderfound = oOrder.GetByKey(int.Parse(rs.Fields.Item("SODocEntry").Value.ToString()));
                                }
                                if (orderfound)
                                {
                                    if (!String.IsNullOrEmpty(oOrder.Address)) oDel.Address = oOrder.Address;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToAddress2)) oDel.AddressExtension.BillToAddress2 = oOrder.AddressExtension.BillToAddress2;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToAddress3)) oDel.AddressExtension.BillToAddress3 = oOrder.AddressExtension.BillToAddress3;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToAddressType)) oDel.AddressExtension.BillToAddressType = oOrder.AddressExtension.BillToAddressType;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToBlock)) oDel.AddressExtension.BillToBlock = oOrder.AddressExtension.BillToBlock;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToBuilding)) oDel.AddressExtension.BillToBuilding = oOrder.AddressExtension.BillToBuilding;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToCity)) oDel.AddressExtension.BillToCity = oOrder.AddressExtension.BillToCity;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToCountry)) oDel.AddressExtension.BillToCountry = oOrder.AddressExtension.BillToCountry;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToCounty)) oDel.AddressExtension.BillToCounty = oOrder.AddressExtension.BillToCounty;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToGlobalLocationNumber)) oDel.AddressExtension.BillToGlobalLocationNumber = oOrder.AddressExtension.BillToGlobalLocationNumber;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToState)) oDel.AddressExtension.BillToState = oOrder.AddressExtension.BillToState;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToStreet)) oDel.AddressExtension.BillToStreet = oOrder.AddressExtension.BillToStreet;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToStreetNo)) oDel.AddressExtension.BillToStreetNo = oOrder.AddressExtension.BillToStreetNo;
                                    if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToZipCode)) oDel.AddressExtension.BillToZipCode = oOrder.AddressExtension.BillToZipCode;
                                    //if (!String.IsNullOrEmpty(oOrder.Address2)) oDel.Address2 = oOrder.Address2;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToAddress2)) oDO.AddressExtension.ShipToAddress2 = oSO.AddressExtension.ShipToAddress2;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToAddress3)) oDO.AddressExtension.ShipToAddress3 = oSO.AddressExtension.ShipToAddress3;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToAddressType)) oDO.AddressExtension.ShipToAddressType = oSO.AddressExtension.ShipToAddressType;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToBlock)) oDO.AddressExtension.ShipToBlock = oSO.AddressExtension.ShipToBlock;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToBuilding)) oDO.AddressExtension.ShipToBuilding = oSO.AddressExtension.ShipToBuilding;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToCity)) oDO.AddressExtension.ShipToCity = oSO.AddressExtension.ShipToCity;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToCountry)) oDO.AddressExtension.ShipToCountry = oSO.AddressExtension.ShipToCountry;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToCounty)) oDO.AddressExtension.ShipToCounty = oSO.AddressExtension.ShipToCounty;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToGlobalLocationNumber)) oDO.AddressExtension.ShipToGlobalLocationNumber = oSO.AddressExtension.ShipToGlobalLocationNumber;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToState)) oDO.AddressExtension.ShipToState = oSO.AddressExtension.ShipToState;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToStreet)) oDO.AddressExtension.ShipToStreet = oSO.AddressExtension.ShipToStreet;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToStreetNo)) oDO.AddressExtension.ShipToStreetNo = oSO.AddressExtension.ShipToStreetNo;
                                    //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToZipCode)) oDO.AddressExtension.ShipToZipCode = oSO.AddressExtension.ShipToZipCode;
                                }
                                if (rs.Fields.Item("ObjType").Value.ToString() == "13")
                                {
                                    orderfound = false;
                                }
                            }

                            oDel.Lines.ItemCode = rs.Fields.Item("U_SOITEMCO").Value.ToString();
                            //oDel.Lines.UserFields.Fields.Item("U_RevisedItemCode").Value = rs.Fields.Item("U_ITEMCODE").Value.ToString();
                            //oDel.Lines.UserFields.Fields.Item("U_RevisedItemDesc").Value = rs.Fields.Item("U_ITEMNAME").Value.ToString();
                            oDel.Lines.Quantity = double.Parse(rs.Fields.Item("U_QUANTITY").Value.ToString());
                            oDel.Lines.WarehouseCode = rs.Fields.Item("U_WHSCODE").Value.ToString();
                            if (rs.Fields.Item("ObjType").Value.ToString() == "17")
                            {
                                oDel.Lines.BaseType = 17;
                                rs1.DoQuery("select * from rdr1 where docentry = " + rs.Fields.Item("SODocEntry").Value.ToString() + " and linenum = " + rs.Fields.Item("SOBaseLine").Value.ToString());
                                oDel.Lines.PriceAfterVAT = double.Parse(rs1.Fields.Item("PriceAfVAT").Value.ToString());
                                oDel.Lines.VatGroup = rs1.Fields.Item("VatGroup").Value.ToString();
                                oDel.Lines.COGSAccountCode = rs1.Fields.Item("CogsAcct").Value.ToString();
                                oDel.Lines.AccountCode = rs1.Fields.Item("AcctCode").Value.ToString();
                                oDel.Lines.LineTotal = double.Parse(rs.Fields.Item("U_QUANTITY").Value.ToString()) * double.Parse(rs1.Fields.Item("Price").Value.ToString());
                            }
                            else if (rs.Fields.Item("ObjType").Value.ToString() == "13")
                            {
                                oDel.Lines.BaseType = 13;
                                rs1.DoQuery("select * from inv1 where docentry = " + rs.Fields.Item("SODocEntry").Value.ToString() + " and linenum = " + rs.Fields.Item("SOBaseLine").Value.ToString());
                                oDel.Lines.PriceAfterVAT = double.Parse(rs1.Fields.Item("PriceAfVAT").Value.ToString());
                                oDel.Lines.VatGroup = rs1.Fields.Item("VatGroup").Value.ToString();
                                oDel.Lines.COGSAccountCode = rs1.Fields.Item("CogsAcct").Value.ToString();
                                oDel.Lines.AccountCode = rs1.Fields.Item("AcctCode").Value.ToString();
                                oDel.Lines.LineTotal = double.Parse(rs.Fields.Item("U_QUANTITY").Value.ToString()) * double.Parse(rs1.Fields.Item("Price").Value.ToString());
                            }
                            oDel.Lines.BaseLine = int.Parse(rs.Fields.Item("SOBaseLine").Value.ToString());
                            oDel.Lines.BaseEntry = int.Parse(rs.Fields.Item("SODocEntry").Value.ToString());

                            cntbatch = 0;
                            cntbin = 0;

                            rs1.DoQuery("select ManBtchNum from OITM where itemcode ='" + rs.Fields.Item("U_SOITEMCO").Value.ToString() + "'");
                            if (rs1.RecordCount > 0)
                            {
                                ManBtchNum = rs1.Fields.Item("ManBtchNum").Value.ToString();
                                if (ManBtchNum == "Y")
                                {
                                    sql = "select sum(T2.U_Quantity) as [Quantity], T2.U_BATCHNUM as [BatchNum] from [@FT_Charge2] T2 inner join [@FT_Charge1] T1 " +
                                        " on T1.docentry = T2.docentry and T2.U_BASEVIS = T1.VisOrder " +
                                        " where T1.docentry = " + docEntry +
                                        " and T1.U_SOBASEOB = " + rs.Fields.Item("ObjType").Value.ToString() +
                                        " and T1.U_SOENTRY = " + rs.Fields.Item("SODocEntry").Value.ToString() +
                                        " and T1.U_SOLINE = " + rs.Fields.Item("SOBaseLine").Value.ToString() +
                                        " group by T2.U_BATCHNUM ";
                                    rs1.DoQuery(sql);
                                    //rs1.DoQuery("select sum(U_Quantity)  as [Quantity], U_BATCHNUM as [BatchNum] from [@FT_Charge2] where docentry =" + docEntry +
                                    //    " and U_BASEVIS =" + rs.Fields.Item("VisOrder").Value.ToString() + " group by U_BATCHNUM ");
                                    while (!rs1.EoF)
                                    {
                                        if (cntbatch > 0)
                                        {
                                            oDel.Lines.BatchNumbers.Add();
                                            oDel.Lines.BatchNumbers.SetCurrentLine(cntbatch);
                                        }
                                        oDel.Lines.BatchNumbers.BatchNumber = rs1.Fields.Item("BatchNum").Value.ToString();
                                        oDel.Lines.BatchNumbers.Quantity = double.Parse(rs1.Fields.Item("Quantity").Value.ToString());

                                        #region bin
                                        rs2.DoQuery("select BinActivat from OWHS where WhsCode ='" + rs.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                        if (rs2.RecordCount > 0)
                                        {
                                            BinActivat = rs2.Fields.Item("BinActivat").Value.ToString();
                                            if (BinActivat == "Y")
                                            {
                                                sql = "select sum(T2.U_Quantity) as [Quantity], T2.U_BINABS as [Bin] from [@FT_Charge2] T2 inner join [@FT_Charge1] T1 " +
                                                    " on T1.docentry = T2.docentry and T2.U_BASEVIS = T1.VisOrder " +
                                                    " where T1.docentry = " + docEntry +
                                                    " and T1.U_SOBASEOB = " + rs.Fields.Item("ObjType").Value.ToString() +
                                                    " and T1.U_SOENTRY = " + rs.Fields.Item("SODocEntry").Value.ToString() +
                                                    " and T1.U_SOLINE = " + rs.Fields.Item("SOBaseLine").Value.ToString() +
                                                    " and T2.U_BATCHNUM = '" + rs1.Fields.Item("BatchNum").Value.ToString() + "' " +
                                                    " group by T2.U_BINABS ";
                                                //sql = "select sum(U_Quantity)  as [Quantity], U_BINABS as [Bin] from [@FT_Charge2] where docentry =" + docEntry +
                                                //    " and U_BASEVIS =" + rs.Fields.Item("VisOrder").Value.ToString() + " and U_BATCHNUM = '" + rs1.Fields.Item("BatchNum").Value.ToString() + "' group by U_BINABS ";
                                                rs2.DoQuery(sql);
                                                while (!rs2.EoF)
                                                {
                                                    if (cntbin > 0)
                                                    {
                                                        oDel.Lines.BinAllocations.Add();
                                                        oDel.Lines.BinAllocations.SetCurrentLine(cntbin);
                                                    }
                                                    oDel.Lines.BinAllocations.BinAbsEntry = int.Parse(rs2.Fields.Item("Bin").Value.ToString());
                                                    oDel.Lines.BinAllocations.Quantity = double.Parse(rs2.Fields.Item("Quantity").Value.ToString());
                                                    oDel.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = cntbatch;

                                                    rs2.MoveNext();
                                                    cntbin++;
                                                }
                                            }
                                        }
                                        #endregion

                                        rs1.MoveNext();
                                        cntbatch++;
                                    }
                                }
                                else
                                {
                                    #region bin
                                    rs2.DoQuery("select BinActivat from OWHS where WhsCode ='" + rs.Fields.Item("U_WHSCODE").Value.ToString() + "'");
                                    if (rs2.RecordCount > 0)
                                    {
                                        BinActivat = rs2.Fields.Item("BinActivat").Value.ToString();
                                        if (BinActivat == "Y")
                                        {
                                            sql = "select sum(T2.U_Quantity) as [Quantity], T2.U_BINABS as [Bin] from [@FT_Charge2] T2 inner join [@FT_Charge1] T1 " +
                                                " on T1.docentry = T2.docentry and T2.U_BASEVIS = T1.VisOrder " +
                                                " where T1.docentry = " + docEntry +
                                                " and T1.U_SOBASEOB = " + rs.Fields.Item("ObjType").Value.ToString() +
                                                " and T1.U_SOENTRY = " + rs.Fields.Item("SODocEntry").Value.ToString() +
                                                " and T1.U_SOLINE = " + rs.Fields.Item("SOBaseLine").Value.ToString() +
                                                " group by T2.U_BINABS ";
                                            //sql = "select sum(U_Quantity)  as [Quantity], U_BINABS as [Bin] from [@FT_Charge2] where docentry =" + docEntry +
                                            //    " and U_BASEVIS =" + rs.Fields.Item("VisOrder").Value.ToString() + " group by U_BINABS ";
                                            rs2.DoQuery(sql);
                                            while (!rs2.EoF)
                                            {
                                                if (cntbin > 0)
                                                {
                                                    oDel.Lines.BinAllocations.Add();
                                                    oDel.Lines.BinAllocations.SetCurrentLine(cntbin);
                                                }
                                                oDel.Lines.BinAllocations.BinAbsEntry = int.Parse(rs2.Fields.Item("Bin").Value.ToString());
                                                oDel.Lines.BinAllocations.Quantity = double.Parse(rs2.Fields.Item("Quantity").Value.ToString());
                                                //oDel.Lines.BinAllocations.SerialAndBatchNumbersBaseLine = cntbatch;

                                                rs2.MoveNext();
                                                cntbin++;
                                            }
                                        }
                                    }
                                    #endregion

                                }
                            }
                            cnt++;
                            rs.MoveNext();
                        }
                        retcode = oDel.Add();
                        if (retcode != 0)
                        {
                            string test = SAP.SBOCompany.GetLastErrorDescription();
                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            if (retcode == -10)
                                SAP.SBOApplication.MessageBox("Delivery Creation Error " + SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                            else
                                SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                            //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                            return rtn;
                        }
                        SAP.SBOCompany.GetNewObjectCode(out oDelKey);

                        if (orderfound)
                        {
                            oDel.GetByKey(int.Parse(oDelKey));

                            if (!String.IsNullOrEmpty(oOrder.Address)) oDel.Address = oOrder.Address;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToAddress2)) oDel.AddressExtension.BillToAddress2 = oOrder.AddressExtension.BillToAddress2;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToAddress3)) oDel.AddressExtension.BillToAddress3 = oOrder.AddressExtension.BillToAddress3;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToAddressType)) oDel.AddressExtension.BillToAddressType = oOrder.AddressExtension.BillToAddressType;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToBlock)) oDel.AddressExtension.BillToBlock = oOrder.AddressExtension.BillToBlock;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToBuilding)) oDel.AddressExtension.BillToBuilding = oOrder.AddressExtension.BillToBuilding;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToCity)) oDel.AddressExtension.BillToCity = oOrder.AddressExtension.BillToCity;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToCountry)) oDel.AddressExtension.BillToCountry = oOrder.AddressExtension.BillToCountry;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToCounty)) oDel.AddressExtension.BillToCounty = oOrder.AddressExtension.BillToCounty;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToGlobalLocationNumber)) oDel.AddressExtension.BillToGlobalLocationNumber = oOrder.AddressExtension.BillToGlobalLocationNumber;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToState)) oDel.AddressExtension.BillToState = oOrder.AddressExtension.BillToState;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToStreet)) oDel.AddressExtension.BillToStreet = oOrder.AddressExtension.BillToStreet;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToStreetNo)) oDel.AddressExtension.BillToStreetNo = oOrder.AddressExtension.BillToStreetNo;
                            if (!String.IsNullOrEmpty(oOrder.AddressExtension.BillToZipCode)) oDel.AddressExtension.BillToZipCode = oOrder.AddressExtension.BillToZipCode;
                            //if (!String.IsNullOrEmpty(oOrder.Address2)) oDel.Address2 = oOrder.Address2;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToAddress2)) oDO.AddressExtension.ShipToAddress2 = oSO.AddressExtension.ShipToAddress2;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToAddress3)) oDO.AddressExtension.ShipToAddress3 = oSO.AddressExtension.ShipToAddress3;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToAddressType)) oDO.AddressExtension.ShipToAddressType = oSO.AddressExtension.ShipToAddressType;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToBlock)) oDO.AddressExtension.ShipToBlock = oSO.AddressExtension.ShipToBlock;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToBuilding)) oDO.AddressExtension.ShipToBuilding = oSO.AddressExtension.ShipToBuilding;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToCity)) oDO.AddressExtension.ShipToCity = oSO.AddressExtension.ShipToCity;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToCountry)) oDO.AddressExtension.ShipToCountry = oSO.AddressExtension.ShipToCountry;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToCounty)) oDO.AddressExtension.ShipToCounty = oSO.AddressExtension.ShipToCounty;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToGlobalLocationNumber)) oDO.AddressExtension.ShipToGlobalLocationNumber = oSO.AddressExtension.ShipToGlobalLocationNumber;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToState)) oDO.AddressExtension.ShipToState = oSO.AddressExtension.ShipToState;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToStreet)) oDO.AddressExtension.ShipToStreet = oSO.AddressExtension.ShipToStreet;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToStreetNo)) oDO.AddressExtension.ShipToStreetNo = oSO.AddressExtension.ShipToStreetNo;
                            //if (!String.IsNullOrEmpty(oOrder.AddressExtension.ShipToZipCode)) oDO.AddressExtension.ShipToZipCode = oSO.AddressExtension.ShipToZipCode;

                            retcode = oDel.Update();
                            if (retcode != 0)
                            {
                                string test = SAP.SBOCompany.GetLastErrorDescription();
                                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                if (retcode == -10)
                                    SAP.SBOApplication.MessageBox("Delivery Update Error " + SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                else
                                    SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                return rtn;
                            }

                        }

                    }

                    if (retcode != 0)
                    {
                        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                        return rtn;
                        //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                    rs.DoQuery("Select T0.docnum, T0.docdate,T1.basetype,T1.baseentry, T1.itemcode, T1.quantity as [quantity], T1.whscode from ODLN T0 inner join DLN1 T1 on T0.docentry = T1.docentry " +
                        " where T0.docentry =" + oDelKey +
                        " order by T1.linenum");// + " group by  T0.docnum, T0.docdate,T1.basetype, T1.baseentry,T1.itemcode,T1.whscode ");
                    if (rs.RecordCount > 0)
                    {
                        oDeldocnum = rs.Fields.Item("docnum").Value.ToString();


                        if (rs.Fields.Item("basetype").Value.ToString() == "13")
                        {
                            double total = 0;
                            string baseentry = "", RInv_docnum = "";
                            rs.MoveFirst();
                            baseentry = rs.Fields.Item("baseentry").Value.ToString();                           

                            rs1.DoQuery("select docnum from oinv where docentry=" + baseentry);
                            if (rs1.RecordCount > 0)
                            {
                                RInv_docnum = rs1.Fields.Item("docnum").Value.ToString();
                            }


                            date = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());

                            if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();

                            oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                            oJE.ReferenceDate = date;
                            oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
                            oJE.Memo = "Provision for Reserved Invoiced";
                            total = 0;
                            currentline = 0;
                            while (!rs.EoF)
                            {
                                rs1.DoQuery("select T0.avgprice, T1.U_CostCenter from oitm T0 inner join oitb T1 on T0.itmsgrpcod = T1.itmsgrpcod where isnull(T0.invntitem,'N') = 'Y' and T0.itemcode='" + rs.Fields.Item("itemcode").Value.ToString() + "' ");
                                if (rs1.RecordCount > 0)
                                {
                                    currentline++;
                                    if (currentline > 1)
                                    {
                                        oJE.Lines.Add();
                                        oJE.Lines.SetCurrentLine(currentline - 1);
                                    }
                                    productgroup = rs1.Fields.Item("U_CostCenter").Value.ToString();
                                    linetotal = double.Parse(rs.Fields.Item("quantity").Value.ToString()) * double.Parse(rs1.Fields.Item("avgprice").Value.ToString());
                                    linetotal = Math.Round(linetotal, 2, MidpointRounding.AwayFromZero);

                                    if (linetotal > 0)
                                    {
                                        oJE.Lines.AccountCode = cogsAcct;
                                        oJE.Lines.Credit = linetotal;
                                        if (!string.IsNullOrEmpty(productgroup))
                                            oJE.Lines.CostingCode = productgroup;
                                        else
                                            productgroup = "";
                                        total = total + linetotal;
                                     }
                                }
                                rs.MoveNext();
                            }
                            total = double.Parse(total.ToString("#########.000000"));
                            total = Math.Round(total, 2, MidpointRounding.AwayFromZero);

                            currentline++;
                            if (currentline > 1)
                            {
                                oJE.Lines.Add();
                                oJE.Lines.SetCurrentLine(currentline - 1);
                            }
                            oJE.Lines.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                            oJE.Lines.Debit = total;

                            if (total != 0)
                            {
                                retcode = oJE.Add();
                                if (retcode != 0)
                                {
                                    if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                    return rtn;
                                }                               
                            }

                            rs1.DoQuery("Select T4.U_CostCenter, isnull(T2.U_InvCost,0) as U_InvCost, isnull(T3.avgprice,0) as avgprice, sum(T1.quantity) as quantity from ODLN T0 inner join DLN1 T1 on T0.docentry = T1.docentry " +
                                " inner join INV1 T2 on T1.baseentry = T2.docentry and T1.baseline = T2.linenum and T1.basetype = 13" +
                                " inner join OITM T3 on T3.itemcode = T1.itemcode" +
                                " inner join OITB T4 on T4.itmsgrpcod = T3.itmsgrpcod" +
                                " where isnull(T3.invntitem,'N') = 'Y' and T0.docentry =" + oDelKey +
                                " group by T4.U_CostCenter, T2.U_InvCost, T3.avgprice");// + " group by  T0.docnum, T0.docdate,T1.basetype, T1.baseentry,T1.itemcode,T1.whscode ");
                            if (rs1.RecordCount > 0)
                            {
                                if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();
                                SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                                oJE1.ReferenceDate = date;
                                oJE1.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
                                oJE1.Memo = "COGS Adjustment";
                                currentline = 0;
                                total_diff = 0;
                                rs1.MoveFirst();
                                while (!rs1.EoF)
                                {
                                    productgroup = rs1.Fields.Item("U_CostCenter").Value.ToString();
                                    temp = double.Parse(rs1.Fields.Item("U_InvCost").Value.ToString()) * double.Parse(rs1.Fields.Item("quantity").Value.ToString());
                                    total = double.Parse(rs1.Fields.Item("avgprice").Value.ToString()) * double.Parse(rs1.Fields.Item("quantity").Value.ToString());
                                    temp = temp - total;
                                    temp = Math.Round(temp, 2, MidpointRounding.AwayFromZero);

                                    if (temp != 0)
                                    {
                                        total_diff = total_diff + temp;

                                        //currentline++;
                                        //if (currentline > 1)
                                        //{
                                        //    oJE1.Lines.Add();
                                        //    oJE1.Lines.SetCurrentLine(currentline - 1);
                                        //}
                                        ////oJE1.Lines.AccountCode = "530700";// "530900"; Stock Gain/Loss
                                        //oJE1.Lines.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                        //if (temp > 0)
                                        //    oJE1.Lines.Debit = temp;
                                        //else if (temp < 0)
                                        //    oJE1.Lines.Credit = -temp;
                                        //if (!string.IsNullOrEmpty(productgroup))
                                        //    oJE1.Lines.CostingCode = productgroup;

                                        currentline++;
                                        if (currentline > 1)
                                        {
                                            oJE1.Lines.Add();
                                            oJE1.Lines.SetCurrentLine(currentline - 1);
                                        }
                                        oJE1.Lines.AccountCode = cogsAcct;
                                        if (temp > 0)
                                            oJE1.Lines.Credit = temp;
                                        else if (temp < 0)
                                            oJE1.Lines.Debit = -temp;
                                        oJE1.Lines.CostingCode = productgroup;
                                    }
                                    rs1.MoveNext();
                                }
                                total_diff = double.Parse(total_diff.ToString("#########.000000"));
                                total_diff = Math.Round(total_diff, 2, MidpointRounding.AwayFromZero);

                                currentline++;
                                if (currentline > 1)
                                {
                                    oJE1.Lines.Add();
                                    oJE1.Lines.SetCurrentLine(currentline - 1);
                                }
                                //oJE1.Lines.AccountCode = "530700";// "530900"; Stock Gain/Loss
                                oJE1.Lines.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                if (total_diff > 0)
                                    oJE1.Lines.Debit = total_diff;
                                else if (total_diff < 0)
                                    oJE1.Lines.Credit = -total_diff;
                                //if (!string.IsNullOrEmpty(productgroup))
                                //    oJE1.Lines.CostingCode = productgroup;

                                if (total_diff != 0)
                                {
                                    retcode = oJE1.Add();
                                    if (retcode != 0)
                                    {
                                        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                        return rtn;
                                    }
                                }
                            }
                        }

                        if (oReceiptKey != "")
                        {
                            oReceipt = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenEntry);
                            oReceipt.GetByKey(int.Parse(oReceiptKey));
                            oReceipt.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
                            oReceipt.Update();
                        }
                        if (oIssueKey != "")
                        {
                            oIssue = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInventoryGenExit);
                            oIssue.GetByKey(int.Parse(oIssueKey));
                            oIssue.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
                            oIssue.Update();
                        }
                        if (oJEKey != "")
                        {
                            oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
                            oJE.GetByKey(int.Parse(oJEKey));
                            oJE.UserFields.Fields.Item("U_DelNo").Value = oDeldocnum;
                            oJE.Update();
                        }
                    }

                    rs.DoQuery("Exec FT_CloseChargeDoc " + docEntry.ToString());
                    if (rs.RecordCount <= 0)
                    {
                        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        SAP.SBOApplication.MessageBox("genDO FT_CloseChargeDoc error.", 1, "Ok", "", "");
                        return rtn;
                    }
                    else
                    {
                        if (rs.Fields.Item(0).Value.ToString() != "0")
                        {
                            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                            SAP.SBOApplication.MessageBox("genDO " + rs.Fields.Item(1).Value.ToString(), 1, "Ok", "", "");
                            return rtn;
                        }
                    }
                    //rs.DoQuery("update [@FT_CHARGE] set Status = 'C' where DocEntry = " + docEntry.ToString());
                    //rs.DoQuery("update [@FT_TPPLAN] set Status = 'C' where DocNum = " + tpdocnum);
                }
                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                rtn = true;

                return rtn;
            }
            catch (Exception ex)
            {
                if (FT_ADDON.SAP.SBOCompany.InTransaction)
                {
                    FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }

                SAP.SBOApplication.MessageBox("genDO " + ex.Message, 1, "Ok", "", "");

                return false;
            }
        }

        private static void sendAppMsg(string ObjType, string TrxType, string DocEntry, string msgsubject, string msg, string linkvalue)
        {

            ApprovalMsgClass oSystemMsgClass = new ApprovalMsgClass();

            oSystemMsgClass.IsApprove = false;
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

        private static void sendMsg(string ObjType, string TrxType, string DocEntry, string msgsubject, string msg, string linkvalue)
        {

            SystemMsgClass oSystemMsgClass = new SystemMsgClass();

            oSystemMsgClass.ObjType = ObjType;
            oSystemMsgClass.TrxType = TrxType;

            string temp = oSystemMsgClass.TrxType == "A" ? "Added." : "Updated.";
            oSystemMsgClass.MsgSubject = msgsubject;
            oSystemMsgClass.Msg = msg;
            oSystemMsgClass.ColumnName = "Description";

            oSystemMsgClass.LinkObj = ObjType;
            oSystemMsgClass.LineValue = linkvalue;
            oSystemMsgClass.LinkKey = DocEntry;

            oSystemMsgClass.SendMsg();

        }

        private static void duplicateDoc(SAPbouiCOM.Form oForm, string sDocEntry)
        {
            try
            {
                string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value;
                string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value;
                SAPbouiCOM.DBDataSource ds = oForm.DataSources.DBDataSources.Item("@" + dsname);
                SAPbouiCOM.DBDataSource ds1 = oForm.DataSources.DBDataSources.Item("@" + dsname1);
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "select * from [@" + dsname + "] where DocEntry = " + sDocEntry;
                string sql1 = "select * from [@" + dsname1 + "] where DocEntry = " + sDocEntry;

                string columnname = "";
                rs.DoQuery(sql);
                if (rs.RecordCount > 0)
                {
                    rs.MoveFirst();
                    foreach (SAPbobsCOM.Field rsf in rs.Fields)
                    {
                        columnname = rsf.Description;
                        if (columnname.Substring(0,2) == "U_")
                        {
                            switch (columnname)
                            {
                                case "U_DOCDATE":
                                    break;
                                default:
                                    ds.SetValue(columnname, 0, rsf.Value.ToString());
                                    break;
                            }
                        }
                    }
                }

                SAPbobsCOM.SBObob dt = (SAPbobsCOM.SBObob)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);

                DateTime date;
                bool detailfound = false;
                string temp = "";
                int cnt = -1;
                rs1.DoQuery(sql1);
                if (rs1.RecordCount > 0)
                {
                    rs1.MoveFirst();

                    while (!rs1.EoF)
                    {
                        if (rs1.Fields.Item("U_ITEMCODE").Value != null)
                        {
                            if (rs1.Fields.Item("U_ITEMCODE").Value.ToString().Trim() != "")
                            {
                                detailfound = true;
                                cnt++;
                                if (cnt > 0)
                                {
                                    ds1.InsertRecord(0);
                                }
                                foreach (SAPbobsCOM.Field rsf in rs1.Fields)
                                {
                                    columnname = rsf.Description;
                                    if (columnname.Substring(0, 2) == "U_")
                                    {
                                        switch (columnname)
                                        {

                                            case "U_ORIQTY":
                                            case "U_RBPONO":
                                            case "U_RBPOQTY":
                                            case "U_RBPOOPEN":
                                            case "U_TPQTY":
                                            case "U_CMQTY":
                                                ds1.SetValue(columnname, cnt, "0");
                                                break;
                                            default:
                                                if (rsf.Type == SAPbobsCOM.BoFieldTypes.db_Date)
                                                {
                                                    if (rsf.Value != null)
                                                    {
                                                        date = DateTime.Parse(rsf.Value.ToString());
                                                        //rs = dt.Format_DateToString(date);
                                                        //rs.MoveFirst();
                                                        //temp = rs.Fields.Item(0).Value.ToString();
                                                        temp = date.ToString("yyyyMMdd");
                                                        ds1.SetValue(columnname, cnt, temp);
                                                    }
                                                }
                                                else
                                                {
                                                    temp = rsf.Value.ToString();
                                                    ds1.SetValue(columnname, cnt, temp);
                                                }

                                                break;
                                        }
                                    }
                                }

                            }
                        }
                        rs1.MoveNext();
                    }

                    if (detailfound)
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                        oMatrix.LoadFromDataSource();
                        SPLANRemainingOpen(oForm, true);
                        int row = ds1.Size;
                        double qty = 0;
                        for (int y = row - 1; y >= 0; y--)
                        {
                            if (double.TryParse(ds1.GetValue("U_QUANTITY", y), out qty))
                            {
                                if (qty <= 0)
                                    ds1.RemoveRecord(y);
                            }
                        }
                        oMatrix.LoadFromDataSource();
                        arrangematrix(oForm, oMatrix, "@" + dsname1);

                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Menu Event Before " + ex.Message, 1, "Ok", "", "");
            }
        }

        private static void checkstock(SAPbouiCOM.Form oForm)
        {
            string sql = @"select T9.U_ITEMCODE, T9.U_WHSCODE, OITW.OnHand, T9.U_QUANTITY, T9.U_CMQTY, $[current] as Current_Plan
from
(
select T1.U_ITEMCODE, T1.U_WHSCODE, sum(T1.U_QUANTITY) as U_QUANTITY, sum(isnull(T00.U_CMQTY,0)) as U_CMQTY
from [@FT_SPLAN] T0 inner join [@FT_SPLAN1] T1 on T0.DocEntry = T1.DocEntry and T0.Canceled = 'N' and T0.DocEntry <> 0 and T1.U_RBPOQTY = 0
inner join OITM on OITM.ItemCode = T1.U_ITEMCODE and OITM.InvntItem = 'Y' and T1.U_ITEMCODE = $[itemcode] and T1.U_WHSCODE = $[whscode]
left join
(
select T1.U_BASEENT, T1.U_BASELINE, T1.U_BASEOBJ, sum(T1.U_CMQTY) as U_CMQTY
from [@FT_TPPLAN] T0 inner join [@FT_TPPLAN1] T1 on T0.DocEntry = T1.DocEntry and T0.Canceled = 'N' and T0.Status = 'C' and T1.U_BASEENT <> 0
inner join OITM on OITM.ItemCode = T1.U_ITEMCODE and OITM.InvntItem = 'Y' and T1.U_ITEMCODE = $[itemcode] and T1.U_WHSCODE = $[whscode]
group by T1.U_BASEENT, T1.U_BASELINE, T1.U_BASEOBJ
) T00 on T00.U_BASEENT = T1.DocEntry and T00.U_BASELINE = T1.LineId and T00.U_BASEOBJ = T1.Object
group by T1.U_ITEMCODE, T1.U_WHSCODE
) T9 left join OITW on T9.U_ITEMCODE = OITW.ItemCode and T9.U_WHSCODE = OITW.WhsCode";

            return;
        }
    }
}
