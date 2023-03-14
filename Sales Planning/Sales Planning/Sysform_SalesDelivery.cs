using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.CHY
{
    class Sysform_SalesDelivery
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
                    SAPbouiCOM.UserDataSource uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                    uds.Value = "";
                    //SAPbouiCOM.Item oItem;
                    //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    //((SAPbouiCOM.EditText)oItem.Specific).String = "";
                    //oItem.Left = 160;
                    //oItem.Width = 10;
                    //oItem.Top = oForm.Height - 60;
                    //oItem.Height = 20;
                    ////oItem.Enabled = false;
                    ////oItem.Visible = false;
                    //oForm.Items.Item("4").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    //oForm.Items.Item("FUID").Visible = false;
                }
                if (oForm.DataSources.UserDataSources.Count > 0)
                {
                    if (oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString() != "")
                    {
                        string FUID = oForm.DataSources.UserDataSources.Item("cfluid").Value.ToString();
                        SAPbouiCOM.Form oSForm = SAP.SBOApplication.Forms.Item(FUID);
                        oSForm.Select();
                        BubbleEvent = false;
                    }
                }
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
                        if (pVal.ItemUID == "38")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.Row >= 0 && pVal.ColUID == "256")
                                {
                                    int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("DLN1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                                    int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("DLN1").GetValue("LINENUM", pVal.Row - 1).ToString());

                                    InitForm.TEXT(oForm.UniqueID, docentry, linenum, pVal.Row, "DLN1", "38", oForm.DataSources.DBDataSources.Item("DLN1").GetValue("TEXT", pVal.Row - 1).ToString());
                                    BubbleEvent = false;
                                }
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "38")
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (pVal.Row >= 0)
                                {
                                    int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("DLN1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                                    int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("DLN1").GetValue("LINENUM", pVal.Row - 1).ToString());
                                    string bookno = oForm.DataSources.DBDataSources.Item("DLN1").GetValue("U_BookNo", pVal.Row - 1).Trim();
                                    if (pVal.ColUID == "U_BookNo")
                                    {
                                        InitForm.CONM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APDOC", "U_CONNO,U_BookNo", bookno);
                                    }
                                    else if (pVal.ColUID == "U_ProdType")
                                    {
                                        InitForm.DOPTM(oForm.UniqueID, docentry, linenum, pVal.Row, "FT_APDOPT", "");
                                    }
                                    else if (pVal.Row >= 0 && pVal.ColUID == "256")
                                    {
                                    //    InitForm.TEXT(oForm.UniqueID, docentry, linenum, pVal.Row, "INV1", "38", oForm.DataSources.DBDataSources.Item("INV1").GetValue("TEXT", pVal.Row - 1).ToString());
                                    }
                                    else
                                    {
                                        InitForm.SDM(oForm.UniqueID, docentry, linenum, pVal.Row, "DLN1", "38");
                                    }
                                }
                            }
                        }
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
        public static void processDataEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        if (!BusinessObjectInfo.ActionSuccess) break;
                        
                        //int retcode = 0;

                        //SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        //SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        //System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        //xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                        //rs.DoQuery("Select T0.docnum, T0.docdate,T1.basetype,T1.baseentry, T1.itemcode, sum(T1.quantity) as [quantity], T1.whscode, T2.U_Type from ODLN T0 inner join DLN1 T1 on T0.docentry = T1.docentry " +
                        //    " inner join ocrd T2 on T0.cardcode = T2.cardcode " +
                        //    " where T0.docentry =" + xmlDoc.InnerText + " group by  T0.docnum, T0.docdate,T1.basetype, T1.baseentry,T1.itemcode,T1.whscode, T2.U_Type ");
                        //if (rs.RecordCount > 0)
                        //{
                        //    if (rs.Fields.Item("basetype").Value.ToString() == "13")
                        //    {
                        //        double total = 0, RInv_Total = 0;
                        //        string baseentry = "", RInv_docnum = "", docnum = "";
                        //        DateTime date = new DateTime();
                        //        rs.MoveFirst();
                        //        baseentry = rs.Fields.Item("baseentry").Value.ToString();
                        //        docnum = rs.Fields.Item("docnum").Value.ToString();
                        //        if (!SAP.SBOCompany.InTransaction) SAP.SBOCompany.StartTransaction();
                        //        SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        //        date = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
                        //        string bpType = rs.Fields.Item("U_Type").Value.ToString();
                        //        string cogsAcct = "";

                        //        switch (bpType.ToUpper())
                        //        {
                        //            case "LOCAL":
                        //                cogsAcct = "500001";
                        //                break;
                        //            case "OVERSEA":
                        //                cogsAcct = "500100";
                        //                break;
                        //            case "INTER-COMPANY":
                        //                cogsAcct = "500200";
                        //                break;
                        //        }

                        //        oJE.TaxDate = date;
                        //        oJE.UserFields.Fields.Item("U_DelNo").Value = docnum;
                        //        while (!rs.EoF)
                        //        {
                        //            rs1.DoQuery("select avgprice from oitm where itemcode='" + rs.Fields.Item("itemcode").Value.ToString() + "' ");
                        //            if (rs1.RecordCount > 0)
                        //            {
                        //                total += double.Parse(rs.Fields.Item("quantity").Value.ToString()) * double.Parse(rs1.Fields.Item("avgprice").Value.ToString());
                        //            }
                        //            rs.MoveNext();
                        //        }

                        //        oJE.Lines.AccountCode = "150500";
                        //        oJE.Lines.Debit = total;
                        //        oJE.Lines.Add();
                        //        oJE.Lines.SetCurrentLine(1);
                        //        oJE.Lines.AccountCode = cogsAcct;// "500300";
                        //        oJE.Lines.Credit = total;

                        //        retcode = oJE.Add();
                        //        xmlDoc = null;

                        //        rs.DoQuery("select * from oinv where docentry=" + baseentry);
                        //        if (rs.RecordCount > 0)
                        //        {
                        //            RInv_docnum = rs.Fields.Item("docnum").Value.ToString();

                        //            rs.DoQuery("select sum(debit) as [total] from ojdt T0 inner join JDT1 T1 on T0.transid = T1.transid where U_RInvNo='" + RInv_docnum + "'");
                        //            if (rs.RecordCount > 0)
                        //            {
                        //                RInv_Total = double.Parse(rs.Fields.Item("total").Value.ToString());
                        //                double different = RInv_Total - total;
                        //                if (different > 0)
                        //                {
                        //                    SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        //                    oJE1.TaxDate = date;
                        //                    oJE.UserFields.Fields.Item("U_DelNo").Value = docnum;

                        //                    oJE1.Lines.AccountCode = "530700";
                        //                    oJE1.Lines.Debit = different;
                        //                    oJE1.Lines.Add();
                        //                    oJE1.Lines.SetCurrentLine(1);
                        //                    oJE1.Lines.AccountCode = cogsAcct;// "500300";
                        //                    oJE1.Lines.Credit = different;

                        //                    retcode = oJE1.Add();
                        //                }
                        //                else if (different < 0)
                        //                {
                        //                    SAPbobsCOM.JournalEntries oJE1 = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        //                    oJE1.TaxDate = date;
                        //                    oJE.UserFields.Fields.Item("U_DelNo").Value = docnum;

                        //                    oJE1.Lines.AccountCode = cogsAcct;// "500300";
                        //                    oJE1.Lines.Debit = different;
                        //                    oJE1.Lines.Add();
                        //                    oJE1.Lines.SetCurrentLine(1);
                        //                    oJE1.Lines.AccountCode = "530700";
                        //                    oJE1.Lines.Credit = different;

                        //                    retcode = oJE1.Add();
                        //                }
                        //            }
                        //        }
                        //        if (retcode != 0)
                        //        {
                        //            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        //            FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                        //            return;
                        //        }
                        //        else
                        //        {
                        //            if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                        //        }
                        //    }

                        //}
                        break;
                }
            }
            catch (Exception ex)
            {
                if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
