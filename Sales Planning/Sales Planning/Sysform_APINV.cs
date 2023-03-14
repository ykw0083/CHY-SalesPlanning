using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.CHY
{
    class Sysform_APINV
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
                    oEdit.DataBind.SetBound(true, "OPCH", "U_ADDONIND");

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
                        //            int docentry = int.Parse(oForm.DataSources.DBDataSources.Item("PCH1").GetValue("DOCENTRY", pVal.Row - 1).ToString());
                        //            int linenum = int.Parse(oForm.DataSources.DBDataSources.Item("PCH1").GetValue("LINENUM", pVal.Row - 1).ToString());
                        //            InitForm.SDM(oForm.UniqueID, docentry, linenum, pVal.Row, "PCH1", "38");
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
        public static void processDataEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            JEHeader jehdr = new JEHeader();
            JEDetails jedtl;
            List<JEDetails> JEdtls = new List<JEDetails>();

            try
            {
                switch (BusinessObjectInfo.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD:
                        if (!BusinessObjectInfo.ActionSuccess) return;
                        int retcode = 0;

                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);

                        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
                        rs.DoQuery("select T4.U_PRebPer, T4.U_PRebAmt,T0.docnum, T0.docdate,T5.U_Type, T3.quantity, T3.LineNum from OPCH T0 " +
                            " inner join PCH1 T1 on T0.docentry = T1.docentry  " +
                            " inner join PDN1 T2 on T1.BaseEntry = T2.DocEntry and T1.BaseType = 20 " +
                            " inner join POR1 T3 on T2.BaseEntry = T3.DocEntry and T2.BaseType = 22 " +
                            " inner join OPOR T4 on T3.docentry = T4.docentry " +
                            " inner join ocrd T5 on T0.cardcode = T5.cardcode " +
                            " where T0.docentry = " + xmlDoc.InnerText + " group by T4.U_PRebPer, T4.U_PRebAmt,T0.docnum, T0.docdate,T5.U_Type,T3.quantity, T3.LineNum");
                        if (rs.RecordCount > 0)
                        {
                            double linetotal = 0, rebAmt = 0, totalQty = 0, rebTotal = 0,rebDiff=0, lineQty=0, finalAmt =0;
                            int cnt = 0;
                            string bpType = "", purchaseAcct = "", ap_docnum = "";
                            string productgroup = "";
                            rebAmt = double.Parse(rs.Fields.Item("U_PRebAmt").Value.ToString());
                            ap_docnum = rs.Fields.Item("docnum").Value.ToString();
                            DateTime date = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());

                            if (rebAmt > 0)
                            {
                                bpType = rs.Fields.Item("U_Type").Value.ToString();
                                switch (bpType.ToUpper())
                                {
                                    case "LOCAL":
                                        purchaseAcct = "520001";
                                        break;
                                    case "OVERSEA":
                                        purchaseAcct = "520100";
                                        break;
                                    case "INTER-COMPANY":
                                        purchaseAcct = "520200";
                                        break;
                                }
                                if (purchaseAcct == "")
                                {
                                    SAP.SBOApplication.MessageBox("Unable to identify Purchase Discount G/L", 1, "Ok", "", "");
                                    return;
                                }
                                // ykw 20181002 start
                                //rs.MoveFirst();
                                //while (!rs.EoF)
                                //{
                                //    lineQty = double.Parse(rs.Fields.Item("quantity").Value.ToString());
                                //    totalQty += lineQty;
                                //    rs.MoveNext();
                                //}
                                // ykw 20181002 end

                                //rs.DoQuery("select sum(linetotal) as [linetotal], ocrcode from pch1 where docentry = " + xmlDoc.InnerText + " group by ocrcode");
                                //rs.DoQuery("select sum(T0.quantity * (isnull(T1.U_Attribute17,1) / 1000)) as [lineQty], ocrcode from pch1 T0 inner join oitm T1 " +
                                //    " on T0.itemcode = T1.itemcode where T0.docentry = " + xmlDoc.InnerText + " group by T0.ocrcode order by sum(T0.quantity / isnull(T1.U_Attribute17,1))");
                                rs.DoQuery("select T0.quantity as [lineQty], T5.U_CostCenter as ocrcode " +
                                   " from pch1 T0 inner join oitm T1 " +
                                   " on T0.itemcode = T1.itemcode " +
                                   " inner join oitb T5 on T1.itmsgrpcod = T5.itmsgrpcod " +
                                   " where T0.docentry = " + xmlDoc.InnerText);
                                rs.MoveFirst();
                                while (!rs.EoF)
                                {
                                    lineQty = double.Parse(rs.Fields.Item("lineQty").Value.ToString());
                                    totalQty += lineQty;
                                    rs.MoveNext();
                                }
                                // ykw 20181002 start
                                //rs.MoveFirst();
                                //while (!rs.EoF)
                                //{
                                //    lineQty = double.Parse(rs.Fields.Item("lineQty").Value.ToString());
                                //    rebTotal = rebTotal + Math.Round(lineQty / totalQty * rebAmt,2);
                                //    rs.MoveNext();
                                //}
                                //rebDiff = rebTotal - rebAmt;
                                // ykw 20181002 end

                                oJE.ReferenceDate = date;
                                //oJE.TaxDate = date;
                                oJE.UserFields.Fields.Item("U_APINVNo").Value = ap_docnum;

                                jehdr.DocType = "APIV";
                                jehdr.RefDate = oJE.ReferenceDate;
                                jehdr.U_APINVNo = oJE.UserFields.Fields.Item("U_APINVNo").Value.ToString();
                                jehdr.Memo = oJE.Memo;

                                //oJE.Lines.AccountCode = "250700";// "150400";
                                //oJE.Lines.Debit = rebAmt;
                                rs.MoveFirst();
                                while (!rs.EoF)
                                {
                                    lineQty = double.Parse(rs.Fields.Item("lineQty").Value.ToString());
                                    finalAmt = Math.Round(lineQty / totalQty * rebAmt, 2, MidpointRounding.AwayFromZero);
                                    //if (cnt > 1)
                                    //{
                                    //    oJE.Lines.Add();
                                    //    oJE.Lines.SetCurrentLine(cnt);
                                    //    cnt++;
                                    //}
                                    //else
                                    //{
                                    //    if (rebDiff > 0)
                                    //    {
                                    //        finalAmt = finalAmt - rebDiff;
                                    //    }
                                    //    else
                                    //    {
                                    //        finalAmt = finalAmt + rebDiff;
                                    //    }
                                    //}
                                    cnt++;
                                    if (cnt > 1)
                                    {
                                        oJE.Lines.Add();
                                        oJE.Lines.SetCurrentLine(cnt - 1);
                                    }
                                    oJE.Lines.AccountCode = "350700";// "250700";// "150400"; Provision for Purchase Discount/Rebate
                                    oJE.Lines.Debit = finalAmt;
                                    productgroup = rs.Fields.Item("ocrcode").Value.ToString();
                                    if (!string.IsNullOrEmpty(productgroup))
                                        oJE.Lines.CostingCode = productgroup;

                                    jedtl = new JEDetails();
                                    jedtl.AccountCode = oJE.Lines.AccountCode;
                                    jedtl.CostingCode = oJE.Lines.CostingCode;
                                    jedtl.Debit = oJE.Lines.Debit;
                                    JEdtls.Add(jedtl);

                                    cnt++;
                                    if (cnt > 1)
                                    {
                                        oJE.Lines.Add();
                                        oJE.Lines.SetCurrentLine(cnt - 1);
                                    }
                                    oJE.Lines.AccountCode = purchaseAcct;
                                    oJE.Lines.Credit = finalAmt;
                                    if (!string.IsNullOrEmpty(productgroup))
                                        oJE.Lines.CostingCode = productgroup;

                                    jedtl = new JEDetails();
                                    jedtl.AccountCode = oJE.Lines.AccountCode;
                                    jedtl.CostingCode = oJE.Lines.CostingCode;
                                    jedtl.Credit = oJE.Lines.Credit;
                                    JEdtls.Add(jedtl);

                                    rs.MoveNext();
                                }

                                retcode = oJE.Add();
                                if (retcode != 0)
                                {
                                    jehdr.ErrMsg = SAP.SBOCompany.GetLastErrorDescription();
                                    jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmss");
                                    ObjectFunctions.ErrorLog(jehdr, JEdtls);
                                    //if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                    SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                    //FT_ADDON.SAP.SBOApplication.StatusBar.SetText(SAP.SBOCompany.GetLastErrorDescription(), SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                    return;
                                }
                                xmlDoc = null;
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                jehdr.ErrMsg = "Data Event After " + ex.Message;
                jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                ObjectFunctions.ErrorLog(jehdr, JEdtls);
                //if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
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
        //                int retcode = 0;

        //                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
        //                xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);
        //                rs.DoQuery("select T4.U_PRebPer, T4.U_PRebAmt from OPCH T0 " +
        //                    " inner join PCH1 T1 on T0.docentry = T1.docentry  " +
        //                    " inner join PDN1 T2 on T1.BaseEntry = T2.DocEntry and T1.BaseType = 20 " +
        //                    " inner join POR1 T3 on T2.BaseEntry = T3.DocEntry and T2.BaseType = 22 " +
        //                    " inner join OPOR T4 on T3.docentry = T4.docentry " + 
        //                    " where T0.docentry = " + xmlDoc.InnerText + " group by T4.U_PRebPer, T4.U_PRebAmt ");
        //                if (rs.RecordCount > 0)
        //                {
        //                    double rebAmt = 0, total = 0, linetotal = 0, rebTotal = 0 ;
        //                    int cnt = 1;
        //                    string bpType = "", purchaseAcct = "";
        //                    rebAmt = double.Parse(rs.Fields.Item("U_PRebAmt").Value.ToString());

        //                    if (rebAmt > 0)
        //                    {
        //                        rs.DoQuery("select T0.docnum, T0.docdate, T0.doctotal, T0.vatsum, T1.ocrcode, T2.U_Type, sum(T1.linetotal) as [linetotal] " +
        //                            " from OPCH T0 " +
        //                            " INNER JOIN PCH1 T1 ON T0.DOCENTRY = T1.DOCENTRY " +
        //                            " INNER JOIN OCRD T2 ON T0.CARDCODE = T2.CARDCODE " + 
        //                            " WHERE T0.DOCENTRY= " + xmlDoc.InnerText + " group by T0.docnum, T0.docdate, T0.doctotal, T0.vatsum, T1.ocrcode,T2.U_Type");
        //                        if (rs.RecordCount > 0)
        //                        {
        //                            rs.MoveFirst();
        //                            bpType = rs.Fields.Item("U_Type").Value.ToString();
        //                            total = double.Parse(rs.Fields.Item("doctotal").Value.ToString()) - double.Parse(rs.Fields.Item("vatsum").Value.ToString());
        //                            switch (bpType.ToUpper())
        //                            {
        //                                case "IMPORT":
        //                                    purchaseAcct = "520001";
        //                                    break;
        //                                case "INTER COMPANY":
        //                                    purchaseAcct = "520100";
        //                                    break;
        //                                case "LOCAL":
        //                                    purchaseAcct = "520200";
        //                                    break;
        //                            }
        //                            if (purchaseAcct == "")
        //                            {
        //                                SAP.SBOApplication.MessageBox("Unable to identify Purchase Discount G/L", 1, "Ok", "", "");
        //                                return;
        //                            }
        //                            SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
        //                            oJE.TaxDate = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
        //                            oJE.UserFields.Fields.Item("U_APInvNo").Value = rs.Fields.Item("docnum").Value.ToString();
        //                            while (!rs.EoF)
        //                            {
        //                                linetotal = double.Parse(rs.Fields.Item("linetotal").Value.ToString());
        //                                rebTotal += (linetotal / total) * rebAmt;
        //                                rs.MoveNext();
        //                            }
        //                            oJE.Lines.AccountCode = "250700";
        //                            oJE.Lines.Debit = rebTotal;
        //                            rs.MoveFirst();
        //                            while (!rs.EoF)
        //                            {
        //                                oJE.Lines.Add();
        //                                oJE.Lines.SetCurrentLine(cnt);
        //                                linetotal = double.Parse(rs.Fields.Item("linetotal").Value.ToString());

        //                                oJE.Lines.AccountCode = purchaseAcct;
        //                                oJE.Lines.Credit = (linetotal / total) * rebAmt;
        //                                oJE.Lines.CostingCode = rs.Fields.Item("ocrcode").Value.ToString();
        //                                cnt++;
        //                                rs.MoveNext();
        //                            }
        //                            retcode = oJE.Add();
        //                            if (retcode != 0)
        //                            {
        //                                SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
        //                            }
        //                            xmlDoc = null;
        //                        }
        //                    }
        //                }
        //                break;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
        //        SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
        //    }
        //}
    }
}
