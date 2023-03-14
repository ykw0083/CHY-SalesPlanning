using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class UserForm_FT_GENSPERR
    {

        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        SAPbouiCOM.Item oItem = oForm.Items.Item("st_msg");


                        if (pVal.ItemUID == "cb_ARIV")
                        {
                            ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Sales Invoice generating...";
                            SalesInvoice("");
                        }
                        else if (pVal.ItemUID == "cb_RIV")
                        {
                            ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Reserve Invoice generating...";
                            ReserveSalesInvoice("");
                        }
                        else if (pVal.ItemUID == "cb_APIV")
                        {
                            ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Purchase Invoice generating...";
                            PurchaseInvoice("");
                        }
                        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Completed";
                        SAP.SBOApplication.MessageBox("Completed", 1, "Ok", "", "");

                        break;
                    default:
                        break;
                }

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void SalesInvoice(string docentry)
        {
            JEHeader jehdr;
            JEDetails jedtl;
            List<JEDetails> JEdtls;
            try
            {
                SAPbobsCOM.Recordset riv = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                DateTime date = new DateTime();
                string sql = "";
                string ivdocentry = "";
                if (docentry == "")
                {
                    sql = "select T0.docentry " +
                        " from oinv T0 " +
                        " left join [@FT_SPERRLOG] T99 on convert(nvarchar,T0.DocNum) = T99.U_ARInvNo" +
                        " where T0.CANCELED = 'N' and T0.isIns = 'N' " +
                        " and T99.docentry is null " +
                        " group by T0.docentry";
                }
                else
                {
                    sql = "select T0.docentry " +
                        " from oinv T0 " +
                        " where T0.CANCELED = 'N' and T0.isIns = 'N' " +
                        " and T0.docentry = " + docentry +
                        " group by T0.docentry";
                }
                riv.DoQuery(sql);
                if (riv.RecordCount > 0)
                {
                    riv.MoveFirst();

                    while (!riv.EoF)
                    {
                        ivdocentry = riv.Fields.Item("docentry").Value.ToString();

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
                        #region FT_SPERRLOG
                        sql = "select T0.docdate, T0.docnum, T4.U_Type, min(T3.docnum) as [DelNo] " +
                            " from oinv T0 inner join inv1 T1 on T0.docentry = T1.docentry " +
                            " inner join dln1 T2 on T1.baseentry = T2.docentry and T1.basetype = 15 " +
                            " and T1.BaseLine = T2.LineNum " +
                            " inner join odln T3 on T2.docentry = T3.docentry " +
                            " inner join ocrd T4 on T0.cardcode = T4.cardcode " +
                            " where T0.CANCELED = 'N' and T0.docentry = " + ivdocentry + " " +
                            " group by T0.docdate, T0.docnum, T4.U_Type";
                        rs.DoQuery(sql);
                        if (rs.RecordCount > 0)
                        {
                            rs.MoveFirst();

                            //different = double.Parse(rs.Fields.Item("U_Diff").Value.ToString());

                            //if (different != 0)
                            {
                                string bpType = "", cogsAcct = "";
                                double temp = 0;
                                double invtotaldiff = 0;
                                string productgroup = "";
                                string ChargeNo = "";
                                string delno = "";

                                date = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
                                jehdr = new JEHeader();
                                JEdtls = new List<JEDetails>();

                                jehdr.DocType = "ARIV";
                                jehdr.RefDate = date;
                                jehdr.U_ARInvNo = rs.Fields.Item("docnum").Value.ToString();
                                jehdr.U_DelNo = rs.Fields.Item("DelNo").Value.ToString();
                                jehdr.Memo = "DO Charge Out";

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
                                    " where T0.docentry = " + ivdocentry +
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
                                                    temp = Math.Round(temp, 2, MidpointRounding.AwayFromZero);
                                                    jedtl = new JEDetails();

                                                    jedtl.AccountCode = cogsAcct;
                                                    if (temp > 0)
                                                        jedtl.Credit = temp;
                                                    else if (temp < 0)
                                                        jedtl.Debit = -temp;
                                                    if (!string.IsNullOrEmpty(productgroup))
                                                        jedtl.CostingCode = productgroup;
                                                    JEdtls.Add(jedtl);

                                                    invtotaldiff = invtotaldiff + temp;

                                                    rs1.MoveNext();
                                                }
                                            }

                                        }
                                        rs.MoveNext();
                                    }

                                    if (invtotaldiff != 0)
                                    {
                                        jedtl = new JEDetails();

                                        jedtl.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                        if (invtotaldiff > 0)
                                            jedtl.Debit = invtotaldiff;
                                        else if (invtotaldiff < 0)
                                            jedtl.Credit = -invtotaldiff;

                                        JEdtls.Add(jedtl);
                                    }

                                }
                                if (invtotaldiff != 0)
                                {
                                    jehdr.ErrMsg = "From GenSPERR";
                                    jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmss");
                                    ObjectFunctions.ErrorLog(jehdr, JEdtls);
                                }
                            }
                        }
                        #endregion


                        riv.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                jehdr = new JEHeader();
                JEdtls = new List<JEDetails>();
                jehdr.ErrMsg = "From GenSPERR Catch " + ex.Message;
                jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                ObjectFunctions.ErrorLog(jehdr, JEdtls);
                SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
            }
        }

        public static void ReserveSalesInvoice(string docentry)
        {
            JEHeader jehdr;
            JEDetails jedtl;
            List<JEDetails> JEdtls;
            try
            {
                SAPbobsCOM.Recordset riv = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


                string ivdocentry = "";

                if (docentry == "")
                {
                    riv.DoQuery("Select T0.docentry " +
                    " from oinv T0 " +
                    " left join [@FT_SPERRLOG] T99 on convert(nvarchar,T0.DocNum) = T99.U_RInvNo" +
                    " where T0.isIns = 'Y' " +
                    " and T0.CANCELED = 'N' and T99.docentry is null " +
                    " group by T0.docentry");
                }
                else
                {
                    riv.DoQuery("Select T0.docentry " +
                    " from oinv T0 " +
                    " where T0.isIns = 'Y' " +
                    " and T0.CANCELED = 'N' and T0.docentry = " + docentry +
                    " group by T0.docentry");
                }
                if (riv.RecordCount > 0)
                {
                    riv.MoveFirst();

                    while (!riv.EoF)
                    {
                        ivdocentry = riv.Fields.Item("docentry").Value.ToString();

                        rs.DoQuery("Select T0.docdate,T0.docnum, T1.basetype,T1.itemcode, T1.quantity as [quantity], T1.whscode, T2.U_Type, T1.StockPrice from oinv T0 inner join inv1 T1 on T0.docentry = T1.docentry " +
                            " inner join ocrd T2 on T0.cardcode = T2.cardcode " +
                            " where T0.CANCELED = 'N' and T0.docentry =" + ivdocentry +
                            " order by T1.linenum");// + " group by  T0.docdate,T1.basetype, T0.docnum,T1.itemcode,T1.whscode, T2.U_Type ");
                        if (rs.RecordCount > 0)
                        {
                            if (rs.Fields.Item("basetype").Value.ToString() == "17")
                            {
                                rs.MoveFirst();
                                string productgroup = "";
                                double line_total = 0;
                                double total = 0;
                                string bpType = rs.Fields.Item("U_Type").Value.ToString();
                                string cogsAcct = "";
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
                                jehdr = new JEHeader();
                                JEdtls = new List<JEDetails>();

                                jehdr.RefDate = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
                                jehdr.U_RInvNo = rs.Fields.Item("docnum").Value.ToString();
                                jehdr.Memo = "Provision for Reserved Invoice";
                                jehdr.DocType = "ARRIV";


                                int currentline = 0;

                                while (!rs.EoF)
                                {
                                    rs1.DoQuery("select T0.avgprice, T1.U_CostCenter from oitm T0 inner join oitb T1 on T0.itmsgrpcod = T1.itmsgrpcod where isnull(T0.invntitem,'N') = 'Y' and T0.itemcode='" + rs.Fields.Item("itemcode").Value.ToString() + "' ");
                                    if (rs1.RecordCount > 0)
                                    {
                                        productgroup = rs1.Fields.Item("U_CostCenter").Value.ToString();
                                        //line_total = double.Parse(rs.Fields.Item("quantity").Value.ToString()) * double.Parse(rs1.Fields.Item("avgprice").Value.ToString());
                                        line_total = double.Parse(rs.Fields.Item("quantity").Value.ToString()) * double.Parse(rs.Fields.Item("StockPrice").Value.ToString());
                                        line_total = Math.Round(line_total, 2, MidpointRounding.AwayFromZero);
                                        if (line_total > 0)
                                        {
                                            jedtl = new JEDetails();
                                            jedtl.AccountCode = cogsAcct;
                                            jedtl.Debit = line_total;
                                            if (!string.IsNullOrEmpty(productgroup))
                                                jedtl.CostingCode = productgroup;
                                            JEdtls.Add(jedtl);

                                            total = total + line_total;
                                        }
                                    }
                                    rs.MoveNext();
                                }

                                jedtl = new JEDetails();
                                jedtl.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                jedtl.Credit = total;
                                JEdtls.Add(jedtl);

                                if (total > 0)
                                {
                                    jehdr.ErrMsg = "From GenSPERR";
                                    jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmss");
                                    ObjectFunctions.ErrorLog(jehdr, JEdtls);

                                }
                            }
                        }

                        riv.MoveNext();
                    }
                }
            }
            catch (Exception ex)
            {
                jehdr = new JEHeader();
                JEdtls = new List<JEDetails>();
                jehdr.ErrMsg = "From GenSPERR Catch " + ex.Message;
                jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                ObjectFunctions.ErrorLog(jehdr, JEdtls);
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
        }

        public static void PurchaseInvoice(string docentry)
        {
            JEHeader jehdr;
            JEDetails jedtl;
            List<JEDetails> JEdtls;

            SAPbobsCOM.Recordset riv = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);


            string ivdocentry = "";

            try
            {

                if (docentry == "")
                {
                    riv.DoQuery("select T0.docentry " +
                    " from OPCH T0 " +
                    " left join [@FT_SPERRLOG] T99 on convert(nvarchar,T0.DocNum) = T99.U_APINVNo" +
                    " where T0.CANCELED = 'N' and T99.docentry is null " +
                    " group by T0.docentry");
                }
                else
                {
                    riv.DoQuery("select T0.docentry " +
                    " from OPCH T0 " +
                    " where T0.CANCELED = 'N' and T0.docentry = " + docentry +
                    " group by T0.docentry");
                }
                if (riv.RecordCount > 0)
                {
                    riv.MoveFirst();

                    while (!riv.EoF)
                    {

                        ivdocentry = riv.Fields.Item("docentry").Value.ToString();

                        rs.DoQuery("select T4.U_PRebPer, T4.U_PRebAmt,T0.docnum, T0.docdate,T5.U_Type, T3.quantity, T3.LineNum from OPCH T0 " +
                        " inner join PCH1 T1 on T0.docentry = T1.docentry  " +
                        " inner join PDN1 T2 on T1.BaseEntry = T2.DocEntry and T1.BaseType = 20 " +
                        " inner join POR1 T3 on T2.BaseEntry = T3.DocEntry and T2.BaseType = 22 " +
                        " inner join OPOR T4 on T3.docentry = T4.docentry " +
                        " inner join ocrd T5 on T0.cardcode = T5.cardcode " +
                        " where T0.CANCELED = 'N' and T0.docentry = " + ivdocentry + " group by T4.U_PRebPer, T4.U_PRebAmt,T0.docnum, T0.docdate,T5.U_Type,T3.quantity, T3.LineNum");
                        if (rs.RecordCount > 0)
                        {
                            double linetotal = 0, rebAmt = 0, totalQty = 0, rebTotal = 0, rebDiff = 0, lineQty = 0, finalAmt = 0;
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
                                   " where T0.docentry = " + ivdocentry);
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
                                jehdr = new JEHeader();
                                JEdtls = new List<JEDetails>();
                                jehdr.RefDate = date;
                                //oJE.TaxDate = date;
                                jehdr.U_APINVNo = ap_docnum;

                                jehdr.DocType = "APIV";
                                jehdr.Memo = "Provision for Purchase Discount/Rebate";

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
                                    jedtl = new JEDetails();
                                    jedtl.AccountCode = "350700";// "250700";// "150400"; Provision for Purchase Discount/Rebate
                                    jedtl.Debit = finalAmt;
                                    productgroup = rs.Fields.Item("ocrcode").Value.ToString();
                                    if (!string.IsNullOrEmpty(productgroup))
                                        jedtl.CostingCode = productgroup;

                                    JEdtls.Add(jedtl);

                                    jedtl = new JEDetails();
                                    jedtl.AccountCode = purchaseAcct;
                                    jedtl.Credit = finalAmt;
                                    if (!string.IsNullOrEmpty(productgroup))
                                        jedtl.CostingCode = productgroup;

                                    JEdtls.Add(jedtl);

                                    rs.MoveNext();
                                }

                                if (finalAmt != 0)
                                {
                                    jehdr.ErrMsg = "From GenSPERR";
                                    jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmss");
                                    ObjectFunctions.ErrorLog(jehdr, JEdtls);
                                }
                            }
                        }

                        riv.MoveNext();

                    }
                }
            }
            catch (Exception ex)
            {
                jehdr = new JEHeader();
                JEdtls = new List<JEDetails>();
                jehdr.ErrMsg = "From GenSPERR Catch " + ex.Message;
                jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmssfff");
                ObjectFunctions.ErrorLog(jehdr, JEdtls);
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
        }
    }
}
