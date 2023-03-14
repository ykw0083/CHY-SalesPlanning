using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.AYS
{
    class Sysform_ARReservedInvoice
    {
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

                        int retcode = 0;
                        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        SAPbobsCOM.Recordset rs1 = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                        System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                        xmlDoc.LoadXml(BusinessObjectInfo.ObjectKey);

                        UserForm_FT_GENSPERR.ReserveSalesInvoice(xmlDoc.InnerText);

                        rs.DoQuery("Select T0.docdate,T0.docnum, T1.basetype,T1.itemcode, T1.quantity as [quantity], T1.whscode, T2.U_Type from oinv T0 inner join inv1 T1 on T0.docentry = T1.docentry " +
                            " inner join ocrd T2 on T0.cardcode = T2.cardcode " +
                            " where T0.CANCELED = 'N' and T0.docentry =" + xmlDoc.InnerText +
                            " order by T1.linenum");// + " group by  T0.docdate,T1.basetype, T0.docnum,T1.itemcode,T1.whscode, T2.U_Type ");
                        if (rs.RecordCount > 0)
                        {
                            if (rs.Fields.Item("basetype").Value.ToString() == "17")
                            {
                                rs.MoveFirst();
                                docnum = rs.Fields.Item("docnum").Value.ToString();
                                SAPbobsCOM.JournalEntries oJE = (SAPbobsCOM.JournalEntries)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oJournalEntries);
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
                                oJE.ReferenceDate = DateTime.Parse(rs.Fields.Item("docdate").Value.ToString());
                                oJE.UserFields.Fields.Item("U_RInvNo").Value = docnum;
                                oJE.Memo = "Provision for Reserved Invoice";

                                //jehdr.DocType = "ARRIV";
                                //jehdr.RefDate = oJE.ReferenceDate;
                                //jehdr.U_RInvNo = oJE.UserFields.Fields.Item("U_RInvNo").Value.ToString();
                                //jehdr.Memo = oJE.Memo;

                                int currentline = 0;

                                string productgroup = "";
                                double line_total = 0;
                                double total = 0;

                                while (!rs.EoF)
                                {
                                    rs1.DoQuery("select T0.avgprice, T1.U_CostCenter from oitm T0 inner join oitb T1 on T0.itmsgrpcod = T1.itmsgrpcod where isnull(T0.invntitem,'N') = 'Y' and T0.itemcode='" + rs.Fields.Item("itemcode").Value.ToString() + "' ");
                                    if (rs1.RecordCount > 0)
                                    {
                                        productgroup = rs1.Fields.Item("U_CostCenter").Value.ToString();
                                        line_total = double.Parse(rs.Fields.Item("quantity").Value.ToString()) * double.Parse(rs1.Fields.Item("avgprice").Value.ToString());
                                        line_total = Math.Round(line_total, 2);
                                        if (line_total > 0)
                                        {
                                            currentline++;
                                            if (currentline > 1)
                                            {
                                                oJE.Lines.Add();
                                                oJE.Lines.SetCurrentLine(currentline - 1);
                                            }
                                            oJE.Lines.AccountCode = cogsAcct;
                                            oJE.Lines.Debit = line_total;
                                            if (!string.IsNullOrEmpty(productgroup))
                                                oJE.Lines.CostingCode = productgroup;

                                            //jedtl = new JEDetails();
                                            //jedtl.AccountCode = oJE.Lines.AccountCode;
                                            //jedtl.CostingCode = oJE.Lines.CostingCode;
                                            //jedtl.Debit = oJE.Lines.Debit;
                                            //JEdtls.Add(jedtl);

                                            total = total + line_total;
                                        }
                                    }
                                    rs.MoveNext();
                                }

                                currentline++;
                                if (currentline > 1)
                                {
                                    oJE.Lines.Add();
                                    oJE.Lines.SetCurrentLine(currentline - 1);
                                }
                                oJE.Lines.AccountCode = "150500";// "150400"; Provision for Cost of Goods Sold
                                oJE.Lines.Credit = total;

                                //jedtl = new JEDetails();
                                //jedtl.AccountCode = oJE.Lines.AccountCode;
                                //jedtl.Credit = oJE.Lines.Credit;
                                //JEdtls.Add(jedtl);

                                if (total > 0)
                                {
                                    retcode = oJE.Add();
                                    if (retcode != 0)
                                    {
                                        //jehdr.ErrMsg = SAP.SBOCompany.GetLastErrorDescription();
                                        //jehdr.ErrDT = DateTime.Now.ToString("yyyyMMddHHmmss");
                                        //ObjectFunctions.ErrorLog(jehdr, JEdtls);
                                        //if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                                        rs.DoQuery("update [@FT_SPERRLOG] set U_ErrMsg = '" + SAP.SBOCompany.GetLastErrorDescription() + "' where U_RInvNo = '" + docnum + "'");
                                        SAP.SBOApplication.MessageBox(SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                                    }
                                    rs.DoQuery("update [@FT_SPERRLOG] set Status = 'C', Canceled = 'Y' where U_RInvNo = '" + docnum + "'");

                                }
                                xmlDoc = null;
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
                rs0.DoQuery("update [@FT_SPERRLOG] set U_ErrMsg = '" + ex.Message + "' where U_RInvNo = '" + docnum + "'");
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
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

    }
}
