using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    public class ft_Functions
    {
        public ft_Functions() { }

        public static int CheckSPNeeded(string type, string formtype, string docnum)
        {
            //return -1 if error
            //return 0 if not found
            //return > 0 if found

            try
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery($"Exec FT_CheckSPAppNeeded '{type}', '{formtype}', '{docnum}'");
                if (rs.RecordCount <= 0)
                {
                    if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    SAP.SBOApplication.MessageBox($"{type} {formtype} {docnum} FT_CheckCreditTerm error.", 1, "Ok", "", "");
                    return -1;
                }
                else
                {
                    if (rs.Fields.Item(0).Value.ToString() != "0")
                        return 1;
                }
                return 0;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("CheckCreditTerm " + ex.Message, 1, "Ok", "", "");
                return -1;
            }
        }

        public static int CheckCreditTerm(SAPbouiCOM.Form oForm, SAPbouiCOM.DBDataSource ods, SAPbouiCOM.DBDataSource ods1, ref string documentnum, ref string documentdate, ref string documentduedate, ref string errMsg)
        {
            //return -1 if error
            //return 0 if not found
            //return > 0 if found

            try
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string yyyy = "", MM = "", dd = "", cardcode = "";
                DateTime date = new DateTime();
                if (ods.TableName.Contains("@"))
                {
                    yyyy = ods.GetValue("U_docdate", 0).ToString().Substring(0, 4);
                    MM = ods.GetValue("U_docdate", 0).ToString().Substring(4, 2);
                    dd = ods.GetValue("U_docdate", 0).ToString().Substring(6, 2);
                    cardcode = ods.GetValue("U_cardcode", 0).ToString();
                }
                else if (oForm.TypeEx == "1250000100")
                {
                    yyyy = ods.GetValue("StartDate", 0).ToString().Substring(0, 4);
                    MM = ods.GetValue("StartDate", 0).ToString().Substring(4, 2);
                    dd = ods.GetValue("StartDate", 0).ToString().Substring(6, 2);
                    cardcode = ods.GetValue("BpCode", 0).ToString();
                }
                else
                {
                    yyyy = ods.GetValue("docdate", 0).ToString().Substring(0, 4);
                    MM = ods.GetValue("docdate", 0).ToString().Substring(4, 2);
                    dd = ods.GetValue("docdate", 0).ToString().Substring(6, 2);
                    cardcode = ods.GetValue("cardcode", 0).ToString();
                }

                date = DateTime.Parse(yyyy + "-" + MM + "-" + dd);
                //rs.DoQuery("select T0.* from oinv T0 inner join OCRD T1 on T0.cardcode = T1.cardcode where T0.cardcode='" + cardcode + "' and docstatus='O' and dateadd(day,isnull(T0.U_Grace,0),DATEADD (dd, -1, DATEADD(mm, DATEDIFF(mm, 0, docduedate) + 1, 0))) <= '" + date.ToString("yyyy-MM-dd") + "' ");
                rs.DoQuery($"exec FT_CheckCreditTerm '{cardcode}', '{date.ToString("yyyy-MM-dd")}'");
                if (rs.RecordCount <= 0)
                {
                    if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                    SAP.SBOApplication.MessageBox($"{cardcode} {date.ToString("yyyy-MM-dd")} FT_CheckCreditTerm error.", 1, "Ok", "", "");
                    return -1;
                }
                else
                {
                    if (rs.Fields.Item(0).Value.ToString() != "0")
                    {
                        documentnum = rs.Fields.Item(1).Value.ToString();
                        documentdate = rs.Fields.Item(2).Value.ToString();
                        documentduedate = rs.Fields.Item(3).Value.ToString();
                        return 1;
                    }

                }
                return 0;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("CheckCreditTerm " + ex.Message, 1, "Ok", "", "");
                return -1;
            }
        }

        public static int CheckCreditLimit(SAPbouiCOM.Form oForm, SAPbouiCOM.DBDataSource ods, SAPbouiCOM.DBDataSource ods1, ref string errMsg, ref string limitType, ref double different, ref double currentUsage, ref double temporaryLimit, ref double customerLimit)
        {
            //return -1 if error
            //return 0 if not found
            //return > 0 if found
            try
            {
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string mainCurncy = "", docCur = "", MasterDBName = "", sql = "", cardcode = "", newCardCode = "", yyyy = "", MM = "", dd = "";
                double docRate = 1, limit = 0, doctotal = 0, totalDue = 0;
                int cnt = 0;
                DateTime date = new DateTime();
           
                if (ods.TableName.Contains("@"))
                {
                    yyyy = ods.GetValue("U_docdate", 0).ToString().Substring(0, 4);
                    MM = ods.GetValue("U_docdate", 0).ToString().Substring(4, 2);
                    dd = ods.GetValue("U_docdate", 0).ToString().Substring(6, 2);
                    cardcode = ods.GetValue("U_cardcode", 0).ToString();
                    doctotal = 0;
                }
                else if (oForm.TypeEx == "1250000100")
                {
                    decimal qty = 0;
                    decimal price = 0;
                    decimal total = 0;
                    string temp = "";
                    string[] tempa;
                    SAPbouiCOM.Matrix grid = oForm.Items.Item("1250000045").Specific as SAPbouiCOM.Matrix;
                    yyyy = ods.GetValue("StartDate", 0).ToString().Substring(0, 4);
                    MM = ods.GetValue("StartDate", 0).ToString().Substring(4, 2);
                    dd = ods.GetValue("StartDate", 0).ToString().Substring(6, 2);
                    cardcode = ods.GetValue("BpCode", 0).ToString();
                    for (int x = 1; x < grid.RowCount; x++)
                    {
                        temp = ((SAPbouiCOM.EditText)grid.Columns.Item("1250000007").Cells.Item(x).Specific).Value;
                        decimal.TryParse(temp, out qty);
                        temp = ((SAPbouiCOM.EditText)grid.Columns.Item("1250000009").Cells.Item(x).Specific).Value;
                        tempa = temp.Split(' ');
                        decimal.TryParse(tempa[tempa.Length - 1], out price);

                        total += Math.Round(qty * price, 6, MidpointRounding.AwayFromZero);
                    }
                    doctotal = Convert.ToDouble(total);
                }
                else
                {
                    yyyy = ods.GetValue("docdate", 0).ToString().Substring(0, 4);
                    MM = ods.GetValue("docdate", 0).ToString().Substring(4, 2);
                    dd = ods.GetValue("docdate", 0).ToString().Substring(6, 2);
                    cardcode = ods.GetValue("cardcode", 0).ToString();
                    doctotal = double.Parse(ods.GetValue("doctotal", 0).ToString());
                }
                date = DateTime.Parse(yyyy + "-" + MM + "-" + dd);

                rs.DoQuery("select * from oadm");
                mainCurncy = rs.Fields.Item("mainCurncy").Value.ToString();
                MasterDBName = rs.Fields.Item("U_MDBName").Value.ToString();

                //newCardCode = cardcode.IndexOf("-") == -1 ? cardcode : cardcode.Substring(0, cardcode.IndexOf("-"));

                //if (docCur.ToUpper() != mainCurncy.ToUpper())
                //{
                //    doctotal = doctotal * docRate;
                //}

                if (MasterDBName != "")
                {
                    // ykw 20180417
                    string tablename = ods.TableName;
                    string temp = "Exec " + MasterDBName + "..FT_CheckCreditLimit '" + cardcode + "', '" + date.ToString("yyyy-MM-dd") + "', " + @doctotal.ToString() + ", '" + tablename + "', '" + SAP.SBOCompany.CompanyDB + "'";
                    rs.DoQuery(temp);
                    if (rs.RecordCount <= 0)
                    {
                        if (SAP.SBOCompany.InTransaction) SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                        SAP.SBOApplication.MessageBox($"CheckCreditLimit error {cardcode}. {MasterDBName}..FT_CheckCreditLimit", 1, "Ok", "", "");
                        return -1;
                    }
                    else
                    {
                        if (rs.Fields.Item(0).Value.ToString() != "0")
                        {
                            limitType = rs.Fields.Item(1).Value.ToString();
                            different = double.Parse(rs.Fields.Item(2).Value.ToString());
                            currentUsage = double.Parse(rs.Fields.Item(3).Value.ToString());
                            temporaryLimit = double.Parse(rs.Fields.Item(4).Value.ToString());
                            customerLimit = double.Parse(rs.Fields.Item(5).Value.ToString());
                            return 1;
                        }
                    }
                    return 0;
                }

                return 0;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("CheckCreditLimit " + ex.Message, 1, "Ok", "", "");
                return -1;
            }
        }

    }
}
