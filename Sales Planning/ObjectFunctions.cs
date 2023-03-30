using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON
{
    class JEHeader
    {
        public string ErrDT = "";
        public string ErrMsg = "";
        public DateTime RefDate;
        public string Memo = "";
        public string DocType = "";
        public string U_RInvNo = "";
        public string U_ARInvNo = "";
        public string U_DelNo = "";
        public string U_APINVNo = "";

    }
    class JEDetails
    {
        public string AccountCode = "";
        public string CostingCode = "";
        public double Debit = 0;
        public double Credit = 0;
    }
    class ObjectFunctions
    {
        public static void ErrorLog(JEHeader JEhdr, List<JEDetails> JEdtls)
        {
            try
            {
                SAPbobsCOM.GeneralService oGeneralService = (SAPbobsCOM.GeneralService)SAP.SBOCompany.GetCompanyService().GetGeneralService("FT_SPERRLOG");
                SAPbobsCOM.GeneralData oGeneralData = (SAPbobsCOM.GeneralData)oGeneralService.GetDataInterface(SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralData);
                SAPbobsCOM.GeneralData oChild;
                SAPbobsCOM.GeneralDataCollection oChildren;
                oChildren = oGeneralData.Child("FT_SPERRLOG1");

                oGeneralData.SetProperty("U_ErrDT", JEhdr.ErrDT);
                oGeneralData.SetProperty("U_ErrMsg", JEhdr.ErrMsg);
                oGeneralData.SetProperty("U_RefDate", JEhdr.RefDate);
                oGeneralData.SetProperty("U_Memo", JEhdr.Memo);
                oGeneralData.SetProperty("U_DocType", JEhdr.DocType);
                oGeneralData.SetProperty("U_RInvNo", JEhdr.U_RInvNo);
                oGeneralData.SetProperty("U_ARInvNo", JEhdr.U_ARInvNo);
                oGeneralData.SetProperty("U_DelNo", JEhdr.U_DelNo);
                oGeneralData.SetProperty("U_APINVNo", JEhdr.U_APINVNo);

                foreach (JEDetails dtl in JEdtls)
                {
                    oChild = oChildren.Add();
                    oChild.SetProperty("U_AccountCode", dtl.AccountCode);
                    oChild.SetProperty("U_CostingCode", dtl.CostingCode);
                    oChild.SetProperty("U_Debit", dtl.Debit);
                    oChild.SetProperty("U_Credit", dtl.Credit);

                }

                oGeneralService.Add(oGeneralData);

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("ObjectFunctions ErrorLog " + ex.Message, 1, "Ok", "", "");
            }
        }
        public static bool Approval(string objcode)
        {
            //return true;
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "select * from [@FT_APPUSER] where U_OBJCODE = '" + objcode + "'";
            ors.DoQuery(sql);
            if (ors.RecordCount > 0)
            {
                return true;
            }
            return false;
        }
        public static bool Approval(string objcode, string username)
        {
            //return true;
            SAPbobsCOM.Recordset ors = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "select * from [@FT_APPUSER] where U_OBJCODE = '" + objcode + "' and U_USERID = '" + username + "'";
            ors.DoQuery(sql);
            if (ors.RecordCount > 0)
            {
                return true;
            }
            return false;
        }
        public static void CopyFromGridColumn(SAPbouiCOM.Form oForm, string griduid, string dsname, string column, int row)
        {
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(griduid).Specific;
            SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "select U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Header = 'N' and U_HColumn = '" + column + "'";
            rc.DoQuery(sql);
            if (rc.RecordCount > 0)
            {
                rc.MoveFirst();
                string btnsql = rc.Fields.Item(0).Value.ToString().Trim();
                btnsql = ObjectFunctions.ReplaceParams(oForm, btnsql, row);
                int size = oForm.DataSources.DBDataSources.Item("@" + dsname).Size;
                if (btnsql != "")
                {
                    string formuid = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
                    CFLInit.usercfl(formuid, "@" + dsname, column, row, griduid, "", btnsql, false);
                    //arrangematrix(oForm, oMatrix, "@" + dsname1);
                }
            }

        }
        public static void CopyFromItem(SAPbouiCOM.Form oForm, string ItemUID)
        {
            string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.ToString();
            SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "select U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Header = 'Y' and U_Btn = '" + ItemUID + "'";
            rc.DoQuery(sql);
            if (rc.RecordCount > 0)
            {
                rc.MoveFirst();
                string btnsql = rc.Fields.Item(0).Value.ToString().Trim();
                btnsql = ObjectFunctions.ReplaceParams(oForm, btnsql, 0);
                if (dsname == "FT_CHARGE" && ItemUID == "C_TPDOCNUM")
                {
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                    {
                        string formuid = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
                        CFLInit.usercfl(formuid, "@" + dsname, "", 0, "", "", btnsql, false);
                    }
                    else
                        SAP.SBOApplication.SetStatusBarMessage("Only new Change Module allowed.", SAPbouiCOM.BoMessageTime.bmt_Medium, true);
                }
                else
                {
                    string formuid = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
                    CFLInit.usercfl(formuid, "@" + dsname, "", 0, "", "", btnsql, false);
                }
            }

        }
        public static void CopyFromGrid1(SAPbouiCOM.Form oForm, SAPbouiCOM.ButtonCombo oButtonCombo)
        {

            if (oButtonCombo.Selected != null)
                if (oButtonCombo.Selected.Value != null)
                    if (oButtonCombo.Selected.Value.Trim() != "")
                    {
                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                        string dsname = oForm.DataSources.UserDataSources.Item("dsname").Value.ToString();
                        string dsname1 = oForm.DataSources.UserDataSources.Item("dsname1").Value.ToString();
                        SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                        string sql = "select U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Btn = '" + oButtonCombo.Selected.Value + "' and U_Header = 'N' and isnull(U_HColumn,'') = ''";
                        rc.DoQuery(sql);
                        if (rc.RecordCount > 0)
                        {
                            rc.MoveFirst();
                            string btnsql = rc.Fields.Item(0).Value.ToString().Trim();
                            btnsql = ObjectFunctions.ReplaceParams(oForm, btnsql, 0);
                            int size = oForm.DataSources.DBDataSources.Item("@" + dsname1).Size;
                            if (btnsql != "")
                            {
                                string formuid = oForm.DataSources.UserDataSources.Item("FormUID").Value.ToString();
                                CFLInit.usercfl(formuid, "@" + dsname1, "", 0, "grid1", "", btnsql, true);
                                //arrangematrix(oForm, oMatrix, "@" + dsname1);
                            }
                        }
                    }

        }
        public static string ReplaceParams(SAPbouiCOM.Form oForm, string syntax, int row)
        {
            string rtn = "";
            int start = -1;
            int end = -1;
            string param = "";
            //syntax = syntax.Replace('*', '%');
            for (int i = 0; i < syntax.Length; i++)
            {
                if (syntax[i] == '$' && start == -1)
                {
                    if (syntax[i + 1] == '[' && start == -1)
                    {
                        start = i;
                        for (int j = start; j < syntax.Length; j++)
                        {
                            if (syntax[j] == ']' && end == -1)
                            {
                                end = j;
                            }
                        }
                    }
                }
            }

            if (start >= 0 && end >= 0)
            {
                param = syntax.Substring(start, end - start + 1);
                string[] paramarr = param.Split('.');
                if (paramarr.Length == 2)
                {
                    string value = "";
                    string ds = paramarr[0].Substring(2, paramarr[0].Length - 2);
                    string column = paramarr[1].Substring(0, paramarr[1].Length - 1);
                    if (ds == "")
                    {
                        value = oForm.DataSources.UserDataSources.Item(column).ValueEx;
                        if (oForm.DataSources.UserDataSources.Item(column).DataType == SAPbouiCOM.BoDataType.dt_DATE)
                        {
                            string temp = value;
                            value = temp.Substring(0, 4) + "-" + temp.Substring(4, 2) + "-" + temp.Substring(6, 2);
                        }
                        else if (oForm.DataSources.UserDataSources.Item(column).DataType == SAPbouiCOM.BoDataType.dt_SHORT_TEXT)
                        {
                            value = "'" + value + "'";
                        }
                        else if (oForm.DataSources.UserDataSources.Item(column).DataType == SAPbouiCOM.BoDataType.dt_LONG_TEXT)
                        {
                            value = "'" + value + "'";
                        }
                    }
                    else
                    {
                        value = oForm.DataSources.DBDataSources.Item(ds).GetValue(column, row).Trim();
                        if (oForm.DataSources.DBDataSources.Item(ds).Fields.Item(column).Type == SAPbouiCOM.BoFieldsType.ft_Date)
                        {
                            string temp = value;
                            value = temp.Substring(0, 4) + "-" + temp.Substring(4, 2) + "-" + temp.Substring(6, 2);
                        }
                        else if (oForm.DataSources.DBDataSources.Item(ds).Fields.Item(column).Type == SAPbouiCOM.BoFieldsType.ft_AlphaNumeric)
                        {
                            value = "'" + value + "'";
                        }
                        else if (oForm.DataSources.DBDataSources.Item(ds).Fields.Item(column).Type == SAPbouiCOM.BoFieldsType.ft_Text)
                        {
                            value = "'" + value + "'";
                        }
                        else
                        {
                        }
                    }
                    value = value.Replace('*', '%');
                    rtn = syntax.Replace(param, value);
                    rtn = ReplaceParams(oForm, rtn, row);
                }
                else
                {
                    string value = param.Substring(2, param.Length - 3);
                    rtn = syntax.Replace(param, "'" + value.Trim() + "'");
                    rtn = ReplaceParams(oForm, rtn, row);
                }
            }
            else
                rtn = syntax;
            return rtn;
        }
        #region COPY Matrix
        public static void copyMatrixColumns(SAPbouiCOM.Form oSForm, SAPbouiCOM.Matrix oSMatrix, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                string columnname;
                int width = 0;
                string temp = "";
                string title = "";
                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                SAPbouiCOM.Column oSColumn = null;
                SAPbouiCOM.Column oColumn = null;

                SAPbouiCOM.UserDataSource oUds = null;
                int size = 0;
                SAPbouiCOM.BoDataType datatype;
                SAPbouiCOM.LinkedButton oLink;
                SAPbouiCOM.LinkedButton oSLink;
                
                for (int col = 0; col < oSMatrix.Columns.Count; col++)
                {
                    if (oSMatrix.Columns.Item(col).Visible)
                    {
                        oSColumn = oSMatrix.Columns.Item(col);
                        title = oSColumn.Title;//oSMatrix.Columns.Item(col).Title;
                        width = oSColumn.Width;//oSMatrix.Columns.Item(col).Width;
                        columnname = oSColumn.UniqueID.ToString();//oSMatrix.Columns.Item(col).UniqueID.ToString();
                        itemtype = oSColumn.Type; ;//oSMatrix.Columns.Item(col).Type;
                        oColumn = oMatrix.Columns.Add(columnname, itemtype);
                        oColumn.TitleObject.Caption = title;
                        oColumn.Width = width;
                        temp = oSColumn.DataBind.Alias;
                        if (temp == null)
                        {
                            size = 100;
                            datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                        }
                        else
                        {
                            size = oSForm.DataSources.UserDataSources.Item(oSColumn.DataBind.Alias).Length;
                            datatype = oSForm.DataSources.UserDataSources.Item(oSColumn.DataBind.Alias).DataType;
                        }
                        oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, size);
                        oColumn.DataBind.SetBound(true, "", columnname);
                        oColumn.Editable = oSColumn.Editable;
                        if (itemtype == SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                        {
                            oSLink = (SAPbouiCOM.LinkedButton)oSColumn.ExtendedObject;
                            oLink = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                            //SAP.SBOApplication.MessageBox(oSLink.LinkedObject.ToString(), 1, "ok", "", "");
                            oLink.LinkedObject = oSLink.LinkedObject;
                        }
                        else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                        {
                            oColumn.DisplayDesc = true;
                            for (int x = 0; x < oSColumn.ValidValues.Count; x++)
                            {
                                if (oSColumn.ValidValues.Item(x).Description.ToUpper() != "DEFINE NEW")
                                    oColumn.ValidValues.Add(oSColumn.ValidValues.Item(x).Value, oSColumn.ValidValues.Item(x).Description);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                
            }
        }
        public static void copyMatrixColumnsValues(SAPbouiCOM.Form oSForm, SAPbouiCOM.Matrix oSMatrix, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                string columnname;
                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;

                string temp = "";

                for (int row = 1; row <= oSMatrix.RowCount; row++)
                {
                    oMatrix.AddRow(1, -1);
                    for (int col = 0; col < oSMatrix.Columns.Count; col++)
                    {
                        if (oSMatrix.Columns.Item(col).Visible)
                        {
                            columnname = oSMatrix.Columns.Item(col).UniqueID.ToString();
                            itemtype = oSMatrix.Columns.Item(col).Type;
                            if (itemtype == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                            {
                                temp = ((SAPbouiCOM.ComboBox)(oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific)).Selected.Value.ToString();
                                ((SAPbouiCOM.ComboBox)(oMatrix.Columns.Item(columnname).Cells.Item(1).Specific)).Select(temp, SAPbouiCOM.BoSearchKey.psk_ByValue);
                            }
                            else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                            {
                                if (((SAPbouiCOM.CheckBox)oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific).Checked)
                                {
                                    oMatrix.Columns.Item(columnname).Cells.Item(row).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                                }
                            }
                            else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                            {
                                temp = ((SAPbouiCOM.EditText)oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific).String;
                                ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnname).Cells.Item(row).Specific)).String = temp;
                            }
                            else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                            {
                                temp = ((SAPbouiCOM.EditText)oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific).String;
                                ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnname).Cells.Item(row).Specific)).String = temp;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        #endregion
        public static SAPbouiCOM.BoDataType changeUIFieldsTypeToUIDataType(SAPbouiCOM.BoFieldsType fieldType)
        {
            SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_RATE;
            switch (fieldType)
            {
                case SAPbouiCOM.BoFieldsType.ft_NotDefined:
                case SAPbouiCOM.BoFieldsType.ft_AlphaNumeric:
                    datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                    break;
                case SAPbouiCOM.BoFieldsType.ft_Date:
                    datatype = SAPbouiCOM.BoDataType.dt_DATE;
                    break;
                case SAPbouiCOM.BoFieldsType.ft_Integer:
                    datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                    break;
               case SAPbouiCOM.BoFieldsType.ft_Text:
                    datatype = SAPbouiCOM.BoDataType.dt_LONG_TEXT;
                    break;
            }
            return datatype;
        }
        public static SAPbouiCOM.BoDataType changeDIFieldTypesToDIDataType(SAPbobsCOM.BoFieldTypes fieldType)
        {
            SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_RATE;
            switch (fieldType)
            {
                case SAPbobsCOM.BoFieldTypes.db_Alpha:
                    datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Date:
                    datatype = SAPbouiCOM.BoDataType.dt_DATE;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Memo:
                    datatype = SAPbouiCOM.BoDataType.dt_LONG_TEXT;
                    break;
                case SAPbobsCOM.BoFieldTypes.db_Numeric:
                    datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                    break;
            }
            return datatype;
        }

        public static void customFormMatrixSetting(SAPbouiCOM.Form oForm, string matrixname, string userid, string table)
        {
            try
            {
                string formname = oForm.TypeEx;
                string code = "";

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT TOP 1 CODE FROM [@FT_CFS] WHERE U_FNAME = '" + formname + "' AND U_USRID = '" + userid + "' AND U_MATRIX = '" + matrixname + "' AND U_DSNAME = '" + table + "' ORDER BY CODE DESC");
                if (rs.RecordCount > 0)
                {
                    rs.MoveFirst();
                    code = rs.Fields.Item(0).Value.ToString();

                    SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixname).Specific;
                    SAPbouiCOM.Column oColumn = null;
                    int nonview = 0;
                    int nonedit = 0;
                    string colname = "";

                    rs.DoQuery("SELECT U_CNAME, U_NONVIEW, U_NONEDIT FROM [@FT_CFSDL] WHERE CODE = '" + code + "'");
                    if (rs.RecordCount > 0)
                    {
                        rs.MoveFirst();
                        while (!rs.EoF)
                        {
                            colname = rs.Fields.Item(0).Value.ToString();
                            nonview = int.Parse(rs.Fields.Item(1).Value.ToString());
                            nonedit = int.Parse(rs.Fields.Item(2).Value.ToString());
                            if (nonview == 1)
                            {
                                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(colname);
                                oColumn.Visible = false;
                            }
                            else if (nonedit == 1)
                            {
                                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(colname);
                                oColumn.Editable = false;
                            }
                            rs.MoveNext();
                        }
                    }
                }
                else
                {
                    rs.DoQuery("SELECT TOP 1 CODE FROM [@FT_CFS] WHERE U_FNAME = '" + formname + "' AND isnull(U_USRID,'') = '' AND U_MATRIX = '" + matrixname + "' AND U_DSNAME = '" + table + "' ORDER BY CODE DESC");
                    if (rs.RecordCount > 0)
                    {
                        rs.MoveFirst();
                        code = rs.Fields.Item(0).Value.ToString();

                        SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item(matrixname).Specific;
                        SAPbouiCOM.Column oColumn = null;
                        int nonview = 0;
                        int nonedit = 0;
                        string colname = "";

                        rs.DoQuery("SELECT U_CNAME, U_NONVIEW, U_NONEDIT FROM [@FT_CFSDL] WHERE CODE = '" + code + "'");
                        if (rs.RecordCount > 0)
                        {
                            try
                            {
                                rs.MoveFirst();
                                while (!rs.EoF)
                                {
                                    colname = rs.Fields.Item(0).Value.ToString();
                                    nonview = int.Parse(rs.Fields.Item(1).Value.ToString());
                                    nonedit = int.Parse(rs.Fields.Item(2).Value.ToString());
                                    if (nonview == 1)
                                    {
                                        oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(colname);
                                        oColumn.Visible = false;
                                    }
                                    else if (nonedit == 1)
                                    {
                                        oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(colname);
                                        oColumn.Editable = false;
                                    }
                                    rs.MoveNext();
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception(colname + ":"  + ex.Message);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }

        public static void documentFormMatrixSetting(string userid, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix)
        {
            try
            {
                string code = "";
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery("SELECT TOP 1 CODE FROM [@FT_DFS] WHERE U_USRID = '" + userid + "' ORDER BY CODE DESC");
                if (rs.RecordCount > 0)
                {
                    rs.MoveFirst();
                    code = rs.Fields.Item(0).Value.ToString();

                    SAPbouiCOM.Column oColumn = null;
                    int nonview = 0;
                    int nonedit = 0;
                    string colname = "";

                    rs.DoQuery("SELECT U_CNAME, U_NONVIEW, U_NONEDIT FROM [@FT_DFSDL] WHERE CODE = '" + code + "'");
                    if (rs.RecordCount > 0)
                    {
                        rs.MoveFirst();
                        while (!rs.EoF)
                        {
                            colname = rs.Fields.Item(0).Value.ToString();
                            nonview = int.Parse(rs.Fields.Item(1).Value.ToString());
                            nonedit = int.Parse(rs.Fields.Item(2).Value.ToString());
                            if (nonview == 1)
                            {
                                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(colname);
                                oColumn.Visible = false;
                            }
                            else if (nonedit == 1)
                            {
                                oColumn = (SAPbouiCOM.Column)oMatrix.Columns.Item(colname);
                                oColumn.Editable = false;
                            }
                            rs.MoveNext();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }

    }
}
