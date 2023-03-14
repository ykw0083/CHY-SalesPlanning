using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON
{
    class ApplicationCommon
    {
        // -------------------------------------------------------------
        // Create UDT, UDF, UDO and Add Menu Item
        // -------------------------------------------------------------
        private Boolean checksuperuser(string userid)
        {
            SAPbobsCOM.Recordset rc = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            string sql = "select superuser from ousr where user_code = '" + userid + "' and superuser = 'Y'";
            rc.DoQuery(sql);
            bool rtn = false;
            if (rc.RecordCount > 0)
            {
                rtn = true;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(rc);
            rc = null;
            return rtn;
       }

        public Boolean createTable(string tableName, string tableDescription, SAPbobsCOM.BoUTBTableType tableType)
        {
            if (!checksuperuser(SAP.SBOCompany.UserName)) return true;
            GC.Collect();

            int IRetCode = 0;
            SAPbobsCOM.UserTablesMD oUserTableMD = null;
            oUserTableMD = ((SAPbobsCOM.UserTablesMD)(FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserTables)));
            if (!oUserTableMD.GetByKey(tableName))
            {
                FT_ADDON.SAP.setStatus("Creating Table : " + tableName + " - " + tableDescription);
                oUserTableMD.TableName = tableName;
                oUserTableMD.TableDescription = tableDescription;
                oUserTableMD.TableType = tableType;
                IRetCode = oUserTableMD.Add();
                if (IRetCode != 0)
                {
                    FT_ADDON.SAP.SBOApplication.MessageBox(FT_ADDON.SAP.SBOCompany.GetLastErrorDescription(), 1, "&Ok", "", "");
                    oUserTableMD = null;
                    return false;
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserTableMD);
            oUserTableMD = null;
            GC.Collect();
            return true;
        }

        public Boolean createField(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int size)
        {
            return createField(tableName, fieldName, fieldDescription, fieldType, size, "", false, SAPbobsCOM.BoFldSubTypes.st_None, "", "");
        }

        public Boolean createField(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int size, string defaultValue)
        {
            return createField(tableName, fieldName, fieldDescription, fieldType, size, defaultValue, false, SAPbobsCOM.BoFldSubTypes.st_None, "", "");
        }

        public Boolean createField(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int size, string defaultValue, Boolean mandatory)
        {
            return createField(tableName, fieldName, fieldDescription, fieldType, size, defaultValue, mandatory, SAPbobsCOM.BoFldSubTypes.st_None, "", "");
        }

        public Boolean createField(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int size, string defaultValue, Boolean mandatory, SAPbobsCOM.BoFldSubTypes subType)
        {
            return createField(tableName, fieldName, fieldDescription, fieldType, size, defaultValue, mandatory, subType, "", "");
        }

        public Boolean createField(string tableName, string fieldName, string fieldDescription, SAPbobsCOM.BoFieldTypes fieldType, int size, string defaultValue, Boolean mandatory, SAPbobsCOM.BoFldSubTypes subType, string validValues, string linkedTable)
        {
            if (!checksuperuser(SAP.SBOCompany.UserName)) return true;
            GC.Collect();

            FT_ADDON.SAP.setStatus("Creating Field : " + fieldName + " - " + fieldDescription);
            int IRetCode = 0;
            int IvalidValues = 0;
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)(FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields));
            oUserFieldsMD.TableName = tableName;
            oUserFieldsMD.Name = fieldName;
            oUserFieldsMD.Description = fieldDescription;
            oUserFieldsMD.Type = fieldType;
            if (subType != SAPbobsCOM.BoFldSubTypes.st_None) { oUserFieldsMD.SubType = subType; }
            if (fieldType != SAPbobsCOM.BoFieldTypes.db_Numeric && size > 0) oUserFieldsMD.Size = size;
            if (fieldType == SAPbobsCOM.BoFieldTypes.db_Numeric && size > 0) { oUserFieldsMD.EditSize = size; }
            if (defaultValue != "") 
            {
                oUserFieldsMD.DefaultValue = defaultValue;
            }
            if (linkedTable != "") oUserFieldsMD.LinkedTable = linkedTable;
            if (validValues != "")
            {
                foreach (string value in validValues.Split('|'))
                {
                    IvalidValues++;
                    string[] parm = value.Split(':');
                    if (IvalidValues != 1) oUserFieldsMD.ValidValues.Add();
                    oUserFieldsMD.ValidValues.SetCurrentLine(IvalidValues - 1);
                    oUserFieldsMD.ValidValues.Value = parm[0];
                    oUserFieldsMD.ValidValues.Description = parm[1];
                }
            }
            if (mandatory) oUserFieldsMD.Mandatory = SAPbobsCOM.BoYesNoEnum.tYES;

            IRetCode = oUserFieldsMD.Add();
            if (IRetCode != 0)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox("Error : " + IRetCode.ToString() + "\n" + FT_ADDON.SAP.SBOCompany.GetLastErrorDescription(),1,"&Ok","","");
                oUserFieldsMD = null;
                GC.Collect();
                return false;
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
            return true;
        }

        public Boolean tableGotField(string tablename)
        {
            if (!checksuperuser(SAP.SBOCompany.UserName)) return true;
            GC.Collect();

            // Check if table has UDF, include @ if table is UDT
            SAPbobsCOM.UserFieldsMD oUserFieldsMD = null;
            oUserFieldsMD = (SAPbobsCOM.UserFieldsMD)(FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields));
            Boolean rtn = oUserFieldsMD.GetByKey(tablename, 0);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserFieldsMD);
            oUserFieldsMD = null;
            GC.Collect();
            return rtn;
        }

        public Boolean udfExist(string tableName, string fieldName)
        {
            if (!checksuperuser(SAP.SBOCompany.UserName)) return true;
            GC.Collect();

            SAPbobsCOM.Recordset oRec = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
            oRec.DoQuery("SELECT AliasID from CUFD where TableID='" + tableName + "' AND AliasID = '" + fieldName + "'");
            Boolean rtn = oRec.RecordCount > 0;
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oRec);
            oRec = null;
            GC.Collect();
            return rtn;
        }

        public Boolean createUDO(string udoName, string udoDescription, SAPbobsCOM.BoUDOObjType objType, string tableName, string childName, string findColumns, SAPbobsCOM.BoYesNoEnum manageSeries, SAPbobsCOM.BoYesNoEnum canCancel, SAPbobsCOM.BoYesNoEnum canClose, SAPbobsCOM.BoYesNoEnum canDelete, SAPbobsCOM.BoYesNoEnum log, string logName)
        {
            if (!checksuperuser(SAP.SBOCompany.UserName)) return true;
            GC.Collect();
           
            SAPbobsCOM.UserObjectsMD oUserObjectMD;
            int IRetCode = 0;
            int Iindex = 0;
            Boolean rtn = true;
            oUserObjectMD = (SAPbobsCOM.UserObjectsMD)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserObjectsMD);
            if (!oUserObjectMD.GetByKey(udoName))
            {
                FT_ADDON.SAP.setStatus("Creating UDO : " + udoDescription);
                oUserObjectMD.Code = udoName;
                oUserObjectMD.Name = udoDescription;
                oUserObjectMD.ObjectType = objType;
                oUserObjectMD.LogTableName = "";
                oUserObjectMD.TableName = tableName;
                oUserObjectMD.CanCancel = canCancel;
                oUserObjectMD.CanClose = canClose;
                oUserObjectMD.CanDelete = canDelete;
                oUserObjectMD.CanFind = SAPbobsCOM.BoYesNoEnum.tYES;
                oUserObjectMD.CanLog = log;
                oUserObjectMD.LogTableName = logName;
                oUserObjectMD.CanYearTransfer = SAPbobsCOM.BoYesNoEnum.tNO;
                oUserObjectMD.ExtensionName = "";
                oUserObjectMD.CanCreateDefaultForm = SAPbobsCOM.BoYesNoEnum.tNO;

                if (childName != "")
                {
                    foreach (string childTable in childName.Split('|'))
                    {
                        Iindex++;
                        if (Iindex > 1) oUserObjectMD.ChildTables.Add();
                        oUserObjectMD.ChildTables.SetCurrentLine(Iindex - 1);
                        oUserObjectMD.ChildTables.TableName = childTable;
                    }
                }
                if (manageSeries == SAPbobsCOM.BoYesNoEnum.tYES && objType == SAPbobsCOM.BoUDOObjType.boud_Document)
                    oUserObjectMD.ManageSeries = manageSeries;
                else oUserObjectMD.ManageSeries = SAPbobsCOM.BoYesNoEnum.tNO;
                Iindex = 0;
                if (findColumns != "")
                {
                    foreach (string colName in findColumns.Split('|'))
                    {
                        Iindex++;
                        if (Iindex > 1) oUserObjectMD.FindColumns.Add();
                        oUserObjectMD.FindColumns.SetCurrentLine(Iindex - 1);
                        oUserObjectMD.FindColumns.ColumnAlias = colName;
                    }
                }
                IRetCode = oUserObjectMD.Add();
                if (IRetCode != 0)
                {
                    FT_ADDON.SAP.SBOApplication.MessageBox("Error : " + IRetCode.ToString() + "\n" + FT_ADDON.SAP.SBOCompany.GetLastErrorDescription(), 1, "Ok", "", "");
                    rtn = false;
                }
            }
            System.Runtime.InteropServices.Marshal.ReleaseComObject(oUserObjectMD);
            oUserObjectMD = null;
            GC.Collect();
            return rtn;
        }

        public Boolean createMenuItem(string menuID, string menuName, string mainMenuID, string refMenuID ,Boolean beforeRef, SAPbouiCOM.BoMenuType menuType)
        {
                Boolean rtn = true;
                SAPbouiCOM.Menus oMenus = null;
                SAPbouiCOM.MenuItem oMenuItem = null, oRefItem = null;
                oMenus = FT_ADDON.SAP.SBOApplication.Menus;
                oMenuItem = FT_ADDON.SAP.SBOApplication.Menus.Item(mainMenuID); // moudles'
                oRefItem = FT_ADDON.SAP.SBOApplication.Menus.Item(refMenuID);
                int IRef = 0;
                oMenus = oMenuItem.SubMenus;
                if (refMenuID == "")
                {
                    IRef = oMenus.Count;
                }
                else
                {
                    for (int i = 0; i < oMenus.Count - 1; i++)
                    {
                        if (oMenus.Item(i).UID == oRefItem.UID)
                        {
                            IRef = i;
                            break;
                        }
                    }
                    if (!beforeRef) IRef++;
                }

                try
                {
                    oMenus.Add(menuID, menuName, menuType, IRef);
                }
            catch { rtn = false; }
            oMenus = null;
            oMenuItem = null;
            oRefItem = null;
            GC.Collect();
            return rtn;
        }
    }
}
