using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace FT_ADDON.CHY
{
    class InitForm
    {
        public static void genFT_SPERRLOG()
        {
            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize Listing window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
            creationPackage.FormType = "FT_GENSPERR";
            SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
            oForm.Title = "Generate SPERRLOG";
            //oForm.Left = oSForm.Left;
            oForm.Width = 600;
            //oForm.Top = oSForm.Top;
            oForm.Height = 500;

            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Add("cb_ARIV", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = 5;
            oItem.Height = 20;
            ((SAPbouiCOM.Button)oItem.Specific).Caption = "Invoice";

            oItem = oForm.Items.Add("cb_RIV", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5 + 70;
            oItem.Width = 65;
            oItem.Top = 5;
            oItem.Height = 20;
            ((SAPbouiCOM.Button)oItem.Specific).Caption = "R Invoice";

            oItem = oForm.Items.Add("cb_APIV", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5 + 70 + 70;
            oItem.Width = 65;
            oItem.Top = 5;
            oItem.Height = 20;
            ((SAPbouiCOM.Button)oItem.Specific).Caption = "AP Invoice";

            oItem = oForm.Items.Add("st_msg", SAPbouiCOM.BoFormItemTypes.it_STATIC);
            oItem.Left = 5;
            oItem.Width = 590;
            oItem.Top = 50;
            oItem.Height = 20;

            oForm.Visible = true;
        }
        public static void FT_SPLANscreenpainter(string dsname, string dsname1, string dsnameb)
        {
            try
            {

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                string path = System.Windows.Forms.Application.StartupPath;
                xmlDoc.Load(path + "\\" + dsname + ".srf");

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.XmlData = xmlDoc.InnerXml;     // Load form from xml 

                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);

                oForm.Freeze(true);
                oForm.AutoManaged = true;

                SAPbobsCOM.Recordset oSeq = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "select * from OUDO where Code = '" + dsname + "'";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();


                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("dsname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = dsname;
                uds = oForm.DataSources.UserDataSources.Add("dsname1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = dsname1;
                uds = oForm.DataSources.UserDataSources.Add("dsnameb", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = dsnameb;

                uds = oForm.DataSources.UserDataSources.Add("FormUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = oForm.UniqueID;
                uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = "";
                uds = oForm.DataSources.UserDataSources.Add("docstatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = "";

                //uds = oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Button oButton = null;
                SAPbouiCOM.EditText oEdit = null;
                SAPbouiCOM.ComboBox oCombo = null;
                SAPbouiCOM.LinkedButton oLinkedButton = null;
                SAPbouiCOM.Folder oFolder = null;

                oItem = oForm.Items.Item("folder1");
                oItem.AffectsFormMode = false;
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oFolder.DataBind.SetBound(true, "", "FolderDS");
                oFolder.Select();

                oItem = oForm.Items.Item("folder2");
                oItem.AffectsFormMode = false;
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oFolder.DataBind.SetBound(true, "", "FolderDS");
                oFolder.GroupWith("folder1");



                //if (dsnameb != "")
                //{
                //    oItem = oForm.Items.Add("cb_batch", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                //    oItem.Left = 150;
                //    oItem.Width = 65;
                //    oItem.Top = oForm.Height - 60;
                //    oItem.Height = 20;
                //    oButton = (SAPbouiCOM.Button)oItem.Specific;
                //    oButton.Caption = "Batch/Bin";
                //}
                //if (dsname == "FT_TPPLAN")
                //{
                //    oItem = oForm.Items.Add("cb_trans", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                //    oItem.Left = oForm.Width - 240;
                //    oItem.Width = 100;
                //    oItem.Top = oForm.Height - 60;
                //    oItem.Height = 20;
                //    oButton = (SAPbouiCOM.Button)oItem.Specific;
                //    oButton.Caption = "Transfer Request";
                //}

                sql = "select U_UDO, U_Btn, U_BtnName, U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Header = 'N' and isnull(U_HColumn,'') = ''";
                oSeq.DoQuery(sql);
                if (oSeq.RecordCount > 0)
                {
                    oSeq.MoveFirst();
                    SAPbouiCOM.ButtonCombo oButtonCombo = null;

                    oItem = oForm.Items.Item("cb_Copy");
                    oItem.DisplayDesc = true;
                    oItem.AffectsFormMode = false;
                    oButtonCombo = (SAPbouiCOM.ButtonCombo)oItem.Specific;

                    while (!oSeq.EoF)
                    {
                        oButtonCombo.ValidValues.Add(oSeq.Fields.Item(1).Value.ToString(), oSeq.Fields.Item(2).Value.ToString());
                        oSeq.MoveNext();
                    }
                    //oButtonCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }

                SAPbouiCOM.Conditions oBPCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oBPCon = oBPCons.Add();

                oBPCon.Alias = "CardType";
                oBPCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oBPCon.CondVal = "C";
                oBPCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oBPCon = oBPCons.Add();
                oBPCon.Alias = "FrozenFor";
                oBPCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oBPCon.CondVal = "N";

                SAPbouiCOM.Conditions oItemCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oItemCon = oItemCons.Add();
                oItemCon.Alias = "SellItem";
                oItemCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oItemCon.CondVal = "Y";

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;

                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLitem";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oItemCons);


                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFLCust";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oBPCons);

                SAPbouiCOM.Conditions oWHSCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oWHSCon = oWHSCons.Add();

                oWHSCon.Alias = "U_DropShip";
                oWHSCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oWHSCon.CondVal = "Y";

                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLwhsH";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oWHSCons);

                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLwhs";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oWHSCons);

                //oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                //oCFLCreationParams.MultiSelection = false;
                //oCFLCreationParams.ObjectType = "TRANSPORTER";
                //oCFLCreationParams.UniqueID = "CFLTrans";
                //oCFL = oCFLs.Add(oCFLCreationParams);

                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                SAPbouiCOM.BoFormItemTypes linktype = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON;
                SAPbouiCOM.BoFormItemTypes itemtypecmb = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX;
                SAPbouiCOM.DataTable dt = null;// oForm.DataSources.DataTables.Add("99");

                SAPbouiCOM.Column oColumn = null;
                //SAPbouiCOM.UserDataSource oUds = null;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                SAPbobsCOM.UserTable oUT = null;
                SAPbobsCOM.UserFields oUF = null;
                string linkedtable = "";
                SAPbouiCOM.Matrix oMatrix = null;

                int top = -15;
                int cnt = -1;
                string columnname = "";
                Boolean end = false;

                //oForm.DataSources.DBDataSources.Add("@" + dsname);

                //top = top + 20;

                //oItem = oForm.Items.Add("8", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                //((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Doc No";
                //oItem.Left = 5;
                //oItem.Top = top;
                //oItem.Width = 150;
                //oItem.Height = 15;

                columnname = "series";
                oItem = oForm.Items.Item(columnname);
                oItem.DisplayDesc = true;
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                #region get series
                oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("series").Specific;
                SAPbobsCOM.CompanyService companyServices = SAP.SBOCompany.GetCompanyService();
                SAPbobsCOM.SeriesService seriesService = (SAPbobsCOM.SeriesService)companyServices.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService);
                SAPbobsCOM.DocumentTypeParams oDocumentTypeParams = (SAPbobsCOM.DocumentTypeParams)seriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams);
                oDocumentTypeParams.Document = dsname;

                SAPbobsCOM.Series oSeries = seriesService.GetDefaultSeries(oDocumentTypeParams);
                int defSeries = 0;
                if (oSeries.Series != 0)
                {
                    defSeries = oSeries.Series;
                }
                //oForm.DataSources.UserDataSources.Add("DEFSERIES", SAPbouiCOM.BoDataType.dt_LONG_NUMBER).ValueEx = defSeries.ToString();
                //SAPbobsCOM.PathAdmin oPathAdmin = companyServices.GetPathAdmin();
                //oForm.DataSources.UserDataSources.Add("ATTACH", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 500).ValueEx = oPathAdmin.AttachmentsFolderPath;
                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCondition = oConditions.Add();

                oCondition.Alias = "ObjectCode";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = dsname;
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCondition = oConditions.Add();
                oCondition.Alias = "Locked";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "N";

                SAPbouiCOM.DBDataSource nnm1 = oForm.DataSources.DBDataSources.Add("NNM1");
                nnm1.Query(oConditions);

                companyServices = null;
                seriesService = null;
                oDocumentTypeParams = null;
                oSeries = null;
                uds = oForm.DataSources.UserDataSources.Add("defseries", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER);
                if (nnm1.Size > 0)
                {
                    for (int i = 0; i < nnm1.Size; i++)
                    {
                        nnm1.Offset = i;
                        oCombo.ValidValues.Add(nnm1.GetValue("Series", i), nnm1.GetValue("SeriesName", i));
                        if (defSeries == 0)
                        {
                            if (i == 0)
                            {
                                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                uds.Value = "0";
                                //oForm.DataSources.DBDataSources.Item("@FT_CCF").SetValue("DocNum", 0, nnm1.GetValue("NextNumber", 0));
                            }
                        }
                        else
                        {
                            if (nnm1.GetValue("Series", i) == defSeries.ToString())
                            {
                                oCombo.Select(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                uds.Value = i.ToString();
                            }
                        }
                    }
                }

                #endregion

                #region header
                columnname = "docnum";
                oItem = oForm.Items.Item(columnname);

                oForm.AutoManaged = true;
                oForm.DataBrowser.BrowseBy = "docnum";
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;

                columnname = "docstatus";
                oItem = oForm.Items.Item(columnname);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "docstatus");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                //oSeq = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                sql = "select AliasID, rtABLE from CUFD where TableID = '@" + dsname + "' and isnull(U_seq,0) >= 0 order by U_seq, FieldID";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();

                string aliasid = "";
                bool btnfound = false;
                SAPbobsCOM.Recordset oBtn = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                while (!oSeq.EoF)
                {
                    aliasid = "U_" + oSeq.Fields.Item(0).Value.ToString();

                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        columnname = oUF.Fields.Item(x).Name;
                        switch (columnname)
                        {
                            case "U_APP":
                            case "U_APPRE":
                            case "U_APPBY":
                            case "U_APPDATE":
                            case "U_APPTIME":
                                continue;
                            default:
                                break;
                        }

                        if (columnname != aliasid) continue;

                        try
                        {
                            oItem = oForm.Items.Item(columnname);
                        }
                        catch
                        {
                            continue;
                        }

                        sql = "select U_Btn, U_BtnName, U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Header = 'Y' and U_HColumn = '" + columnname + "'";
                        oBtn.DoQuery(sql);
                        btnfound = false;
                        if (oBtn.RecordCount > 0)
                        {
                            oBtn.MoveFirst();
                            oItem = oForm.Items.Add(oBtn.Fields.Item(0).Value.ToString(), SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oButton = (SAPbouiCOM.Button)oItem.Specific;
                            oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                            oButton.Image = Application.StartupPath + @"\CFL.BMP";
                            btnfound = true;
                        }


                        linkedtable = oSeq.Fields.Item(1).Value.ToString().Trim();
                        //linkedtable = oUF.Fields.Item(x).LinkedTable;
                        if (linkedtable != "")
                        {
                            oItem = oForm.Items.Item(columnname);
                            oItem.DisplayDesc = true;
                            oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;

                            dt = oForm.DataSources.DataTables.Add(linkedtable);
                            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                            for (int y = 0; y < dt.Rows.Count; y++)
                            {
                                ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                            }
                        }
                        else if (oUF.Fields.Item(x).ValidValues.Count > 0)
                        {
                            oItem = oForm.Items.Item(columnname);
                            oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                            oItem.DisplayDesc = true;

                            //for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                            //{
                            //    ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                            //}
                        }
                        else
                        {
                            oItem = oForm.Items.Item(columnname);
                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;

                            if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                            {
                                //oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            }
                            //if (columnname == "U_ADDRESS")
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            //if (columnname == "U_REMARKS")
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            //if (columnname == "U_FREMARKS")
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            if (columnname == "U_DOCDATE")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_CARDNAME")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_DOCTOTAL")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            //if (columnname == "U_APRICE")
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_TPDOCNUM")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_CARDCODE")
                            {
                                //oItemtemp = oForm.Items.Add("c" + columnname.Substring(1, columnname.Length - 1), SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                                //if (end)
                                //    oItemtemp.Left = oForm.Width - 66;
                                //else
                                //    oItemtemp.Left = 155 + 149;
                                //oItemtemp.Top = top - 2;
                                //oItemtemp.Width = 20;
                                //oItemtemp.Height = 20;
                                //oButton = ((SAPbouiCOM.Button)(oItemtemp.Specific));
                                //oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                                //oButton.Image = Application.StartupPath + @"\CFL.BMP";
                                //oCFLCreationParams.UniqueID = "c" + columnname.Substring(1, columnname.Length - 1);
                                //oCFL = oCFLs.Add(oCFLCreationParams);
                                //oButton.ChooseFromListUID = "c" + columnname.Substring(1, columnname.Length - 1);
                                oEdit.ChooseFromListUID = "CFLCust";
                                oEdit.ChooseFromListAlias = "CardCode";
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                                SAPbouiCOM.Item oLinkItem = oForm.Items.Add("LBP", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                oLinkItem.LinkTo = "U_CARDCODE";
                                ((SAPbouiCOM.LinkedButton)oLinkItem.Specific).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                            }
                            if (columnname == "U_WHSCODE")
                            {
                                oEdit.ChooseFromListUID = "CFLwhsH";
                                oEdit.ChooseFromListAlias = "WhsCode";
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                                SAPbouiCOM.Item oLinkItem = oForm.Items.Add("LWH", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                //oLinkItem.Top = oForm.Items.Item("U_WHSCODE").Top;
                                //oLinkItem.Left = oForm.Items.Item("U_WHSCODE").Left - 10;
                                oLinkItem.LinkTo = "U_WHSCODE";
                                ((SAPbouiCOM.LinkedButton)oLinkItem.Specific).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                            }
                            //if (columnname == "U_TPCODE")
                            //{
                            //    oEdit.ChooseFromListUID = "CFLTrans";
                            //    oEdit.ChooseFromListAlias = "Name";
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            //}
                        }

                        if (columnname == "U_CARDCODE")
                        {
                            oForm.Items.Item("LBP").Top = oForm.Items.Item("U_CARDCODE").Top;
                            oForm.Items.Item("LBP").Left = oForm.Items.Item("U_CARDCODE").Left - 20;
                        }
                        if (columnname == "U_WHSCODE")
                        {
                            oForm.Items.Item("LWH").Top = oForm.Items.Item("U_WHSCODE").Top;
                            oForm.Items.Item("LWH").Left = oForm.Items.Item("U_WHSCODE").Left - 20;
                        }

                        if (btnfound)
                        {
                            oButton.Item.Top = oItem.Top - 2;
                            oButton.Item.Left = oItem.Left + oItem.Width;
                            oButton.Item.Width = 15;
                        }
                    }
                    oSeq.MoveNext();
                }
                #endregion

                #region grid 1
                oItem = oForm.Items.Item("grid1");
                oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

                string dtldsname = dsname1;
                oForm.DataSources.DBDataSources.Add("@" + dtldsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dtldsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);
                oColumn.Editable = false;

                oMatrix.Columns.Remove("V_-1");
                oMatrix.Columns.Remove("V_0");

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dtldsname);
                oUF = oUT.UserFields;

                sql = "select AliasID from CUFD where TableID = '@" + dtldsname + "' order by U_seq, FieldID";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();

                aliasid = "";

                while (!oSeq.EoF)
                {
                    aliasid = "U_" + oSeq.Fields.Item(0).Value.ToString();
                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        columnname = oUF.Fields.Item(x).Name;
                        if (columnname != aliasid) continue;

                        linkedtable = oUF.Fields.Item(x).LinkedTable;
                        if (linkedtable == "")
                        {
                            //if (columnname == "U_size")
                            //    linkedtable = "0009";
                            //else if (columnname == "U_jcclr")
                            //    linkedtable = "0007";
                            //else if (columnname == "U_brand")
                            //    linkedtable = "0004";
                            //else if (columnname == "U_perfcl")
                            //    linkedtable = "0001";
                        }
                        if (linkedtable != "")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                            oColumn.DisplayDesc = true;
                            dt = oForm.DataSources.DataTables.Add(linkedtable);
                            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                            for (int y = 0; y < dt.Rows.Count; y++)
                            {
                                oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                            }

                        }
                        else if (oUF.Fields.Item(x).ValidValues.Count > 0)
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                            oColumn.DisplayDesc = true;
                            oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);
                            for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                            {
                                oColumn.ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                            }

                            //oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            //oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                            //oCombo.DataBind.SetBound(true, "@" + dsname, columnname);
                            //oItem.DisplayDesc = true;

                            //for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                            //{
                            //    ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                            //}
                        }
                        else if (columnname == "U_SOITEMCO" || columnname == "U_ITEMCODE")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, linktype);
                            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                        }
                        else if (columnname == "U_WHSCODE")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, linktype);
                            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                        }
                        else if (columnname == "U_CARDCODE")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, linktype);
                            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                        }
                        else
                        {
                            //if (columnname == "U_itemcode")
                            //{
                            //    oColumn = oMatrix.Columns.Add(columnname, linktype);
                            //}
                            //else
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                        }
                        oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                        if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                            oColumn.Width = 200;
                        else
                            oColumn.Width = 100;

                        oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);

                        if (columnname == "U_LTOTAL")
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        else if (columnname == "U_WHSCODE")
                        {
                            oColumn.ChooseFromListUID = "CFLwhs";
                            oColumn.ChooseFromListAlias = "WhsCode";
                            //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                        }
                        else if (columnname == "U_ITEMCODE")
                        {
                            oColumn.ChooseFromListUID = "CFLitem";
                            oColumn.ChooseFromListAlias = "ItemCode";

                        }
                        else
                        {
                            sql = "select U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname1 + "' and U_Header = 'N' and U_HColumn = '" + columnname + "'";
                            oBtn.DoQuery(sql);
                            if (oBtn.RecordCount > 0)
                            {
                                oCFLCreationParams.UniqueID = columnname;
                                oCFL = oCFLs.Add(oCFLCreationParams);
                                oColumn.ChooseFromListUID = columnname;
                            }

                        }
                        //else if (columnname == "U_ITEMCODE")
                        //{
                        //    oColumn.ChooseFromListUID = "CFLitem";
                        //    oColumn.ChooseFromListAlias = "ItemCode";
                        //    //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                        //}


                    }
                    oSeq.MoveNext();
                }
                #endregion


                oForm.EnableFormatSearch();
                oForm.DataBrowser.BrowseBy = "docnum";

                oForm.PaneLevel = 1;
                oForm.Freeze(false);
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                UserForm_SalesPlanning.AddNew(oForm);

                //oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }
        public static void approval(string FormID, string appcode, string objecttype, string table, string table1, string title, bool sysobj, string sptype)
        {
            try
            {

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                string path = System.Windows.Forms.Application.StartupPath;
                xmlDoc.Load(path + "\\APPROVAL2.srf");

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.XmlData = xmlDoc.InnerXml;     // Load form from xml 
                creationPackage.FormType = FormID;
                creationPackage.ObjectType = objecttype;

                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = title;

                oForm.Freeze(true);
                SAPbouiCOM.UserDataSource uds = null;
                uds = oForm.DataSources.UserDataSources.Add("sptype", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = sptype;
                uds = oForm.DataSources.UserDataSources.Add("table", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = table;
                uds = oForm.DataSources.UserDataSources.Add("table1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = table1;
                uds = oForm.DataSources.UserDataSources.Add("appcode", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = appcode;
                uds = oForm.DataSources.UserDataSources.Add("objecttype", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = objecttype;
                uds = oForm.DataSources.UserDataSources.Add("sysobj", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                uds.Value = sysobj ? "Y" : "N";

                uds = oForm.DataSources.UserDataSources.Add("Status", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                uds.Value = "W";
                uds = oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 0);
                uds.Value = "-1";
                uds = oForm.DataSources.UserDataSources.Add("WddCode", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER, 0);
                uds.Value = "0";

                uds = oForm.DataSources.UserDataSources.Add("U_APP", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                uds.Value = "N";

                SAPbouiCOM.Item oItem = oForm.Items.Item("Approval");
                oItem.DisplayDesc = true;
                SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
                oComboBox.DataBind.SetBound(true, "", "U_APP");
                oComboBox.ValidValues.Add("W", "Pending");
                oComboBox.ValidValues.Add("Y", "Approved");
                oComboBox.ValidValues.Add("N", "Reject");

                oForm.DataSources.DataTables.Add("cfl");

                UserForm_Approval.retrievedata(oForm);
                //SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                //SAPbouiCOM.Condition oCondition = oConditions.Add();

                //oCondition.Alias = "U_APPROVAL";
                //oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //oCondition.CondVal = "P";
                //oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                //oCondition = oConditions.Add();
                //if (sysobj)
                //    oCondition.Alias = "DocStatus";
                //else
                //    oCondition.Alias = "Status";
                //oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //oCondition.CondVal = "O";

                //SAPbouiCOM.DBDataSource ds = null;

                //ds = oForm.DataSources.DBDataSources.Add(table);
                //ds.Query(oConditions);

                SAPbobsCOM.UserTable oUT = null;
                SAPbobsCOM.UserFields oUF = null;
                SAPbobsCOM.Documents doc = null;

                if (sysobj)
                {
                    doc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), objecttype));
                    oUF = doc.UserFields;
                }
                else
                {
                    oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(objecttype);
                    oUF = oUT.UserFields;
                }

                oItem = oForm.Items.Item("DocNum");
                SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, table, "DocNum");

                if (sysobj)
                {
                    oItem = oForm.Items.Item("DocDate");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    oEditText.DataBind.SetBound(true, table, "DocDate");
                }
                else
                {
                    oItem = oForm.Items.Item("DocDate");
                    oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                    oEditText.DataBind.SetBound(true, table, "U_DOCDATE");
                }

                //if (oUF.Fields.Item("U_APPROVAL").ValidValues.Count > 0)
                //{
                //    for (int y = 0; y < oUF.Fields.Item("U_APPROVAL").ValidValues.Count; y++)
                //    {
                //        oComboBox.ValidValues.Add(oUF.Fields.Item("U_APPROVAL").ValidValues.Item(y).Value, oUF.Fields.Item("U_APPROVAL").ValidValues.Item(y).Description);
                //    }
                //}


                oItem = oForm.Items.Item("ApprovalRe");
                oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                oEditText.DataBind.SetBound(true, table, "U_APPRE");

                oItem = oForm.Items.Item("Status");
                oItem.DisplayDesc = true;
                oItem.AffectsFormMode = false;
                oComboBox = (SAPbouiCOM.ComboBox)oItem.Specific;
                oComboBox.DataBind.SetBound(true, "", "Status");
                oComboBox.ValidValues.Add("W", "Pending");
                oComboBox.ValidValues.Add("Y", "Approved");
                oComboBox.ValidValues.Add("N", "Reject");
                //if (oUF.Fields.Item("U_APPROVAL").ValidValues.Count > 0)
                //{
                //    for (int y = 0; y < oUF.Fields.Item("U_APPROVAL").ValidValues.Count; y++)
                //    {
                //        oComboBox.ValidValues.Add(oUF.Fields.Item("U_APPROVAL").ValidValues.Item(y).Value, oUF.Fields.Item("U_APPROVAL").ValidValues.Item(y).Description);
                //    }
                //}
                oComboBox.Select("W", SAPbouiCOM.BoSearchKey.psk_ByValue);

                oForm.PaneLevel = 1;

                oForm.Freeze(false);

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize Completed", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void approvallist(string FormUID, string objecttype, string table, string title, bool sysobj)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                string path = System.Windows.Forms.Application.StartupPath;
                xmlDoc.Load(path + "\\APPROVAL.srf");

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.XmlData = xmlDoc.InnerXml;     // Load form from xml 
                creationPackage.FormType = FormUID;
                //creationPackage.ObjectType = objecttype;

                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = title;

                oForm.Freeze(true);
                SAPbouiCOM.UserDataSource uds = null;
                uds = oForm.DataSources.UserDataSources.Add("table", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = table;
                uds = oForm.DataSources.UserDataSources.Add("objecttype", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = objecttype;
                uds = oForm.DataSources.UserDataSources.Add("sysobj", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);
                uds.Value = sysobj ? "Y" : "N";


                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCondition = oConditions.Add();

                oCondition.Alias = "U_APPROVAL";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "P";
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCondition = oConditions.Add();
                if (sysobj)
                    oCondition.Alias = "DocStatus";
                else
                    oCondition.Alias = "Status";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "O";

                SAPbouiCOM.DBDataSource ds = null;

                ds = oForm.DataSources.DBDataSources.Add(table);
                ds.Query(oConditions);

                SAPbobsCOM.UserTable oUT = null;
                SAPbobsCOM.UserFields oUF = null;
                SAPbobsCOM.Documents doc = null;
                                   
                if (sysobj)
                {
                    doc = (SAPbobsCOM.Documents)SAP.SBOCompany.GetBusinessObject((SAPbobsCOM.BoObjectTypes)Enum.Parse(typeof(SAPbobsCOM.BoObjectTypes), objecttype));
                    oUF = doc.UserFields;
                }
                else
                {
                    oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(objecttype);
                    oUF = oUT.UserFields;
                }

                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
                SAPbouiCOM.Column oColumn = null;

                oColumn = oMatrix.Columns.Item(1);
                oMatrix.Columns.Remove(1);

                oColumn = oMatrix.Columns.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oColumn.Editable = false;
                oColumn.Visible = false;
                oColumn.DataBind.SetBound(true, table, "DocEntry");

                oColumn = oMatrix.Columns.Add("DocNum", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                oColumn.Editable = false;
                oColumn.DataBind.SetBound(true, table, "DocNum");
                oColumn.TitleObject.Caption = "Doc No";
                oColumn.Width = 100;

                if (sysobj)
                {
                    oColumn = oMatrix.Columns.Add("DocDate", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.Editable = false;
                    oColumn.DataBind.SetBound(true, table, "DocDate");
                    oColumn.TitleObject.Caption = "Doc Date";
                    oColumn.Width = 100;

                }
                else
                {
                    oColumn = oMatrix.Columns.Add("DOCDATE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.Editable = false;
                    oColumn.DataBind.SetBound(true, table, "U_DOCDATE");
                    oColumn.TitleObject.Caption = "Doc Date";
                    oColumn.Width = 100;

                }

                oColumn = oMatrix.Columns.Add("U_APPROVAL", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                oColumn.Editable = true;
                oColumn.DataBind.SetBound(true, table, "U_APPROVAL");
                oColumn.TitleObject.Caption = oUF.Fields.Item("U_APPROVAL").Description;
                oColumn.Width = 100;
                oColumn.DisplayDesc = true;
                if (oUF.Fields.Item("U_APPROVAL").ValidValues.Count > 0)
                {
                    for (int y = 0; y < oUF.Fields.Item("U_APPROVAL").ValidValues.Count; y++)
                    {
                        oColumn.ValidValues.Add(oUF.Fields.Item("U_APPROVAL").ValidValues.Item(y).Value, oUF.Fields.Item("U_APPROVAL").ValidValues.Item(y).Description);
                    }
                }

                oColumn = oMatrix.Columns.Add("U_APPRE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                oColumn.Editable = true;
                oColumn.DataBind.SetBound(true, table, "U_APPRE");
                oColumn.TitleObject.Caption = oUF.Fields.Item("U_APPRE").Description;
                oColumn.Width = 200;


                oMatrix.LoadFromDataSource();
                oForm.Freeze(false);

                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void FT_SPLAN(string dsname, string dsname1, string dsnameb)
        {
            try
            {

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = dsname;
                creationPackage.ObjectType = dsname;

                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);

                SAPbobsCOM.Recordset oSeq = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "select * from OUDO where Code = '" + dsname + "'";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();

                oForm.Title = oSeq.Fields.Item(1).Value.ToString();

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("dsname", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = dsname;
                uds = oForm.DataSources.UserDataSources.Add("dsname1", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = dsname1;
                uds = oForm.DataSources.UserDataSources.Add("dsnameb", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = dsnameb;

                uds = oForm.DataSources.UserDataSources.Add("FormUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = oForm.UniqueID;
                uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = "";
                uds = oForm.DataSources.UserDataSources.Add("docstatus", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = "";

                oForm.AutoManaged = true;

                //oForm.Left = oSForm.Left;
                oForm.Width = 1000;
                //oForm.Top = oSForm.Top;
                oForm.Height = 600;

                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Button oButton = null;
                SAPbouiCOM.EditText oEdit = null;
                SAPbouiCOM.ComboBox oCombo = null;
                SAPbouiCOM.LinkedButton oLinkedButton = null; 

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                if (dsnameb != "")
                {
                    oItem = oForm.Items.Add("cb_batch", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    oItem.Left = 150;
                    oItem.Width = 65;
                    oItem.Top = oForm.Height - 60;
                    oItem.Height = 20;
                    oButton = (SAPbouiCOM.Button)oItem.Specific;
                    oButton.Caption = "Batch/Bin";
                }
                if (dsname == "FT_TPPLAN")
                {
                    oItem = oForm.Items.Add("cb_trans", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                    oItem.Left = oForm.Width - 240;
                    oItem.Width = 100;
                    oItem.Top = oForm.Height - 60;
                    oItem.Height = 20;
                    oButton = (SAPbouiCOM.Button)oItem.Specific;
                    oButton.Caption = "Transfer Request";
                }
                sql = "select U_UDO, U_Btn, U_BtnName, U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Header = 'N' and isnull(U_HColumn,'') = ''";
                oSeq.DoQuery(sql);
                if (oSeq.RecordCount > 0)
                {
                    oSeq.MoveFirst();
                    SAPbouiCOM.ButtonCombo oButtonCombo = null;

                    oItem = oForm.Items.Add("cb_Copy", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    oItem.DisplayDesc = true;
                    oItem.Left = oForm.Width - 120;
                    oItem.Width = 80;
                    oItem.Top = oForm.Items.Item("2").Top;
                    oItem.Height = oForm.Items.Item("2").Height;
                    oItem.AffectsFormMode = false;
                    oButtonCombo = (SAPbouiCOM.ButtonCombo)oItem.Specific;
                    oButtonCombo.Caption = "Copy From";

                    while (!oSeq.EoF)
                    {
                        oButtonCombo.ValidValues.Add(oSeq.Fields.Item(1).Value.ToString(), oSeq.Fields.Item(2).Value.ToString());
                        oSeq.MoveNext();
                    }
                    //oButtonCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
               
                SAPbouiCOM.Conditions oBPCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oBPCon = oBPCons.Add();

                oBPCon.Alias = "CardType";
                oBPCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oBPCon.CondVal = "C";
                oBPCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oBPCon = oBPCons.Add();
                oBPCon.Alias = "FrozenFor";
                oBPCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oBPCon.CondVal = "N";

                SAPbouiCOM.Conditions oItemCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oItemCon = oItemCons.Add();
                oItemCon.Alias = "SellItem";
                oItemCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oItemCon.CondVal = "Y";


                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;

                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLitem";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oItemCons);


                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "2";
                oCFLCreationParams.UniqueID = "CFLCust";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oBPCons);

                SAPbouiCOM.Conditions oWHSCons = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oWHSCon = oWHSCons.Add();

                oWHSCon.Alias = "U_DropShip";
                oWHSCon.Operation = SAPbouiCOM.BoConditionOperation.co_NOT_EQUAL;
                oWHSCon.CondVal = "Y";

                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLwhsH";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oWHSCons);

                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFLwhs";
                oCFL = oCFLs.Add(oCFLCreationParams);
                oCFL.SetConditions(oWHSCons);

                //oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                //oCFLCreationParams.MultiSelection = false;
                //oCFLCreationParams.ObjectType = "TRANSPORTER";
                //oCFLCreationParams.UniqueID = "CFLTrans";
                //oCFL = oCFLs.Add(oCFLCreationParams);

                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                SAPbouiCOM.BoFormItemTypes linktype = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON;
                SAPbouiCOM.BoFormItemTypes itemtypecmb = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX;
                SAPbouiCOM.DataTable dt = null;// oForm.DataSources.DataTables.Add("99");

                SAPbouiCOM.Column oColumn = null;
                //SAPbouiCOM.UserDataSource oUds = null;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                SAPbobsCOM.UserTable oUT = null;
                SAPbobsCOM.UserFields oUF = null;
                string linkedtable = "";
                SAPbouiCOM.Matrix oMatrix = null;

                int top = -15;
                int cnt = -1;
                string columnname = "";
                Boolean end = false;

                oForm.DataSources.DBDataSources.Add("@" + dsname);

                top = top + 20;

                oItem = oForm.Items.Add("8", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Doc No";
                oItem.Left = 5;
                oItem.Top = top;
                oItem.Width = 150;
                oItem.Height = 15;

                columnname = "series";
                oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                ((SAPbouiCOM.ComboBox)oItem.Specific).DataBind.SetBound(true, "@" + dsname, columnname);
                oItem.Left = 155;
                oItem.Top = top;
                oItem.Width = 100;
                oItem.Height = 15;
                oItem.LinkTo = ("8").ToString();
                oItem.DisplayDesc = true;
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                #region get series
                oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("series").Specific;
                SAPbobsCOM.CompanyService companyServices = SAP.SBOCompany.GetCompanyService();
                SAPbobsCOM.SeriesService seriesService = (SAPbobsCOM.SeriesService)companyServices.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService);
                SAPbobsCOM.DocumentTypeParams oDocumentTypeParams = (SAPbobsCOM.DocumentTypeParams)seriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams);
                oDocumentTypeParams.Document = dsname;

                SAPbobsCOM.Series oSeries = seriesService.GetDefaultSeries(oDocumentTypeParams);
                int defSeries = 0;
                if (oSeries.Series != 0)
                {
                    defSeries = oSeries.Series;
                }
                //oForm.DataSources.UserDataSources.Add("DEFSERIES", SAPbouiCOM.BoDataType.dt_LONG_NUMBER).ValueEx = defSeries.ToString();
                //SAPbobsCOM.PathAdmin oPathAdmin = companyServices.GetPathAdmin();
                //oForm.DataSources.UserDataSources.Add("ATTACH", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 500).ValueEx = oPathAdmin.AttachmentsFolderPath;
                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCondition = oConditions.Add();

                oCondition.Alias = "ObjectCode";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = dsname;
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCondition = oConditions.Add();
                oCondition.Alias = "Locked";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "N";

                SAPbouiCOM.DBDataSource nnm1 = oForm.DataSources.DBDataSources.Add("NNM1");
                nnm1.Query(oConditions);

                companyServices = null;
                seriesService = null;
                oDocumentTypeParams = null;
                oSeries = null;
                uds = oForm.DataSources.UserDataSources.Add("defseries", SAPbouiCOM.BoDataType.dt_SHORT_NUMBER);
                if (nnm1.Size > 0)
                {
                    for (int i = 0; i < nnm1.Size; i++)
                    {
                        nnm1.Offset = i;
                        oCombo.ValidValues.Add(nnm1.GetValue("Series", i), nnm1.GetValue("SeriesName", i));
                        if (defSeries == 0)
                        {
                            if (i == 0)
                            {
                                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                uds.Value = "0";
                                //oForm.DataSources.DBDataSources.Item("@FT_CCF").SetValue("DocNum", 0, nnm1.GetValue("NextNumber", 0));
                            }
                        }
                        else
                        {
                            if (nnm1.GetValue("Series", i) == defSeries.ToString())
                            {
                                oCombo.Select(i, SAPbouiCOM.BoSearchKey.psk_Index);
                                uds.Value = i.ToString();
                            }
                        }
                    }
                }

                #endregion

                #region header
                columnname = "docnum";
                oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "@" + dsname, columnname);
                oItem.Left = 255;
                oItem.Top = top;
                oItem.Width = 150;
                oItem.Height = 15;
                oItem.LinkTo = ("8").ToString();
                //oItem.Enabled = false;

                oForm.AutoManaged = true;
                oForm.DataBrowser.BrowseBy = "docnum";
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;

                oItem = oForm.Items.Add("9", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Status";
                oItem.Left = oForm.Width - 60 - 155 - 155;
                oItem.Top = top;
                oItem.Width = 150;
                oItem.Height = 15;

                columnname = "docstatus";
                oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "docstatus");
                oItem.Left = oForm.Width - 60 - 155;
                oItem.Top = top;
                oItem.Width = 150;
                oItem.Height = 15;
                oItem.LinkTo = ("9").ToString();
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                //oSeq = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                sql = "select AliasID, rtABLE from CUFD where TableID = '@" + dsname + "' and isnull(U_seq,0) >= 0 order by U_seq, FieldID";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();

                string aliasid = "";
                bool btnfound = false;
                SAPbobsCOM.Recordset oBtn = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                while (!oSeq.EoF)
                {
                    aliasid = "U_" + oSeq.Fields.Item(0).Value.ToString();

                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        columnname = oUF.Fields.Item(x).Name;
                        switch (columnname)
                        {
                            case "U_APP":
                            case "U_APPRE":
                            case "U_APPBY":
                            case "U_APPDATE":
                            case "U_APPTIME":
                                continue;
                            default:
                                break;
                        }

                        if (columnname != aliasid) continue;

                        sql = "select U_Btn, U_BtnName, U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname + "' and U_Header = 'Y' and U_HColumn = '" + columnname + "'";
                        oBtn.DoQuery(sql);
                        btnfound = false;
                        if (oBtn.RecordCount > 0)
                        {
                            oBtn.MoveFirst();
                            oItem = oForm.Items.Add(oBtn.Fields.Item(0).Value.ToString(), SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                            oButton = (SAPbouiCOM.Button)oItem.Specific;
                            oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                            oButton.Image = Application.StartupPath + @"\CFL.BMP";
                            btnfound = true;
                        }

                        cnt++;

                        if (cnt % 2 == 0)
                        {
                            top = top + 20;
                            end = false;
                        }
                        else
                            end = true;

                        oItem = oForm.Items.Add((x * 10).ToString(), SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        ((SAPbouiCOM.StaticText)oItem.Specific).Caption = oUF.Fields.Item(x).Description;
                        if (end)
                            oItem.Left = oForm.Width - 60 - 155 - 155;
                        else
                            oItem.Left = 5;
                        oItem.Top = top;
                        oItem.Width = 150;
                        oItem.Height = 15;

                        linkedtable = oSeq.Fields.Item(1).Value.ToString().Trim();
                        //linkedtable = oUF.Fields.Item(x).LinkedTable;
                        if (linkedtable != "")
                        {
                            oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oItem.DisplayDesc = true;
                            oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                            oCombo.DataBind.SetBound(true, "@" + dsname, columnname);

                            dt = oForm.DataSources.DataTables.Add(linkedtable);
                            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                            for (int y = 0; y < dt.Rows.Count; y++)
                            {
                                ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                            }
                        }
                        else if (oUF.Fields.Item(x).ValidValues.Count > 0)
                        {
                            oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                            oCombo.DataBind.SetBound(true, "@" + dsname, columnname);
                            oItem.DisplayDesc = true;

                            for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                            {
                                ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                            }
                        }
                        else
                        {
                            oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                            oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                            oEdit.DataBind.SetBound(true, "@" + dsname, columnname);

                            if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                            {
                                //oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            }
                            if (columnname == "U_ADDRESS")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_REMARKS")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_FREMARKS")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                            if (columnname == "U_DOCDATE")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_CARDNAME")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_DOCTOTAL")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            //if (columnname == "U_APRICE")
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_TPDOCNUM")
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            if (columnname == "U_CARDCODE")
                            {
                                //oItemtemp = oForm.Items.Add("c" + columnname.Substring(1, columnname.Length - 1), SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                                //if (end)
                                //    oItemtemp.Left = oForm.Width - 66;
                                //else
                                //    oItemtemp.Left = 155 + 149;
                                //oItemtemp.Top = top - 2;
                                //oItemtemp.Width = 20;
                                //oItemtemp.Height = 20;
                                //oButton = ((SAPbouiCOM.Button)(oItemtemp.Specific));
                                //oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                                //oButton.Image = Application.StartupPath + @"\CFL.BMP";
                                //oCFLCreationParams.UniqueID = "c" + columnname.Substring(1, columnname.Length - 1);
                                //oCFL = oCFLs.Add(oCFLCreationParams);
                                //oButton.ChooseFromListUID = "c" + columnname.Substring(1, columnname.Length - 1);
                                oEdit.ChooseFromListUID = "CFLCust";
                                oEdit.ChooseFromListAlias = "CardCode";
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                                SAPbouiCOM.Item oLinkItem = oForm.Items.Add("LBP", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                oLinkItem.LinkTo = "U_CARDCODE";
                                ((SAPbouiCOM.LinkedButton)oLinkItem.Specific).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                            }
                            if (columnname == "U_WHSCODE")
                            {
                                oEdit.ChooseFromListUID = "CFLwhsH";
                                oEdit.ChooseFromListAlias = "WhsCode";
                                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                                SAPbouiCOM.Item oLinkItem = oForm.Items.Add("LWH", SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                                //oLinkItem.Top = oForm.Items.Item("U_WHSCODE").Top;
                                //oLinkItem.Left = oForm.Items.Item("U_WHSCODE").Left - 10;
                                oLinkItem.LinkTo = "U_WHSCODE";
                                ((SAPbouiCOM.LinkedButton)oLinkItem.Specific).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                            }
                            //if (columnname == "U_TPCODE")
                            //{
                            //    oEdit.ChooseFromListUID = "CFLTrans";
                            //    oEdit.ChooseFromListAlias = "Name";
                            //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            //}
                        }

                        if (end)
                            oItem.Left = oForm.Width - 60 - 155;
                        else
                            oItem.Left = 155;
                        oItem.Top = top;
                        oItem.Width = 150;
                        oItem.Height = 15;
                        oItem.LinkTo = (x * 10).ToString();
                        if (columnname == "U_CARDCODE")
                        {
                            oForm.Items.Item("LBP").Top = oForm.Items.Item("U_CARDCODE").Top;
                            oForm.Items.Item("LBP").Left = oForm.Items.Item("U_CARDCODE").Left - 20;
                        }
                        if (columnname == "U_WHSCODE")
                        {
                            oForm.Items.Item("LWH").Top = oForm.Items.Item("U_WHSCODE").Top;
                            oForm.Items.Item("LWH").Left = oForm.Items.Item("U_WHSCODE").Left - 20;
                        }

                        if (btnfound)
                        {
                            oButton.Item.Top = oItem.Top - 2;
                            oButton.Item.Left = oItem.Left + oItem.Width;
                            oButton.Item.Width = 15;
                        }
                        //switch (columnname)
                        //{
                        //    case "U_shipper":
                        //    case "U_itemdesc":
                        //    case "U_booking":
                        //    case "U_country":
                        //        oItem = oForm.Items.Add("c" + columnname.Substring(1, columnname.Length - 1), SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                        //        if (end)
                        //            oItem.Left = oForm.Width - 66;
                        //        else
                        //            oItem.Left = 155 + 149;
                        //        oItem.Top = top - 2;
                        //        oItem.Width = 20;
                        //        oItem.Height = 20;
                        //        oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                        //        oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                        //        oButton.Image = Application.StartupPath + @"\CFL.BMP";
                        //        oCFLCreationParams.UniqueID = "c" + columnname.Substring(1, columnname.Length - 1);
                        //        oCFL = oCFLs.Add(oCFLCreationParams);
                        //        oButton.ChooseFromListUID = "c" + columnname.Substring(1, columnname.Length - 1);

                        //        break;
                        //    case "":
                        //        oCFLCreationParams.UniqueID = "c" + columnname.Substring(1, columnname.Length - 1);
                        //        oCFL = oCFLs.Add(oCFLCreationParams);
                        //        oEdit.ChooseFromListUID = "c" + columnname.Substring(1, columnname.Length - 1);
                        //        //oEdit.ChooseFromListAlias = "U_booking";
                        //        break;
                        //    case "U_consigne":
                        //    case "U_notify":
                        //    case "U_loading":
                        //    case "U_discharg":
                        //        break;
                        //}
                    }
                    oSeq.MoveNext();
                }
                #endregion
                #region folder
                //top = top + 20;
                //oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                //SAPbouiCOM.Folder oFolder = null;

                //oItem = oForm.Items.Add("fgrid2", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                //oItem.Left = 5;
                //oItem.Width = 100;
                //oItem.Top = top;
                //oItem.Height = 19;
                //oItem.AffectsFormMode = false;
                //oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                //oFolder.Caption = "Container Details";
                //oFolder.DataBind.SetBound(true, "", "FolderDS");
                //oFolder.Select();

                //oItem = oForm.Items.Add("fgrid3", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                //oItem.Left = 105;
                //oItem.Width = 100;
                //oItem.Top = top;
                //oItem.Height = 19;
                //oItem.AffectsFormMode = false;
                //oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                //oFolder.Caption = "COA Result";
                //oFolder.DataBind.SetBound(true, "", "FolderDS");
                //oFolder.GroupWith("fgrid2");
                #endregion
                top = top + 20;

                #region grid 1
                oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = top;
                oItem.Height = oForm.Height - top - 65;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

                string dtldsname = dsname + "1";
                oForm.DataSources.DBDataSources.Add("@" + dtldsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dtldsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);
                oColumn.Editable = false;

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dtldsname);
                oUF = oUT.UserFields;

                sql = "select AliasID from CUFD where TableID = '@" + dtldsname + "' order by U_seq, FieldID";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();

                aliasid = "";

                while (!oSeq.EoF)
                {
                    aliasid = "U_" + oSeq.Fields.Item(0).Value.ToString();
                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        columnname = oUF.Fields.Item(x).Name;
                        if (columnname != aliasid) continue;

                        linkedtable = oUF.Fields.Item(x).LinkedTable;
                        if (linkedtable == "")
                        {
                            //if (columnname == "U_size")
                            //    linkedtable = "0009";
                            //else if (columnname == "U_jcclr")
                            //    linkedtable = "0007";
                            //else if (columnname == "U_brand")
                            //    linkedtable = "0004";
                            //else if (columnname == "U_perfcl")
                            //    linkedtable = "0001";
                        }
                        if (linkedtable != "")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                            oColumn.DisplayDesc = true;
                            dt = oForm.DataSources.DataTables.Add(linkedtable);
                            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                            for (int y = 0; y < dt.Rows.Count; y++)
                            {
                                oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                            }

                        }
                        else if (oUF.Fields.Item(x).ValidValues.Count > 0)
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                            oColumn.DisplayDesc = true;
                            oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);
                            for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                            {
                                oColumn.ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                            }

                            //oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                            //oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                            //oCombo.DataBind.SetBound(true, "@" + dsname, columnname);
                            //oItem.DisplayDesc = true;

                            //for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                            //{
                            //    ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                            //}
                        }
                        else if (columnname == "U_SOITEMCO" || columnname == "U_ITEMCODE")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, linktype);
                            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                        }
                        else if (columnname == "U_WHSCODE")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, linktype);
                            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses;
                        }
                        else if (columnname == "U_CARDCODE")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, linktype);
                            ((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_BusinessPartner;
                        }
                        else
                        {
                            //if (columnname == "U_itemcode")
                            //{
                            //    oColumn = oMatrix.Columns.Add(columnname, linktype);
                            //}
                            //else
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                        }
                        oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                        if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                            oColumn.Width = 200;
                        else
                            oColumn.Width = 100;

                        oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);

                        if (columnname == "U_LTOTAL")
                            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
                        else if (columnname == "U_WHSCODE")
                        {
                            oColumn.ChooseFromListUID = "CFLwhs";
                            oColumn.ChooseFromListAlias = "WhsCode";
                            //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                        }
                        else if (columnname == "U_ITEMCODE")
                        {
                            oColumn.ChooseFromListUID = "CFLitem";
                            oColumn.ChooseFromListAlias = "ItemCode";
                            
                        }
                        else
                        {
                            sql = "select U_BtnSQL from [@FT_SPCFSQL] where U_UDO = '" + dsname1 + "' and U_Header = 'N' and U_HColumn = '" + columnname + "'";
                            oBtn.DoQuery(sql);
                            if (oBtn.RecordCount > 0)
                            {
                                oCFLCreationParams.UniqueID = columnname;
                                oCFL = oCFLs.Add(oCFLCreationParams);
                                oColumn.ChooseFromListUID = columnname;
                            }

                        }
                        //else if (columnname == "U_ITEMCODE")
                        //{
                        //    oColumn.ChooseFromListUID = "CFLitem";
                        //    oColumn.ChooseFromListAlias = "ItemCode";
                        //    //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                        //}


                    }
                    oSeq.MoveNext();
                }
                #endregion

                //#region grid 1
                //oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                //oItem.Left = 5;
                //oItem.Width = oForm.Width - 25;
                //oItem.Top = top;
                //oItem.Height = oForm.Height - top - 65;
                //oItem.FromPane = 1;
                //oItem.ToPane = 1;
                //oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                //string dtldsname = dsname + "1";
                //oForm.DataSources.DBDataSources.Add("@" + dtldsname);
                //datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dtldsname).Fields.Item("VisOrder").Type);
                ////datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                //columnname = "VisOrder";
                //oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //oColumn.TitleObject.Caption = "#";
                //oColumn.Width = 20;
                ////oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                //oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);
                //oColumn.Editable = false;

                //oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dtldsname);
                //oUF = oUT.UserFields;
                //for (int x = 0; x < oUF.Fields.Count; x++)
                //{
                //    columnname = oUF.Fields.Item(x).Name;

                //    linkedtable = oUF.Fields.Item(x).LinkedTable;
                //    //if (linkedtable == "")
                //    //{
                //    //    //if (columnname == "U_size")
                //    //    //    linkedtable = "0009";
                //    //    //else if (columnname == "U_jcclr")
                //    //    //    linkedtable = "0007";
                //    //    //else if (columnname == "U_brand")
                //    //    //    linkedtable = "0004";
                //    //    //else if (columnname == "U_perfcl")
                //    //    //    linkedtable = "0001";
                //    //}
                //    if (linkedtable != "")
                //    {
                //        oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                //        oColumn.DisplayDesc = true;
                //        dt = oForm.DataSources.DataTables.Add(linkedtable);
                //        dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                //        for (int y = 0; y < dt.Rows.Count; y++)
                //        {
                //            oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                //        }

                //    }
                //    else if (oUF.Fields.Item(x).ValidValues.Count > 0)
                //    {
                //        oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                //        oColumn.DisplayDesc = true;
                //        oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);
                //        for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                //        {
                //            oColumn.ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                //        }

                //        //oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                //        //oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                //        //oCombo.DataBind.SetBound(true, "@" + dsname, columnname);
                //        //oItem.DisplayDesc = true;

                //        //for (int y = 0; y < oUF.Fields.Item(x).ValidValues.Count; y++)
                //        //{
                //        //    ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(oUF.Fields.Item(x).ValidValues.Item(y).Value, oUF.Fields.Item(x).ValidValues.Item(y).Description);
                //        //}
                //    }

                //    else
                //    {
                //        //if (columnname == "U_itemcode")
                //        //{
                //        //    oColumn = oMatrix.Columns.Add(columnname, linktype);
                //        //}
                //        //else
                //        oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //    }
                //    oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                //    if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                //        oColumn.Width = 200;
                //    else
                //        oColumn.Width = 100;

                //    oColumn.DataBind.SetBound(true, "@" + dtldsname, columnname);

                //    if (columnname == "U_ITEMCODE")
                //    {
                //        oColumn.ChooseFromListUID = "CFLitem";
                //        //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                //    }

                //    if (columnname == "U_WHSCODE")
                //    {
                //        oColumn.ChooseFromListUID = "CFLwhs";
                //        //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                //    }

                //}
                //#endregion
                ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dtldsname);
                #region grid 2
                //oItem = oForm.Items.Add("grid2", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                //oItem.Left = 5;
                //oItem.Width = oForm.Width - 25;
                //oItem.Top = top;
                //oItem.Height = oForm.Height - top - 60;
                //oItem.FromPane = 2;
                //oItem.ToPane = 2;
                //oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                //dsname = "FT_SPLAN1";
                //oForm.DataSources.DBDataSources.Add("@" + dsname);
                //datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                ////datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                //columnname = "VisOrder";
                //oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //oColumn.TitleObject.Caption = "#";
                //oColumn.Width = 20;
                ////oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                //oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                //oColumn.Editable = false;

                //oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                //oUF = oUT.UserFields;

                //// aaaaaa
                //for (int x = 0; x < oUF.Fields.Count; x++)
                //{
                //    columnname = oUF.Fields.Item(x).Name;
                //    {
                //        if (columnname == "U_conqty" || columnname == "U_conuom" || columnname == "U_conno" || columnname == "U_itemcode" || columnname == "U_itemname" || columnname == "U_size" || columnname == "U_jcclr" || columnname == "U_brand" || columnname == "U_perfcl")
                //        {
                //        }
                //        else
                //            continue;

                //        linkedtable = oUF.Fields.Item(x).LinkedTable;
                //        if (linkedtable == "")
                //        {
                //            if (columnname == "U_size")
                //                linkedtable = "0009";
                //            if (columnname == "U_jcclr")
                //                linkedtable = "0007";
                //            if (columnname == "U_brand")
                //                linkedtable = "0004";
                //            if (columnname == "U_perfcl")
                //                linkedtable = "0001";
                //        }
                //        if (linkedtable != "")
                //        {
                //            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                //            oColumn.DisplayDesc = true;
                //            dt = oForm.DataSources.DataTables.Add(linkedtable);
                //            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                //            for (int y = 0; y < dt.Rows.Count; y++)
                //            {
                //                oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                //            }
                //        }
                //        else
                //        {
                //            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //        }
                //        datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                //        oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                //        if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                //            oColumn.Width = 200;
                //        else
                //            oColumn.Width = 100;

                //        oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                //        if (columnname == "U_conno")
                //        {
                //            oColumn.ChooseFromListUID = "CFLcon";
                //        }
                //        if (columnname == "U_item1")
                //        {
                //            oColumn.ChooseFromListUID = "CFLitem1";
                //        }
                //        if (columnname == "U_item2")
                //        {
                //            oColumn.ChooseFromListUID = "CFLitem2";
                //        }
                //        if (columnname == "U_docentry" || columnname == "U_lineid")
                //            oColumn.Editable = false;
                //    }

                //}

                //// bbbbbb
                //for (int x = 0; x < oUF.Fields.Count; x++)
                //{
                //    columnname = oUF.Fields.Item(x).Name;
                //    {
                //        if (columnname == "U_conqty" || columnname == "U_conuom" || columnname == "U_conno" || columnname == "U_itemcode" || columnname == "U_itemname" || columnname == "U_size" || columnname == "U_jcclr" || columnname == "U_brand" || columnname == "U_perfcl")
                //            continue;

                //        linkedtable = oUF.Fields.Item(x).LinkedTable;
                //        if (linkedtable == "")
                //        {
                //            if (columnname == "U_consize")
                //                linkedtable = "0014";
                //        }
                //        if (linkedtable != "")
                //        {
                //            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                //            oColumn.DisplayDesc = true;
                //            dt = oForm.DataSources.DataTables.Add(linkedtable);
                //            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                //            for (int y = 0; y < dt.Rows.Count; y++)
                //            {
                //                oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                //            }
                //        }
                //        else
                //        {
                //            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //        }
                //        datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                //        oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                //        if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                //            oColumn.Width = 200;
                //        else
                //            oColumn.Width = 100;

                //        oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                //        if (columnname == "U_conno")
                //        {
                //            oColumn.ChooseFromListUID = "CFLcon";
                //        }
                //        if (columnname == "U_item1")
                //        {
                //            oColumn.ChooseFromListUID = "CFLitem1";
                //        }
                //        if (columnname == "U_item2")
                //        {
                //            oColumn.ChooseFromListUID = "CFLitem2";
                //        }
                //        if (columnname == "U_docentry" || columnname == "U_lineid")
                //            oColumn.Editable = false;
                //    }

                //}
                #endregion
                #region grid 3
                //oItem = oForm.Items.Add("grid3", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                //oItem.Left = 5;
                //oItem.Width = oForm.Width - 25;
                //oItem.Top = top;
                //oItem.Height = oForm.Height - top - 60;
                //oItem.FromPane = 3;
                //oItem.ToPane = 3;
                //oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                //dsname = "FT_SHIP3";
                //oForm.DataSources.DBDataSources.Add("@" + dsname);
                //datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                ////datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                //columnname = "VisOrder";
                //oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //oColumn.TitleObject.Caption = "#";
                //oColumn.Width = 20;
                ////oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                //oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                //oColumn.Editable = false;

                //oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                //oUF = oUT.UserFields;
                //for (int x = 0; x < oUF.Fields.Count; x++)
                //{
                //    columnname = oUF.Fields.Item(x).Name;

                //    linkedtable = oUF.Fields.Item(x).LinkedTable;
                //    if (columnname == "U_prodtype")
                //    {
                //        linkedtable = "0010";
                //    }
                //    if (linkedtable != "")
                //    {
                //        oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                //        oColumn.DisplayDesc = true;
                //        dt = oForm.DataSources.DataTables.Add(linkedtable);
                //        dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                //        for (int y = 0; y < dt.Rows.Count; y++)
                //        {
                //            oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                //        }
                //    }
                //    else
                //    {
                //        oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //    }
                //    datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                //    oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                //    if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                //        oColumn.Width = 200;
                //    else
                //        oColumn.Width = 100;

                //    oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                //}
                #endregion


                #region condition retrieve
                //SAPbouiCOM.Condition oCon = null;
                //SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                //SAPbouiCOM.Conditions oCons1 = new SAPbouiCOM.Conditions();

                //SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                //if (shipdocentry > 0)
                //{
                //    oCon = oCons1.Add();
                //    oCon.BracketOpenNum = 1;
                //    oCon.Alias = "docentry";
                //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //    oCon.CondVal = shipdocentry.ToString();
                //    oCon.BracketCloseNum = 1;

                //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons1);
                //    oForm.DataSources.DBDataSources.Item("@FT_SHIP1").Query(oCons1);
                //    oForm.DataSources.DBDataSources.Item("@FT_SHIP2").Query(oCons1);
                //    oForm.DataSources.DBDataSources.Item("@FT_SHIP3").Query(oCons1);
                //    ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                //    ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                //    ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();

                //    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                //}
                //else
                //{
                //    oCon = oCons.Add();
                //    oCon.BracketOpenNum = 1;
                //    oCon.Alias = "U_sdoc";
                //    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //    oCon.CondVal = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                //    oCon.BracketCloseNum = 1;

                //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons);

                //    if (oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Size == 0)
                //    {
                //        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                //        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_sdoc", 0, docentry.ToString());
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_pino", 0, docnum);

                //        oForm.DataSources.DBDataSources.Item("@FT_SHIP1").SetValue("VisOrder", 0, "1");
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIP2").SetValue("VisOrder", 0, "1");
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIP3").SetValue("VisOrder", 0, "1");
                //        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                //        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                //        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();
                //        //((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).AddRow(1, -1);                   
                //        //((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).AddRow(1, -1);
                //        //((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).AddRow(1, -1);

                //        rs.DoQuery("select max(U_set) from [@FT_SHIPD] where U_sdoc = " + docentry.ToString());
                //        if (rs.RecordCount > 0)
                //        {
                //            rs.MoveFirst();
                //            int set = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                //            oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_set", 0, set.ToString());
                //        }
                //        //rs.DoQuery("select max(docentry) from [@FT_SHIPD]");
                //        //if (rs.RecordCount > 0)
                //        //{
                //        //    long docnum = long.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                //        //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("docnum", 0, docnum.ToString());
                //        //}
                //    }
                //    else
                //    {
                //        oCon = oCons1.Add();
                //        oCon.BracketOpenNum = 1;
                //        oCon.Alias = "docentry";
                //        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                //        oCon.CondVal = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("docentry", 0).ToString();
                //        oCon.BracketCloseNum = 1;

                //        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons1);
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIP1").Query(oCons1);
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIP2").Query(oCons1);
                //        oForm.DataSources.DBDataSources.Item("@FT_SHIP3").Query(oCons1);
                //        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                //        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                //        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();

                //        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                //    }
                //}
                //rs = null;
                #endregion
                //oForm.Items.Item("fgrid2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                //oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized;
                //oSForm.Freeze(true);


                oForm.EnableFormatSearch();
                oForm.DataBrowser.BrowseBy = "docnum";

                oForm.PaneLevel = 1;
                oForm.Visible = true;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;
                UserForm_SalesPlanning.AddNew(oForm);

                //oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }

        public static void batchno(string FormUID, string DetailDS, string BatchDS)
        {
            SAPbouiCOM.Form oSForm = null;
            oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);

            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
            string path = System.Windows.Forms.Application.StartupPath;
            xmlDoc.Load(path + "\\BATCHNO.srf");

            SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
            creationPackage.XmlData = xmlDoc.InnerXml;     // Load form from xml 

            SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);

            oForm.Freeze(true);

            #region with bin
            string sql = @"SELECT 
t01.ItemCode, 
t01.WhsCode, 
t01.BinCode,
t01.BinAbs,
t01.BatchNum,
(case when t01.BatchNum = '' then TotalStockBin else t01.TotalStockBinBatch end) as OnHand
from
(SELECT 
T0.ItemCode, 
T5.ItemName, 
T0.WhsCode, 
T1.BinCode,
T2.BinAbs,
isnull(T4.DistNumber,'') as BatchNum,
SUM(T3.OnHandQty) as 'TotalStockBinBatch',
SUM(T2.OnHandQty) as 'TotalStockBin',
SUM(T0.OnHand) AS 'TotalStockWarehouse'
FROM 
OITW T0
INNER JOIN OBIN T1 ON T0.WhsCode = T1.WhsCode
INNER JOIN OIBQ T2 ON T2.WhsCode = T0.WhsCode AND T1.AbsEntry = T2.BinAbs AND T0.ItemCode = T2.ItemCode
left JOIN OBBQ T3 ON T3.ItemCode = T0.ItemCode AND T3.BinAbs = T1.AbsEntry AND T3.WhsCode = T2.WhsCode
left JOIN OBTN T4 ON T4.AbsEntry = T3.SnBMDAbs
INNER JOIN OITM T5 ON T5.ItemCode = T0.ItemCode
INNER JOIN OWHS ON OWHS.BinActivat = 'Y' and OWHS.WhsCode = T0.WhsCode
WHERE 
T2.OnHandQty > 0 
or T3.OnHandQty > 0
GROUP BY 
T0.ItemCode, T5.ItemName, T0.WhsCode, 
T0.WhsCode, T1.BinCode, T2.BinAbs, T4.DistNumber
) t01";

            SAPbouiCOM.UserDataSource oUDS = null;
            oUDS = oForm.DataSources.UserDataSources.Add("sqlbatch2", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
            oUDS.ValueEx = sql;
            #endregion

            #region without bin
            sql = "select t01.ItemCode, t01.WhsCode, '' as BinCode, 0 as BinAbs, t01.BatchNum, t01.OnHand from (select T0.[ItemCode], T0.[WhsCode], T0.[BatchNum], SUM(CASE T0.[Direction] when 0 then 1 else -1 end * T0.[Quantity]) as [OnHand] FROM IBT1 T0 inner join OITM T9 on T9.ItemCode = T0.ItemCode and T9.InvntItem = 'Y' and T9.SellItem = 'Y' and T9.ManBtchNum = 'Y' group by T0.[ItemCode], T0.[WhsCode], T0.[BatchNum]) t01";

            oUDS = oForm.DataSources.UserDataSources.Add("sqlbatch1", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
            oUDS.ValueEx = sql;
            #endregion

            oUDS = oForm.DataSources.UserDataSources.Add("sform", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            oUDS.ValueEx = FormUID;
            oUDS = oForm.DataSources.UserDataSources.Add("sdetailds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            oUDS.ValueEx = DetailDS;
            oUDS = oForm.DataSources.UserDataSources.Add("sbatchds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
            oUDS.ValueEx = BatchDS;

            SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Add("ITEMBATCH");
            //oDT.ExecuteQuery(sql + " where 1 = 0" + sqlgroup);

            oForm.DataSources.DBDataSources.Add(DetailDS);
            oForm.DataSources.DBDataSources.Add(BatchDS);
            oForm.DataSources.DBDataSources.Add("@FT_BATCH");

            #region grid1_FT_BATCH
            SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;
            SAPbouiCOM.Column oColumn = null;

            oColumn = oMatrix.Columns.Item(0);
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "Code");

            oColumn = oMatrix.Columns.Add("U_ITEMCODE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Item Code";
            oColumn.Width = 80;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_ITEMCODE");
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("U_WHSCODE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Code";
            oColumn.Width = 100;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_WHSCODE");
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("U_BIN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Bin";
            oColumn.Width = 130;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_BIN");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_BINABS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Bin #";
            oColumn.Width = 30;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_BINABS");
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("U_BATCHNUM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Batch No";
            oColumn.Width = 80;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_BATCHNUM");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_ONHAND", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "On Hand";
            oColumn.Width = 50;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_ONHAND");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_QUANTITY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Quantity";
            oColumn.Width = 50;
            oColumn.DataBind.SetBound(true, "@FT_BATCH", "U_QUANTITY");
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            oColumn.AffectsFormMode = false;
            #endregion

            #region grid2_BatchDS
            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific;
            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            oColumn = oMatrix.Columns.Item(0);
            oColumn.DataBind.SetBound(true, BatchDS, "U_BASEVIS");

            oColumn = oMatrix.Columns.Add("U_ITEMCODE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Item Code";
            oColumn.Width = 80;
            oColumn.DataBind.SetBound(true, BatchDS, "U_ITEMCODE");
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("U_WHSCODE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Code";
            oColumn.Width = 100;
            oColumn.DataBind.SetBound(true, BatchDS, "U_WHSCODE");
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("U_BIN", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Bin";
            oColumn.Width = 150;
            oColumn.DataBind.SetBound(true, BatchDS, "U_BIN");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_BINABS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Bin #";
            oColumn.Width = 30;
            oColumn.DataBind.SetBound(true, BatchDS, "U_BINABS");
            oColumn.Editable = false;
            oColumn.Visible = false;

            oColumn = oMatrix.Columns.Add("U_BATCHNUM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Batch No";
            oColumn.Width = 80;
            oColumn.DataBind.SetBound(true, BatchDS, "U_BATCHNUM");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_QUANTITY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Quantity";
            oColumn.Width = 50;
            oColumn.DataBind.SetBound(true, BatchDS, "U_QUANTITY");
            oColumn.Editable = false;
            oColumn.ColumnSetting.SumType = SAPbouiCOM.BoColumnSumType.bst_Auto;
            #endregion

            #region grid3_DetailDS
            oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific;
            oColumn = oMatrix.Columns.Item(0);
            oColumn.DataBind.SetBound(true, DetailDS, "VisOrder");

            oColumn = oMatrix.Columns.Add("U_ITEMCODE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Item Code";
            oColumn.Width = 100;
            oColumn.DataBind.SetBound(true, DetailDS, "U_ITEMCODE");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_ITEMNAME", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Item Name";
            oColumn.Width = 300;
            oColumn.DataBind.SetBound(true, DetailDS, "U_ITEMNAME");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_WHSCODE", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Warehouse Code";
            oColumn.Width = 100;
            oColumn.DataBind.SetBound(true, DetailDS, "U_WHSCODE");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_QUANTITY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Quantity";
            oColumn.Width = 100;
            oColumn.DataBind.SetBound(true, DetailDS, "U_QUANTITY");
            oColumn.Editable = false;

            oColumn = oMatrix.Columns.Add("U_BQTY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
            oColumn.TitleObject.Caption = "Batch Selected";
            oColumn.Width = 100;
            oColumn.DataBind.SetBound(true, DetailDS, "U_BQTY");
            oColumn.Editable = false;
            #endregion

            oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            SAPbouiCOM.DBDataSource dbs = oForm.DataSources.DBDataSources.Item(DetailDS);
            SAPbouiCOM.DBDataSource fdbs = oSForm.DataSources.DBDataSources.Item(DetailDS);
            for (int y = 0; y < fdbs.Size; y++)
            {
                if (y > 0)
                {
                    dbs.InsertRecord(dbs.Size - 1);
                }

                dbs.SetValue("VisOrder", dbs.Size - 1, fdbs.GetValue("VisOrder", y).Trim());
                dbs.SetValue("U_ITEMCODE", dbs.Size - 1, fdbs.GetValue("U_ITEMCODE", y).Trim());
                dbs.SetValue("U_ITEMNAME", dbs.Size - 1, fdbs.GetValue("U_ITEMNAME", y).Trim());
                dbs.SetValue("U_WHSCODE", dbs.Size - 1, fdbs.GetValue("U_WHSCODE", y).Trim());
                dbs.SetValue("U_QUANTITY", dbs.Size - 1, fdbs.GetValue("U_QUANTITY", y).Trim());
                dbs.SetValue("U_BQTY", dbs.Size - 1, fdbs.GetValue("U_BQTY", y).Trim());
            }


            oMatrix.LoadFromDataSource();
            oMatrix.SelectRow(1, true, false);

            UserForm_Batch.retrieveBatch(oForm);

            //oForm.AutoManaged = true;
            oForm.Freeze(false);

        }

        public static void shipingadvice(string FormUID, bool SourceForm)
        {
            try
            {
                string dsname = "FT_APSA";

                SAPbouiCOM.Form oSForm = null;
                if (SourceForm)
                {
                    oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);
                }

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                System.Xml.XmlDocument xmlDoc = new System.Xml.XmlDocument();
                string path = System.Windows.Forms.Application.StartupPath;
                xmlDoc.Load(path + "\\APSA.srf");

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.XmlData = xmlDoc.InnerXml;     // Load form from xml 
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);

                oForm.Freeze(true);
                //oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized;
                //oSForm.Freeze(true);

                SAPbouiCOM.UserDataSource uds = null;
                uds = oForm.DataSources.UserDataSources.Add("fuid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                SAPbouiCOM.DBDataSource dbs = oForm.DataSources.DBDataSources.Item("@" + dsname);
                SAPbouiCOM.DataTable dt = null;// (SAPbouiCOM.DataTable)oForm.DataSources.DataTables.Item("DT_NNM1");


                //SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("series").Specific;

                //for (int x = 0; x < dt.Rows.Count; x++)
                //{
                //    oComboBox.ValidValues.Add(dt.GetValue(0, x).ToString(), dt.GetValue(1, x).ToString());
                //    if (dt.GetValue(2, x).ToString() == "Default")
                //    {
                //        dbs.SetValue("Series", 0, dt.GetValue(0, x).ToString());
                //    }
                //}
                ///////////////////////
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("series").Specific;
                SAPbobsCOM.CompanyService companyServices = SAP.SBOCompany.GetCompanyService();
                SAPbobsCOM.SeriesService seriesService = (SAPbobsCOM.SeriesService)companyServices.GetBusinessService(SAPbobsCOM.ServiceTypes.SeriesService);
                SAPbobsCOM.DocumentTypeParams oDocumentTypeParams = (SAPbobsCOM.DocumentTypeParams)seriesService.GetDataInterface(SAPbobsCOM.SeriesServiceDataInterfaces.ssdiDocumentTypeParams);
                oDocumentTypeParams.Document = dsname;

                SAPbobsCOM.Series oSeries = seriesService.GetDefaultSeries(oDocumentTypeParams);
                int defSeries = 0;
                if (oSeries.Series != 0)
                {
                    defSeries = oSeries.Series;
                }
                //oForm.DataSources.UserDataSources.Add("DEFSERIES", SAPbouiCOM.BoDataType.dt_LONG_NUMBER).ValueEx = defSeries.ToString();
                //SAPbobsCOM.PathAdmin oPathAdmin = companyServices.GetPathAdmin();
                //oForm.DataSources.UserDataSources.Add("ATTACH", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 500).ValueEx = oPathAdmin.AttachmentsFolderPath;
                SAPbouiCOM.Conditions oConditions = new SAPbouiCOM.Conditions();
                SAPbouiCOM.Condition oCondition = oConditions.Add();

                oCondition.Alias = "ObjectCode";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = dsname;
                oCondition.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                oCondition = oConditions.Add();
                oCondition.Alias = "Locked";
                oCondition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCondition.CondVal = "N";

                SAPbouiCOM.DBDataSource nnm1 = oForm.DataSources.DBDataSources.Add("NNM1");
                nnm1.Query(oConditions);

                companyServices = null;
                seriesService = null;
                oDocumentTypeParams = null;
                oSeries = null;

                if (nnm1.Size > 0)
                {
                    for (int i = 0; i < nnm1.Size; i++)
                    {
                        nnm1.Offset = i;
                        oCombo.ValidValues.Add(nnm1.GetValue("Series", i), nnm1.GetValue("SeriesName", i));
                        if (defSeries == 0)
                        {
                            if (i == 0)
                            {
                                oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                //oForm.DataSources.DBDataSources.Item("@FT_CCF").SetValue("DocNum", 0, nnm1.GetValue("NextNumber", 0));
                            }
                        }
                        else
                        {
                            if (nnm1.GetValue("Series", i) == defSeries.ToString())
                            {
                                oCombo.Select(i, SAPbouiCOM.BoSearchKey.psk_Index);
                            }
                        }
                    }
                }
                ///////////////////////
                string linkedtable = "";
                string columnname = "";

                if (SourceForm)
                {
                    dbs.SetValue("U_CARDCODE", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("CardCode", 0).Trim());
                    dbs.SetValue("U_CARDNAME", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("CardName", 0).Trim());
                    dbs.SetValue("U_SODOCENTRY", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocEntry", 0).Trim());
                    dbs.SetValue("U_SODOCNUM", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("DocNum", 0).Trim());
                    dbs.SetValue("U_NUMATCARD", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("NumAtCard", 0).Trim());
                    //dbs.SetValue("U_PRICE", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_PRICE", 0));
                    //dbs.SetValue("U_BookNo", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_BookNo", 0));
                    //dbs.SetValue("U_CONT", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_CONT", 0));
                    //dbs.SetValue("U_CONTNO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_CONTNO", 0));
                    //dbs.SetValue("U_BVMOH", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_BVMOH", 0));
                    //dbs.SetValue("U_BVMOHNO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_BVMOHNO", 0));
                    //dbs.SetValue("U_MOH", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_MOH", 0));
                    //dbs.SetValue("U_MOHNO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_MOHNO", 0));
                    //dbs.SetValue("U_PCert", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_PCert", 0));
                    //dbs.SetValue("U_PCertNo", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_PCertNo", 0));
                    //dbs.SetValue("U_SGS", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_SGS", 0));
                    //dbs.SetValue("U_SGSNO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_SGSNO", 0));
                    //dbs.SetValue("U_GCHEM", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_GCHEM", 0));
                    //dbs.SetValue("U_GCHEMNO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_GCHEMNO", 0));
                    //dbs.SetValue("U_IIS", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_IIS", 0));
                    //dbs.SetValue("U_IISNO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_IISNO", 0));
                    //dbs.SetValue("U_COA", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_COA", 0));
                    //dbs.SetValue("U_COANO", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_COANO", 0));
                    //dbs.SetValue("U_Remarks", 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue("U_Remarks", 0));

                    SAPbobsCOM.UserTable oUT = null;
                    SAPbobsCOM.UserFields oUF = null;
                    //SAPbobsCOM.UserFieldsMD oUFMD = (SAPbobsCOM.UserFieldsMD)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUserFields);
                    oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                    oUF = oUT.UserFields;

                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        //oUFMD.GetByKey("@" + dsname, oUF.Fields.Item(x).FieldID);
                        //columnname = "U_" + oUFMD.Name;
                        columnname = oUF.Fields.Item(x).Name;

                        switch (columnname)
                        {
                            case "U_CARDCODE":
                            case "U_CARDNAME":
                            case "U_SODOCENTRY":
                            case "U_SODOCNUM":
                            case "U_NUMATCARD":
                            case "U_DOCDATE":
                            case "U_DODATE":
                            case "U_DODOCNUM":
                                continue;
                        }
                        dbs.SetValue(columnname, 0, oSForm.DataSources.DBDataSources.Item("ORDR").GetValue(columnname, 0).Trim());
                    }
                }



                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                //SAPbouiCOM.BoFormItemTypes linktype = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON;
                SAPbouiCOM.BoFormItemTypes itemtypecmb = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX;

                SAPbouiCOM.Column oColumn = null;
                //SAPbouiCOM.UserDataSource oUds = null;
                SAPbouiCOM.Matrix oMatrix = null;

                oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific;

                dsname = "FT_APSA1";
                oForm.DataSources.DBDataSources.Add("@" + dsname);
                //datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                ////datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                //columnname = "VisOrder";
                //oColumn = oMatrix.Columns.Add(columnname, itemtype);
                //oColumn.TitleObject.Caption = "#";
                //oColumn.Width = 20;
                ////oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                //oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                //oColumn.Editable = false;

                oColumn = oMatrix.Columns.Item(0);
                oColumn.DataBind.SetBound(true, "@" + dsname, "VisOrder");

                //oMatrix.Columns.Item("V_0").Visible = false;
                //oMatrix.Columns.Item("V_-1").Visible = false;

                SAPbouiCOM.DataTable dtUF = oForm.DataSources.DataTables.Add("CUFD");
                dtUF.ExecuteQuery("select 'U_' + AliasID, Descr, isnull(RTable,'') from CUFD where tableid = '@" + dsname + "' order by u_seq, fieldid");

                SAPbouiCOM.ChooseFromListCollection oCFLs = oForm.ChooseFromLists;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = (SAPbouiCOM.ChooseFromListCreationParams)SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "64";
                oCFLCreationParams.UniqueID = "CFL_WHS";

                SAPbouiCOM.ChooseFromList oCFL = oCFLs.Add(oCFLCreationParams);
                SAPbouiCOM.Conditions oCons = oCFL.GetConditions();
                SAPbouiCOM.Condition oCon = oCons.Add();

                oCon.Alias = "Inactive";
                oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                oCon.CondVal = "N";

                oCFL.SetConditions(oCons);

                SAPbouiCOM.LinkedButton oLink;

                //SAPbobsCOM.UserTable oUT = null;
                //SAPbobsCOM.UserFields oUF = null;
                //oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                //oUF = oUT.UserFields;
                for (int x = 0; x < dtUF.Rows.Count; x++)
                {
                    columnname = dtUF.GetValue(0, x).ToString();
                    //switch (columnname)
                    //{
                    //    case "U_SOENTRY":
                    //    case "U_SOLINE":
                    //        continue;
                    //}
                    linkedtable = "";
                    linkedtable = dtUF.GetValue(2, x).ToString();
                    if (linkedtable != "")
                    {
                        oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                        oColumn.DisplayDesc = true;
                        dt = oForm.DataSources.DataTables.Add(linkedtable);
                        dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                        for (int y = 0; y < dt.Rows.Count; y++)
                        {
                            oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                        }

                    }
                    else
                    {
                        //if (columnname == "U_WHSCODE" || columnname == "U_ITEMCODE")
                        //    oColumn = oMatrix.Columns.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON);
                        //else
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                    }
                    oColumn.TitleObject.Caption = dtUF.GetValue(1, x).ToString();
                    oColumn.Width = 100;

                    oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                    switch (columnname)
                    {
                        case "U_SOENTRY":
                        case "U_SOLINE":
                            oColumn.Editable = false;
                            oColumn.Visible = false;
                            break;
                        case "U_BookNo":
                        case "U_ITEMNAME":
                        case "U_OPENQTY":
                            oColumn.Editable = false;
                            break;
                        case "U_QUANTITY":                           
                            //((SAPbouiCOM.Item)oColumn.ExtendedObject).SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 2, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                            break;
                        case "U_WHSCODE":
                            oColumn.ChooseFromListUID = "CFL_WHS";
                            oColumn.ChooseFromListAlias = "WhsCode";
                            //oLink = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                            //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Warehouses;

                            //oColumn.Editable = false;
                            break;
                        case "U_ITEMCODE":
                        case "U_BQTY":
                            oColumn.Editable = false;

                            //oLink = (SAPbouiCOM.LinkedButton)oColumn.ExtendedObject;
                            //oLink.LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                            break;

                    }

                }

                if (SourceForm)
                {
                    dbs = oForm.DataSources.DBDataSources.Item("@" + dsname);
                    SAPbouiCOM.DBDataSource fdbs = oSForm.DataSources.DBDataSources.Item("RDR1");

                    int rows = fdbs.Size;
                    int cnt = 0;
                    double rs = 0;
                    for (int y = 0; y < rows - 1; y++)
                    {
                        if (fdbs.GetValue("OpenQty", y) == null) continue;

                        if (double.TryParse(fdbs.GetValue("OpenQty", y), out rs))
                        {
                            if (rs > 0)
                            {
                                cnt++;

                                if (cnt > 1)
                                {
                                    dbs.InsertRecord(dbs.Size - 1);
                                }
                                dbs.SetValue("VisOrder", dbs.Size - 1, cnt.ToString());

                                for (int x = 0; x < dtUF.Rows.Count; x++)
                                {
                                    columnname = dtUF.GetValue(0, x).ToString();
                                    switch (columnname)
                                    {
                                        case "U_WHSCODE":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("WhsCode", y).Trim());
                                            break;
                                        case "U_SOENTRY":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("DocEntry", y).Trim());
                                            break;
                                        case "U_SOLINE":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("LineNum", y).Trim());
                                            break;
                                        case "U_ITEMCODE":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("ItemCode", y).Trim());
                                            break;
                                        case "U_ITEMNAME":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("Dscription", y).Trim());
                                            break;
                                        case "U_OPENQTY":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("OpenQty", y).Trim());
                                            break;
                                        case "U_QUANTITY":
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue("OpenQty", y).Trim());
                                            break;
                                        case "U_BQTY":
                                            break;
                                        default:
                                            dbs.SetValue(columnname, dbs.Size - 1, fdbs.GetValue(columnname, y).Trim());
                                            break;
                                    }

                                }
                            }
                        }
                    }
                }
                oMatrix.LoadFromDataSource();

                SAPbouiCOM.Item oItem = oForm.Items.Item("cardcode");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                //oItem = oForm.Items.Item("bookno");
                //oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("cardname");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("cancel");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("numatcard");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("sodocnum");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("status");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("docnum");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("series");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 9, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                oItem = oForm.Items.Item("cb_gendo");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 1, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                oItem = oForm.Items.Item("cb_batch");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_True);
                oItem = oForm.Items.Item("dodocnum");
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, 11, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                oForm.EnableFormatSearch();
                oForm.AutoManaged = true;
                oForm.DataBrowser.BrowseBy = "docnum";

                oForm.Freeze(false);
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }


        }
        
        public static void shiplist(string FUID, long docentry)
        {
            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize Listing window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
            creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
            creationPackage.FormType = "FT_SHIPL";
            SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
            oForm.Title = "Shipping Document List";
            SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FUID);
            oForm.Left = oSForm.Left;
            oForm.Width = 600;
            oForm.Top = oSForm.Top;
            oForm.Height = 500;

            oForm.DataSources.UserDataSources.Add("FUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
            oForm.DataSources.UserDataSources.Add("sdoc", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);

            oForm.DataSources.UserDataSources.Item("FUID").Value = FUID;
            oForm.DataSources.UserDataSources.Item("sdoc").Value = docentry.ToString();

            SAPbouiCOM.Grid oGrid = null;
            SAPbouiCOM.Item oItem = null;

            oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
            oItem.Left = 5;
            oItem.Width = 65;
            oItem.Top = oForm.Height - 60;
            oItem.Height = 20;

            oItem = oForm.Items.Add("grid", SAPbouiCOM.BoFormItemTypes.it_GRID);
            oItem.Left = 5;
            oItem.Width = oForm.Width - 25;
            oItem.Top = 5;
            oItem.Height = oForm.Height - 70;

            oGrid = (SAPbouiCOM.Grid)oItem.Specific;
            oForm.DataSources.DataTables.Add("list");
            try
            {
                string sql = "select docentry, U_sdoc, U_pino, U_docdate, U_booking, U_set, U_shipper, U_consigne, U_notify, U_loading, U_discharg, U_vessel, U_country, U_itemdesc from [@FT_SHIPD] where U_sdoc = " + docentry.ToString();
                oForm.DataSources.DataTables.Item("list").ExecuteQuery(sql);
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }

            oGrid.DataTable = oForm.DataSources.DataTables.Item("list");

            foreach (SAPbouiCOM.GridColumn column in oGrid.Columns)
            {
                column.Editable = false;
                if (column.UniqueID == "docentry" || column.UniqueID == "U_sdoc")
                    column.Visible = false;
                else
                {
                    if (column.UniqueID == "U_pino")
                        column.TitleObject.Caption = "SC No";
                    if (column.UniqueID == "U_docdate")
                        column.TitleObject.Caption = "Date";
                }
            }

            oGrid.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Single;

            oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized;
            oSForm.Freeze(true);
            oForm.Visible = true;

            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_None);
        }
        public static void shipdoc(string FormUID, long docentry, long shipdocentry, string docnum)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_SHIPD";
                creationPackage.ObjectType = "FT_SHIPD";
                
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);

                oForm.Title = "Shipping Document";
                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);

                //oForm.Left = oSForm.Left;
                oForm.Width = 1000;
                //oForm.Top = oSForm.Top;
                oForm.Height = 600;

                SAPbouiCOM.Item oItem = null;
                SAPbouiCOM.Button oButton = null;
                SAPbouiCOM.EditText oEdit = null;
                SAPbouiCOM.ComboBox oCombo = null;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                uds.Value = "";

                //oItem = oForm.Items.Add("ROW", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "row");
                //oItem.Left = 55;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;

                //oItem = oForm.Items.Add("st_ROW", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                //((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Row No#";
                //oItem.Left = 5;
                //oItem.Width = 50;
                //oItem.Top = 5;
                //oItem.Height = 15;

                uds = oForm.DataSources.UserDataSources.Add("fuid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                uds = oForm.DataSources.UserDataSources.Add("docentry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();
                uds = oForm.DataSources.UserDataSources.Add("docnum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docnum;

                SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                oCFLs = oForm.ChooseFromLists;

                SAPbouiCOM.ChooseFromList oCFL = null;
                SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                oCFLCreationParams.MultiSelection = false;
                oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLitem";
                oCFL = oCFLs.Add(oCFLCreationParams);

                //oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));
                //oCFLCreationParams.MultiSelection = false;
                //oCFLCreationParams.ObjectType = "4";
                oCFLCreationParams.UniqueID = "CFLcon";
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCFLCreationParams.UniqueID = "CFLitem1";
                oCFL = oCFLs.Add(oCFLCreationParams);

                oCFLCreationParams.UniqueID = "CFLitem2";
                oCFL = oCFLs.Add(oCFLCreationParams);

                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                //SAPbouiCOM.BoFormItemTypes linktype = SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON;
                SAPbouiCOM.BoFormItemTypes itemtypecmb = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX;
                SAPbouiCOM.DataTable dt = null;// oForm.DataSources.DataTables.Add("99");

                SAPbouiCOM.Column oColumn = null;
                //SAPbouiCOM.UserDataSource oUds = null;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                SAPbobsCOM.UserTable oUT = null;
                SAPbobsCOM.UserFields oUF = null;
                string linkedtable = "";
                SAPbouiCOM.Matrix oMatrix = null;

                string dsname = "";
                int top = -15;
                int cnt = -1;
                string columnname = "";
                Boolean end = false;
                dsname = "FT_SHIPD";
                oForm.DataSources.DBDataSources.Add("@" + dsname);

                top = top + 20;

                columnname = "docnum";
                oItem = oForm.Items.Add("8", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Doc No";
                oItem.Left = 5;
                oItem.Top = top;
                oItem.Width = 150;
                oItem.Height = 15;

                oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "@" + dsname, columnname);
                oItem.Left = 155;
                oItem.Top = top;
                oItem.Width = 150;
                oItem.Height = 15;
                oItem.LinkTo = ("8").ToString();
                //oItem.Enabled = false;

                oForm.AutoManaged = true;
                oForm.DataBrowser.BrowseBy = "docnum";
                oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                //columnname = "U_sdoc";
                //oItem = oForm.Items.Add("9", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                //((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Docentry";
                //oItem.Left = oForm.Width - 20 - 155 - 155;
                //oItem.Top = top;
                //oItem.Width = 150;
                //oItem.Height = 15;

                //oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "@" + dsname, columnname);
                //oItem.Left = oForm.Width - 20 - 155;
                //oItem.Top = top;
                //oItem.Width = 150;
                //oItem.Height = 15;
                //oItem.LinkTo = ("9").ToString();
                //oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;

                SAPbobsCOM.Recordset oSeq = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                string sql = "select AliasID from CUFD where TableID = '@" + dsname + "' order by U_seq, FieldID";
                string aliasid = "";
                oSeq.DoQuery(sql);
                oSeq.MoveFirst();

                while (!oSeq.EoF)
                {
                    aliasid = "U_" + oSeq.Fields.Item(0).Value.ToString();

                    for (int x = 0; x < oUF.Fields.Count; x++)
                    {
                        columnname = oUF.Fields.Item(x).Name;

                        if (columnname != aliasid) continue;

                        if (columnname != "U_sdoc")
                        {
                            cnt++;

                            if (cnt % 2 == 0)
                            {
                                top = top + 20;
                                end = false;
                            }
                            else
                                end = true;

                            oItem = oForm.Items.Add((x * 10).ToString(), SAPbouiCOM.BoFormItemTypes.it_STATIC);
                            ((SAPbouiCOM.StaticText)oItem.Specific).Caption = oUF.Fields.Item(x).Description;
                            if (end)
                                oItem.Left = oForm.Width - 60 - 155 - 155;
                            else
                                oItem.Left = 5;
                            oItem.Top = top;
                            oItem.Width = 150;
                            oItem.Height = 15;

                            linkedtable = oUF.Fields.Item(x).LinkedTable;
                            if (linkedtable != "")
                            {
                                oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                                oCombo = (SAPbouiCOM.ComboBox)oItem.Specific;
                                oCombo.DataBind.SetBound(true, "@" + dsname, columnname);

                                dt = oForm.DataSources.DataTables.Add(linkedtable);
                                dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                                for (int y = 0; y < dt.Rows.Count; y++)
                                {
                                    ((SAPbouiCOM.ComboBox)oItem.Specific).ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                                }
                            }
                            else
                            {
                                oItem = oForm.Items.Add(columnname, SAPbouiCOM.BoFormItemTypes.it_EDIT);
                                oEdit = (SAPbouiCOM.EditText)oItem.Specific;
                                oEdit.DataBind.SetBound(true, "@" + dsname, columnname);

                                if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                                {
                                    //oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                                }
                                //if (columnname == "U_set")
                                //    oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, -1, SAPbouiCOM.BoModeVisualBehavior.mvb_False);
                            }

                            if (end)
                                oItem.Left = oForm.Width - 60 - 155;
                            else
                                oItem.Left = 155;
                            oItem.Top = top;
                            oItem.Width = 150;
                            oItem.Height = 15;
                            oItem.LinkTo = (x * 10).ToString();

                            switch (columnname)
                            {
                                case "U_shipper":
                                case "U_itemdesc":
                                case "U_booking":
                                case "U_country":
                                    oItem = oForm.Items.Add("c" + columnname.Substring(1, columnname.Length - 1), SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                                    if (end)
                                        oItem.Left = oForm.Width - 66;
                                    else
                                        oItem.Left = 155 + 149;
                                    oItem.Top = top - 2;
                                    oItem.Width = 20;
                                    oItem.Height = 20;
                                    oButton = ((SAPbouiCOM.Button)(oItem.Specific));
                                    oButton.Type = SAPbouiCOM.BoButtonTypes.bt_Image;
                                    oButton.Image = Application.StartupPath + @"\CFL.BMP";
                                    oCFLCreationParams.UniqueID = "c" + columnname.Substring(1, columnname.Length - 1);
                                    oCFL = oCFLs.Add(oCFLCreationParams);
                                    oButton.ChooseFromListUID = "c" + columnname.Substring(1, columnname.Length - 1);

                                    break;
                                case "":
                                    oCFLCreationParams.UniqueID = "c" + columnname.Substring(1, columnname.Length - 1);
                                    oCFL = oCFLs.Add(oCFLCreationParams);
                                    oEdit.ChooseFromListUID = "c" + columnname.Substring(1, columnname.Length - 1);
                                    //oEdit.ChooseFromListAlias = "U_booking";
                                    break;
                                case "U_consigne":
                                case "U_notify":
                                case "U_loading":
                                case "U_discharg":
                                    break;
                            }
                        }
                    }
                    oSeq.MoveNext();
                }

                top = top + 20;
                oForm.DataSources.UserDataSources.Add("FolderDS", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 1);

                SAPbouiCOM.Folder oFolder = null;
                //oItem = oForm.Items.Add("fgrid1", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                //oItem.Left = 5;
                //oItem.Width = 100;
                //oItem.Top = top;
                //oItem.Height = 19;
                //oItem.AffectsFormMode = false;
                //oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                //oFolder.Caption = "Item Details";
                //oFolder.DataBind.SetBound(true, "", "FolderDS");
                //oFolder.Select();

                oItem = oForm.Items.Add("fgrid2", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                oItem.Left = 5;
                oItem.Width = 100;
                oItem.Top = top;
                oItem.Height = 19;
                oItem.AffectsFormMode = false;
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oFolder.Caption = "Container Details";
                oFolder.DataBind.SetBound(true, "", "FolderDS");
                oFolder.Select();

                oItem = oForm.Items.Add("fgrid3", SAPbouiCOM.BoFormItemTypes.it_FOLDER);
                oItem.Left = 105;
                oItem.Width = 100;
                oItem.Top = top;
                oItem.Height = 19;
                oItem.AffectsFormMode = false;
                oFolder = (SAPbouiCOM.Folder)oItem.Specific;
                oFolder.Caption = "COA Result";
                oFolder.DataBind.SetBound(true, "", "FolderDS");
                oFolder.GroupWith("fgrid2");

                top = top + 20;

                //grid 1 - start
                oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = top;
                oItem.Height = oForm.Height - top - 60;
                oItem.FromPane = 1;
                oItem.ToPane = 1;
                oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
                
                dsname = "FT_SHIP1";
                oForm.DataSources.DBDataSources.Add("@" + dsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Editable = false;

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;
                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    columnname = oUF.Fields.Item(x).Name;

                    linkedtable = oUF.Fields.Item(x).LinkedTable;
                    if (linkedtable == "")
                    {
                        //if (columnname == "U_size")
                        //    linkedtable = "0009";
                        //else if (columnname == "U_jcclr")
                        //    linkedtable = "0007";
                        //else if (columnname == "U_brand")
                        //    linkedtable = "0004";
                        //else if (columnname == "U_perfcl")
                        //    linkedtable = "0001";
                    }
                    if (linkedtable != "")
                    {
                        oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                        oColumn.DisplayDesc = true;
                        dt = oForm.DataSources.DataTables.Add(linkedtable);
                        dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                        for (int y = 0; y < dt.Rows.Count; y++)
                        {
                            oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());                            
                        }
                        
                    }
                    else
                    {
                        //if (columnname == "U_itemcode")
                        //{
                        //    oColumn = oMatrix.Columns.Add(columnname, linktype);
                        //}
                        //else
                        oColumn = oMatrix.Columns.Add(columnname, itemtype);
                    }
                    oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                    if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                        oColumn.Width = 200;
                    else
                        oColumn.Width = 100;

                    oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                    if (columnname == "U_itemcode")
                    {
                        oColumn.ChooseFromListUID = "CFLitem";
                        //((SAPbouiCOM.LinkedButton)oColumn.ExtendedObject).LinkedObject = SAPbouiCOM.BoLinkedObject.lf_Items;
                    }

                }
                //grid 2 - start
                oItem = oForm.Items.Add("grid2", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = top;
                oItem.Height = oForm.Height - top - 60;
                oItem.FromPane = 2;
                oItem.ToPane = 2;
                oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                dsname = "FT_SHIP2";
                oForm.DataSources.DBDataSources.Add("@" + dsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Editable = false;

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;

                // aaaaaa
                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    columnname = oUF.Fields.Item(x).Name;
                    {
                        if (columnname == "U_conqty" || columnname == "U_conuom" || columnname == "U_conno" || columnname == "U_itemcode" || columnname == "U_itemname" || columnname == "U_size" || columnname == "U_jcclr" || columnname == "U_brand" || columnname == "U_perfcl")
                        {
                        }
                        else
                            continue;

                        linkedtable = oUF.Fields.Item(x).LinkedTable;
                        if (linkedtable == "")
                        {
                            if (columnname == "U_size")
                                linkedtable = "0009";
                            if (columnname == "U_jcclr")
                                linkedtable = "0007";
                            if (columnname == "U_brand")
                                linkedtable = "0004";
                            if (columnname == "U_perfcl")
                                linkedtable = "0001";
                        }
                        if (linkedtable != "")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                            oColumn.DisplayDesc = true;
                            dt = oForm.DataSources.DataTables.Add(linkedtable);
                            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                            for (int y = 0; y < dt.Rows.Count; y++)
                            {
                                oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                            }
                        }
                        else
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                        }
                        datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                        oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                        if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                            oColumn.Width = 200;
                        else
                            oColumn.Width = 100;

                        oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                        if (columnname == "U_conno")
                        {
                            oColumn.ChooseFromListUID = "CFLcon";
                        }
                        if (columnname == "U_item1")
                        {
                            oColumn.ChooseFromListUID = "CFLitem1";
                        }
                        if (columnname == "U_item2")
                        {
                            oColumn.ChooseFromListUID = "CFLitem2";
                        }
                        if (columnname == "U_docentry" || columnname == "U_lineid")
                            oColumn.Editable = false;
                    }

                }

                // bbbbbb
                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    columnname = oUF.Fields.Item(x).Name;
                    {
                        if (columnname == "U_conqty" || columnname == "U_conuom" || columnname == "U_conno" || columnname == "U_itemcode" || columnname == "U_itemname" || columnname == "U_size" || columnname == "U_jcclr" || columnname == "U_brand" || columnname == "U_perfcl")
                            continue;

                        linkedtable = oUF.Fields.Item(x).LinkedTable;
                        if (linkedtable == "")
                        {
                            if (columnname == "U_consize")
                                linkedtable = "0014";
                        }
                        if (linkedtable != "")
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                            oColumn.DisplayDesc = true;
                            dt = oForm.DataSources.DataTables.Add(linkedtable);
                            dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                            for (int y = 0; y < dt.Rows.Count; y++)
                            {
                                oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                            }
                        }
                        else
                        {
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                        }
                        datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                        oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                        if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                            oColumn.Width = 200;
                        else
                            oColumn.Width = 100;

                        oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                        if (columnname == "U_conno")
                        {
                            oColumn.ChooseFromListUID = "CFLcon";
                        }
                        if (columnname == "U_item1")
                        {
                            oColumn.ChooseFromListUID = "CFLitem1";
                        }
                        if (columnname == "U_item2")
                        {
                            oColumn.ChooseFromListUID = "CFLitem2";
                        }
                        if (columnname == "U_docentry" || columnname == "U_lineid")
                            oColumn.Editable = false;
                    }

                }
                //grid 3 - start
                oItem = oForm.Items.Add("grid3", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = top;
                oItem.Height = oForm.Height - top - 60;
                oItem.FromPane = 3;
                oItem.ToPane = 3;
                oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                dsname = "FT_SHIP3";
                oForm.DataSources.DBDataSources.Add("@" + dsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Editable = false;

                oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                oUF = oUT.UserFields;
                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    columnname = oUF.Fields.Item(x).Name;

                    linkedtable = oUF.Fields.Item(x).LinkedTable;
                    if (columnname == "U_prodtype")
                    {
                        linkedtable = "0010";
                    }
                    if (linkedtable != "")
                    {
                        oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                        oColumn.DisplayDesc = true;
                        dt = oForm.DataSources.DataTables.Add(linkedtable);
                        dt.ExecuteQuery("select code, name from [@" + linkedtable + "]");
                        for (int y = 0; y < dt.Rows.Count; y++)
                        {
                            oColumn.ValidValues.Add(dt.GetValue(0, y).ToString(), dt.GetValue(1, y).ToString());
                        }
                    }
                    else
                    {
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                    }
                    datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                    oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                    if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                        oColumn.Width = 200;
                    else
                        oColumn.Width = 100;

                    oColumn.DataBind.SetBound(true, "@" + dsname, columnname);

                }
                // grid - end
                //UserForm_CONmodified.retrieveRow(oForm, docentry, linenum, dsname);

                //ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dsname);

                SAPbouiCOM.Condition oCon = null;
                SAPbouiCOM.Conditions oCons = new SAPbouiCOM.Conditions();

                SAPbouiCOM.Conditions oCons1 = new SAPbouiCOM.Conditions();

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                if (shipdocentry > 0)
                {
                    oCon = oCons1.Add();
                    oCon.BracketOpenNum = 1;
                    oCon.Alias = "docentry";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = shipdocentry.ToString();
                    oCon.BracketCloseNum = 1;

                    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons1);
                    oForm.DataSources.DBDataSources.Item("@FT_SHIP1").Query(oCons1);
                    oForm.DataSources.DBDataSources.Item("@FT_SHIP2").Query(oCons1);
                    oForm.DataSources.DBDataSources.Item("@FT_SHIP3").Query(oCons1);
                    ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                    ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                    ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();

                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                }
                else
                {                    
                    oCon = oCons.Add();
                    oCon.BracketOpenNum = 1;
                    oCon.Alias = "U_sdoc";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = oForm.DataSources.UserDataSources.Item("docentry").Value.ToString();
                    oCon.BracketCloseNum = 1;

                    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons);

                    if (oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Size == 0)
                    {
                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_ADD_MODE;


                        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_sdoc", 0, docentry.ToString());
                        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_pino", 0, docnum);
                        
                        oForm.DataSources.DBDataSources.Item("@FT_SHIP1").SetValue("VisOrder", 0, "1");
                        oForm.DataSources.DBDataSources.Item("@FT_SHIP2").SetValue("VisOrder", 0, "1");
                        oForm.DataSources.DBDataSources.Item("@FT_SHIP3").SetValue("VisOrder", 0, "1");
                        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();
                        //((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).AddRow(1, -1);                   
                        //((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).AddRow(1, -1);
                        //((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).AddRow(1, -1);

                        rs.DoQuery("select max(U_set) from [@FT_SHIPD] where U_sdoc = " + docentry.ToString());
                        if (rs.RecordCount > 0)
                        {
                            rs.MoveFirst();
                            int set = int.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                            oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("U_set", 0, set.ToString());
                        }
                        //rs.DoQuery("select max(docentry) from [@FT_SHIPD]");
                        //if (rs.RecordCount > 0)
                        //{
                        //    long docnum = long.Parse(rs.Fields.Item(0).Value.ToString()) + 1;
                        //    oForm.DataSources.DBDataSources.Item("@FT_SHIPD").SetValue("docnum", 0, docnum.ToString());
                        //}
                    }
                    else
                    {
                        oCon = oCons1.Add();
                        oCon.BracketOpenNum = 1;
                        oCon.Alias = "docentry";
                        oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                        oCon.CondVal = oForm.DataSources.DBDataSources.Item("@FT_SHIPD").GetValue("docentry", 0).ToString();
                        oCon.BracketCloseNum = 1;

                        oForm.DataSources.DBDataSources.Item("@FT_SHIPD").Query(oCons1);
                        oForm.DataSources.DBDataSources.Item("@FT_SHIP1").Query(oCons1);
                        oForm.DataSources.DBDataSources.Item("@FT_SHIP2").Query(oCons1);
                        oForm.DataSources.DBDataSources.Item("@FT_SHIP3").Query(oCons1);
                        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid1").Specific).LoadFromDataSource();
                        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid2").Specific).LoadFromDataSource();
                        ((SAPbouiCOM.Matrix)oForm.Items.Item("grid3").Specific).LoadFromDataSource();

                        oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                    }
                }
                rs = null;

                oForm.Items.Item("fgrid2").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized;
                oSForm.Freeze(true);
                //oForm.PaneLevel = 1;
                oForm.Visible = true;
                //oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }
        public static void TEXT(string FormUID, long docentry, int linenum, int row, string dsname, string matrixname, string value)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_TEXT";
                creationPackage.BorderStyle = SAPbouiCOM.BoFormBorderStyle.fbs_Fixed;
                //creationPackage.FormType 
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = "Text Detail...";
                //oForm.Left = oSForm.Left;
                oForm.Width = 500;
                //oForm.Top = oSForm.Top;
                oForm.Height = 400;

                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("DocEntry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();

                //oItem = oForm.Items.Add("DocEntry", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).String = docentry.ToString();

                //oItem.Left = 140;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;
                //oItem.Enabled = true;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("LineNo", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = linenum.ToString();

                //oItem = oForm.Items.Add("LineNo", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).String = linenum.ToString();

                //oItem.Left = 150;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;
                //oItem.Enabled = true;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("FUID", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "fuid");
                //oItem.Left = 160;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("DSNAME", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = dsname;

                //oItem = oForm.Items.Add("DSNAME", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).String = dsname;
                //oItem.Left = 160;
                //oItem.Width = 10;
                //oItem.Top = oForm.Height - 60;
                //oItem.Height = 20;

                uds = oForm.DataSources.UserDataSources.Add("text", SAPbouiCOM.BoDataType.dt_LONG_TEXT);
                uds.Value = value;

                oItem = oForm.Items.Add("TEXT", SAPbouiCOM.BoFormItemTypes.it_EXTEDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "text");

                oItem.Left = 5;
                oItem.Width = oForm.Width - 30;
                oItem.Top = 5;
                oItem.Height = oForm.Height - 90;
                oItem.Enabled = true;
                oItem.Visible = true;

                //oForm.Items.Item("FUID").Visible = false;
                //oForm.Items.Item("LineNo").Visible = false;
                //oForm.Items.Item("DocEntry").Visible = false;
                //oForm.Items.Item("DSNAME").Visible = false;

                oForm.DataSources.DBDataSources.Add("INV1");

                oItem = oForm.Items.Item("1");
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";
                oForm.Visible = true;
                //oForm.Modal = true;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                //oForm
                //oSForm.State = SAPbouiCOM.BoFormStateEnum.;
                //oSForm.Freeze(true);

                //((SAPbouiCOM.EditText)oSForm.Items.Item("FUID").Specific).Value = oForm.UniqueID.ToString();
                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);
                oSForm.DataSources.UserDataSources.Item("cfluid").Value = oForm.UniqueID;

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
        public static void CONM(string FormUID, long docentry, int linenum, int row, string dsname, string mustcol, string bookno)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning); 
                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());

                creationPackage.FormType = "FT_CONM";

                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = "Container";
                //oForm.Left = oSForm.Left;
                oForm.Width = 600;
                //oForm.Top = oSForm.Top;
                oForm.Height = 500;

                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = row.ToString();

                uds = oForm.DataSources.UserDataSources.Add("bookno", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = bookno;

                oItem = oForm.Items.Add("ROW", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "row");
                oItem.Left = 55;
                oItem.Width = 65;
                oItem.Top = 5;
                oItem.Height = 15;
                oItem.Enabled = false;

                oItem = oForm.Items.Add("st_ROW", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Row No#";
                oItem.Left = 5;
                oItem.Width = 50;
                oItem.Top = 5;
                oItem.Height = 15;

                uds = oForm.DataSources.UserDataSources.Add("fuid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "fuid");
                //oItem.Left = 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("docentry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();

                //oItem = oForm.Items.Add("DOCENTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "docentry");
                //oItem.Left = 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("linenum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = linenum.ToString();

                //oItem = oForm.Items.Add("LINENUM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "linenum");
                //oItem.Left = 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("ds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = dsname;

                //oItem = oForm.Items.Add("DS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "ds");
                //oItem.Left = 5 + 65 + 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("mustcol", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = mustcol;

                //oItem = oForm.Items.Add("MUSTCOL", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "mustcol");
                //oItem.Left = 5 + 65 + 5 + 65 + 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = 25;
                oItem.Height = oForm.Height - 90;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                SAPbouiCOM.BoFormItemTypes itemtypecmb = SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX;

                SAPbouiCOM.Column oColumn = null;
                //SAPbouiCOM.UserDataSource oUds = null;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                string columnname = "";

                oForm.DataSources.DBDataSources.Add("@" + dsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Editable = false;
                
                SAPbobsCOM.UserTable oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                SAPbobsCOM.UserFields oUF = null;
                oUF = oUT.UserFields;
                string linkedtable = "";
                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    columnname = oUF.Fields.Item(x).Name;
                    //if (columnname == "U_DOCNO" || columnname == "U_LINENO")
                    //    continue;
                    linkedtable = oUF.Fields.Item(x).LinkedTable;
                    //if (linkedtable != "")
                    //{
                    //    oColumn = oMatrix.Columns.Add(columnname, itemtypecmb);
                    //}
                    //else
                    //{
                        oColumn = oMatrix.Columns.Add(columnname, itemtype);
                    //}
                    datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                    oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                    if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                        oColumn.Width = 200;
                    else
                        oColumn.Width = 100;
                    //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, oUF.Fields.Item(x).Size);
                    oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                    if (columnname == "U_LINENO")
                        oColumn.Visible = false;
                    
                    //else if (columnname == "U_CONNO")
                    //{
                    //    SAPbouiCOM.ChooseFromListCollection oCFLs = null;
                    //    oCFLs = oForm.ChooseFromLists;

                    //    SAPbouiCOM.ChooseFromList oCFL = null;
                    //    SAPbouiCOM.ChooseFromListCreationParams oCFLCreationParams = null;
                    //    oCFLCreationParams = ((SAPbouiCOM.ChooseFromListCreationParams)(SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams)));

                    //    oCFLCreationParams.MultiSelection = false;
                    //    oCFLCreationParams.ObjectType = "1";
                    //    oCFLCreationParams.UniqueID = "CFL1";

                    //    oCFL = oCFLs.Add(oCFLCreationParams);

                    //    oColumn.ChooseFromListUID = "CFL1";
                    //}
                    
                }
                /*
                datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "#";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "", columnname);
                oColumn.Editable = false;

                datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                columnname = "U_CONNO";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Container No";
                oColumn.Width = 100;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 100);
                oColumn.DataBind.SetBound(true, "", columnname);

                datatype = SAPbouiCOM.BoDataType.dt_DATE;
                columnname = "U_CONDATE";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Container Date";
                oColumn.Width = 100;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "", columnname);

                datatype = SAPbouiCOM.BoDataType.dt_LONG_TEXT;
                columnname = "U_REF";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Reference";
                oColumn.Width = 300;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 500);
                oColumn.DataBind.SetBound(true, "", columnname);

                datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "U_ITEMNO";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Item No";
                oColumn.Width = 100;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "", columnname);
                oColumn.Visible = false;
                */

                datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "LineId";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Item No";
                oColumn.Width = 100;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Visible = false;

                UserForm_CONmodified.retrieveRow(oForm, docentry, linenum, dsname);

                ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dsname);

                //uds = oForm.DataSources.UserDataSources.Add("cfluid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT, 50);
                //uds.Value = "";

                oForm.Visible = true;
                oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success); 

            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }
        public static void DOPTM(string FormUID, long docentry, int linenum, int row, string dsname, string mustcol)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_DOPTM";
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = "GRN Sales Planning Doc";
                //oForm.Left = oSForm.Left;
                oForm.Width = 600;
                //oForm.Top = oSForm.Top;
                oForm.Height = 500;

                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = row.ToString();

                oItem = oForm.Items.Add("ROW", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "row");
                oItem.Left = 55;
                oItem.Width = 65;
                oItem.Top = 5;
                oItem.Height = 15;
                oItem.Enabled = false;

                oItem = oForm.Items.Add("st_ROW", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Row No#";
                oItem.Left = 5;
                oItem.Width = 50;
                oItem.Top = 5;
                oItem.Height = 15;

                uds = oForm.DataSources.UserDataSources.Add("fuid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "fuid");
                //oItem.Left = 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("docentry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();

                //oItem = oForm.Items.Add("DOCENTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "docentry");
                //oItem.Left = 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("linenum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = linenum.ToString();

                //oItem = oForm.Items.Add("LINENUM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "linenum");
                //oItem.Left = 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("ds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = dsname;

                //oItem = oForm.Items.Add("DS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "ds");
                //oItem.Left = 5 + 65 + 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("mustcol", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = mustcol;

                //oItem = oForm.Items.Add("MUSTCOL", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "mustcol");
                //oItem.Left = 5 + 65 + 5 + 65 + 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = 25;
                oItem.Height = oForm.Height - 90;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                SAPbouiCOM.Column oColumn = null;
                //SAPbouiCOM.UserDataSource oUds = null;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                string columnname = "";

                oForm.DataSources.DBDataSources.Add("@" + dsname);
                datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(oForm.DataSources.DBDataSources.Item("@" + dsname).Fields.Item("VisOrder").Type);
                //datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "VisOrder";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Editable = false;

                SAPbobsCOM.UserTable oUT = (SAPbobsCOM.UserTable)SAP.SBOCompany.UserTables.Item(dsname);
                SAPbobsCOM.UserFields oUF = null;
                oUF = oUT.UserFields;
                for (int x = 0; x < oUF.Fields.Count; x++)
                {
                    columnname = oUF.Fields.Item(x).Name;
                    //if (columnname == "U_DOCNO" || columnname == "U_LINENO")
                    //    continue;
                    oColumn = oMatrix.Columns.Add(columnname, itemtype);
                    datatype = ObjectFunctions.changeDIFieldTypesToDIDataType(oUF.Fields.Item(x).Type);
                    oColumn.TitleObject.Caption = oUF.Fields.Item(x).Description;
                    if (oUF.Fields.Item(x).Type == SAPbobsCOM.BoFieldTypes.db_Memo)
                        oColumn.Width = 200;
                    else
                        oColumn.Width = 200;
                    //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, oUF.Fields.Item(x).Size);
                    oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                    if (columnname == "U_LINENO")
                        oColumn.Visible = false;
                }
                /*
                datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "#";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "#";
                oColumn.Width = 20;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "", columnname);
                oColumn.Editable = false;

                datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
                columnname = "U_CONNO";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Container No";
                oColumn.Width = 100;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 100);
                oColumn.DataBind.SetBound(true, "", columnname);

                datatype = SAPbouiCOM.BoDataType.dt_DATE;
                columnname = "U_CONDATE";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Container Date";
                oColumn.Width = 100;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "", columnname);

                datatype = SAPbouiCOM.BoDataType.dt_LONG_TEXT;
                columnname = "U_REF";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Reference";
                oColumn.Width = 300;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 500);
                oColumn.DataBind.SetBound(true, "", columnname);

                datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "U_ITEMNO";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Item No";
                oColumn.Width = 100;
                oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "", columnname);
                oColumn.Visible = false;
                */

                datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                columnname = "LineId";
                oColumn = oMatrix.Columns.Add(columnname, itemtype);
                oColumn.TitleObject.Caption = "Item No";
                oColumn.Width = 100;
                //oUds = oForm.DataSources.UserDataSources.Add(columnname, datatype, 0);
                oColumn.DataBind.SetBound(true, "@" + dsname, columnname);
                oColumn.Visible = false;

                UserForm_CONmodified.retrieveRow(oForm, docentry, linenum, dsname);

                //ObjectFunctions.customFormMatrixSetting(oForm, "grid1", SAP.SBOCompany.UserName, dsname);
                
                oForm.Visible = true;
                oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success); 

            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }
        public static void SOM(string FormUID, long docentry, int linenum, int row, string dsname, string matrixname)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_SOM";
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                oForm.Title = "Sales Order Modified";
                //oForm.Left = oSForm.Left;
                oForm.Width = 700;
                //oForm.Top = oSForm.Top;
                oForm.Height = 500;

                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = row.ToString();

                oItem = oForm.Items.Add("ROW", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "row");
                oItem.Left = 55;
                oItem.Width = 65;
                oItem.Top = 5;
                oItem.Height = 15;
                oItem.Enabled = false;

                oItem = oForm.Items.Add("st_ROW", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Row No#";
                oItem.Left = 5;
                oItem.Width = 50;
                oItem.Top = 5;
                oItem.Height = 15;

                uds = oForm.DataSources.UserDataSources.Add("fuid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "fuid");
                //oItem.Left = 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("docentry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();

                //oItem = oForm.Items.Add("DOCENTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "docentry");
                //oItem.Left = 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("linenum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = linenum.ToString();

                //oItem = oForm.Items.Add("LINENUM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "linenum");
                //oItem.Left = 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("ds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = dsname;

                //oItem = oForm.Items.Add("DS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "ds");
                //oItem.Left = 5 + 65 + 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                SAPbouiCOM.Matrix oSMatrix = (SAPbouiCOM.Matrix)oSForm.Items.Item(matrixname).Specific;

                oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = 25;
                oItem.Height = oForm.Height - 90;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;

                if (oSMatrix.RowCount > 0)
                {
                    SAPbouiCOM.Conditions oCons = (SAPbouiCOM.Conditions)SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Add(dsname);

                    oDS.Clear();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.BracketOpenNum = 1;
                    oCon.Alias = "DocEntry";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = docentry.ToString();
                    oCon.BracketCloseNum = 1;
                    oCon.Relationship = SAPbouiCOM.BoConditionRelationship.cr_AND;

                    oCon = oCons.Add();
                    oCon.BracketOpenNum = 1;
                    oCon.Alias = "DelivrdQty";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_GRATER_THAN;
                    oCon.CondVal = "0";
                    oCon.BracketCloseNum = 1;

                    SAPbouiCOM.Column oColumn = null;
                    oColumn = oMatrix.Columns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.DataBind.SetBound(true, dsname, "LineNum");
                    oColumn.Visible = false;

                    copyUDFMatrixColumns(oSForm, oSMatrix, oForm, oMatrix, dsname, dsname);

                    oDS.Query(oCons);
                    oMatrix.LoadFromDataSource();

                    string templinenum = "";
                    string temprow = "";
                    for (int x = 0; x < oSForm.DataSources.DBDataSources.Item(dsname).Size; x++)
                    {
                        if (decimal.Parse(oSForm.DataSources.DBDataSources.Item(dsname).GetValue("DelivrdQty", x).ToString()) > 0)
                        {
                            templinenum = oSForm.DataSources.DBDataSources.Item(dsname).GetValue("LineNum", x).ToString();
                            temprow = ((SAPbouiCOM.EditText)oSMatrix.Columns.Item("0").Cells.Item(x + 1).Specific).String;
                            for (int y = 1; y <= oMatrix.RowCount; y++)
                            {
                                if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("LineNum").Cells.Item(y).Specific).String == templinenum)
                                {
                                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("0").Cells.Item(y).Specific).String = temprow;
                                }
                            }
                        }
                    }
                }

                oItem = oForm.Items.Item("1");
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                //oForm.Items.Item("TEMP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                //oForm.Items.Item("DS").Visible = false;
                //oForm.Items.Item("LINENUM").Visible = false;
                //oForm.Items.Item("DOCENTRY").Visible = false;
                //oForm.Items.Item("FUID").Visible = false;
                //oForm.Items.Item("st_ROW").Visible = false;
                //oForm.Items.Item("ROW").Visible = false;


                oForm.Visible = true;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized;
                oSForm.Freeze(true);
                oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }
        public static void SDM(string FormUID, long docentry, int linenum, int row, string dsname, string matrixname)
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Initialize popup window...", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
                SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(FormUID);

                SAPbouiCOM.FormCreationParams creationPackage = (SAPbouiCOM.FormCreationParams)FT_ADDON.SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_FormCreationParams);
                creationPackage.UniqueID = "FT_" + (FT_ADDON.SAP.getNewformUID().ToString());
                creationPackage.FormType = "FT_SDM";
                SAPbouiCOM.Form oForm = FT_ADDON.SAP.SBOApplication.Forms.AddEx(creationPackage);
                switch (dsname)
                {
                    case "INV1":
                        oForm.Title = "A/R Invoice UDF Modified";
                        break;
                    case "PDN1":
                        oForm.Title = "Goods Receipt PO UDF Modified";
                        break;
                    case "DLN1":
                        oForm.Title = "Delivery UDF Modified";
                        break;
                    case "RIN1":
                        oForm.Title = "A/R Credit Note UDF Modified";
                        break;
                    case "PCH1":
                        oForm.Title = "A/P Invoice UDF Modified";
                        break;
                    case "RPC1":
                        oForm.Title = "A/P Credit Memo UDF Modified";
                        break;
                    case "POR1":
                        oForm.Title = "Purchase Order UDF Modified";
                        break;
                    case "RDN1":
                        oForm.Title = "Return UDF Modified";
                        break;
                    case "RPD1":
                        oForm.Title = "Goods Return UDF Modified";
                        break;
                    default:
                        oForm.Title = "User Define Fields UDF Modified";
                        break;

                }
                //oForm.Left = oSForm.Left;
                oForm.Width = 700;
                //oForm.Top = oSForm.Top;
                oForm.Height = 500;

                SAPbouiCOM.Item oItem;
                SAPbouiCOM.Button oButton;

                oItem = oForm.Items.Add("1", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 5;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";

                oItem = oForm.Items.Add("2", SAPbouiCOM.BoFormItemTypes.it_BUTTON);
                oItem.Left = 75;
                oItem.Width = 65;
                oItem.Top = oForm.Height - 60;
                oItem.Height = 20;
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "Cancel";

                SAPbouiCOM.UserDataSource uds = null;

                uds = oForm.DataSources.UserDataSources.Add("row", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = row.ToString();

                oItem = oForm.Items.Add("ROW", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                ((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "row");
                oItem.Left = 55;
                oItem.Width = 65;
                oItem.Top = 5;
                oItem.Height = 15;
                oItem.Enabled = false;

                oItem = oForm.Items.Add("st_ROW", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                ((SAPbouiCOM.StaticText)oItem.Specific).Caption = "Row No#";
                oItem.Left = 5;
                oItem.Width = 50;
                oItem.Top = 5;
                oItem.Height = 15;

                uds = oForm.DataSources.UserDataSources.Add("fuid", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = FormUID;

                //oItem = oForm.Items.Add("FUID", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "fuid");
                //oItem.Left = 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("docentry", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = docentry.ToString();

                //oItem = oForm.Items.Add("DOCENTRY", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "docentry");
                //oItem.Left = 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("linenum", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = linenum.ToString();

                //oItem = oForm.Items.Add("LINENUM", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "linenum");
                //oItem.Left = 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;

                uds = oForm.DataSources.UserDataSources.Add("ds", SAPbouiCOM.BoDataType.dt_SHORT_TEXT);
                uds.Value = dsname;

                //oItem = oForm.Items.Add("DS", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                //((SAPbouiCOM.EditText)oItem.Specific).DataBind.SetBound(true, "", "ds");
                //oItem.Left = 5 + 65 + 5 + 65 + 5 + 65 + 5;
                //oItem.Width = 65;
                //oItem.Top = 5;
                //oItem.Height = 15;
                //oItem.Enabled = false;
                //oItem.Visible = false;
 
                SAPbouiCOM.Matrix oSMatrix = (SAPbouiCOM.Matrix)oSForm.Items.Item(matrixname).Specific;

                oItem = oForm.Items.Add("grid1", SAPbouiCOM.BoFormItemTypes.it_MATRIX);
                oItem.Left = 5;
                oItem.Width = oForm.Width - 25;
                oItem.Top = 25;
                oItem.Height = oForm.Height - 90;
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oItem.Specific;
                if (oSMatrix.RowCount > 0)
                {
                    SAPbouiCOM.Conditions oCons = (SAPbouiCOM.Conditions)SAP.SBOApplication.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_Conditions);
                    SAPbouiCOM.DBDataSource oDS = (SAPbouiCOM.DBDataSource)oForm.DataSources.DBDataSources.Add(dsname);

                    oDS.Clear();
                    SAPbouiCOM.Condition oCon = oCons.Add();
                    oCon.BracketOpenNum = 1;
                    oCon.Alias = "DocEntry";
                    oCon.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    oCon.CondVal = docentry.ToString();
                    oCon.BracketCloseNum = 1;

                    SAPbouiCOM.Column oColumn = null;
                    oColumn = oMatrix.Columns.Add("LineNum", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    oColumn.DataBind.SetBound(true, dsname, "LineNum");
                    oColumn.Visible = false;

                    copyUDFMatrixColumns(oSForm, oSMatrix, oForm, oMatrix, dsname, dsname);

                    oDS.Query(oCons);
                    oMatrix.LoadFromDataSource();

                    string templinenum = "";
                    string temprow = "";
                    for (int x = 0; x < oSForm.DataSources.DBDataSources.Item(dsname).Size; x++)
                    {
                        templinenum = oSForm.DataSources.DBDataSources.Item(dsname).GetValue("LineNum", x).ToString();
                        temprow = ((SAPbouiCOM.EditText)oSMatrix.Columns.Item("0").Cells.Item(x + 1).Specific).String;
                        for (int y = 1; y <= oMatrix.RowCount; y++)
                        {
                            if (((SAPbouiCOM.EditText)oMatrix.Columns.Item("LineNum").Cells.Item(y).Specific).String == templinenum)
                            {
                                ((SAPbouiCOM.EditText)oMatrix.Columns.Item("0").Cells.Item(y).Specific).String = temprow;
                            }
                        }
                    }
                }

                oItem = oForm.Items.Item("1");
                oButton = (SAPbouiCOM.Button)oItem.Specific;
                oButton.Caption = "OK";
                //oForm.Items.Item("TEMP").Click(SAPbouiCOM.BoCellClickType.ct_Regular);

                //oForm.Items.Item("DS").Visible = false;
                //oForm.Items.Item("LINENUM").Visible = false;
                //oForm.Items.Item("DOCENTRY").Visible = false;
                //oForm.Items.Item("FUID").Visible = false;
                //oForm.Items.Item("st_ROW").Visible = false;
                //oForm.Items.Item("ROW").Visible = false;


                oForm.Visible = true;
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Minimized;
                oSForm.Freeze(true);
                oMatrix.Columns.Item(1).Cells.Item(1).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("popup window initialize completed!", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
            }
            catch (Exception ex)
            {
                FT_ADDON.SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");

            }
        }
        public static void copyUDFMatrixColumns(SAPbouiCOM.Form oSForm, SAPbouiCOM.Matrix oSMatrix, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, string dsname, string targetdsname)
        {
            string columnname = "";
            try
            {
                int width = 0;
                string temp = "";
                string title = "";
                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;
                SAPbouiCOM.Column oSColumn = null;
                SAPbouiCOM.Column oColumn = null;

                SAPbouiCOM.BoFieldsType fieldType = SAPbouiCOM.BoFieldsType.ft_Text;
                int size = 0;
                SAPbouiCOM.BoDataType datatype = SAPbouiCOM.BoDataType.dt_SHORT_TEXT;
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
                        switch (columnname)
                        {
                            case "0":
                                break;
                            default:
                                if (!columnname.Contains("U_"))
                                    continue;
                                break;
                        }
                        itemtype = oSColumn.Type;//oSMatrix.Columns.Item(col).Type;
                        if (columnname == "0")
                            oColumn = oMatrix.Columns.Add(columnname, itemtype);
                        else
                            oColumn = oMatrix.Columns.Add("U_" + col.ToString(), itemtype);
                        oColumn.TitleObject.Caption = title;
                        oColumn.Width = width;
                        temp = oSColumn.DataBind.Alias;
                        if (temp == null || columnname == "0")
                        {
                            size = 0;
                            datatype = SAPbouiCOM.BoDataType.dt_SHORT_NUMBER;
                            oColumn.DataBind.SetBound(true, targetdsname, "BaseLine");
                        }
                        else
                        {
                            size = oSForm.DataSources.DBDataSources.Item(dsname).Fields.Item(oSColumn.DataBind.Alias).Size;
                            fieldType = oSForm.DataSources.DBDataSources.Item(dsname).Fields.Item(oSColumn.DataBind.Alias).Type;
                            datatype = ObjectFunctions.changeUIFieldsTypeToUIDataType(fieldType);
                            oColumn.DataBind.SetBound(true, targetdsname, columnname);
                        }
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
                            for (int x = 0; x < oSColumn.ValidValues.Count; x++ )
                            {
                                if (oSColumn.ValidValues.Item(x).Description.ToUpper() != "DEFINE NEW")
                                    if (x > 0 && oSColumn.ValidValues.Item(x).Value.Trim() == "")
                                    { }
                                    else
                                        oColumn.ValidValues.Add(oSColumn.ValidValues.Item(x).Value, oSColumn.ValidValues.Item(x).Description);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(columnname + "-" + ex.Message, 1, "Ok", "", "");

            }
        }
        public static void copyUDFMatrixColumnsValues(SAPbouiCOM.Form oSForm, SAPbouiCOM.Matrix oSMatrix, SAPbouiCOM.Form oForm, SAPbouiCOM.Matrix oMatrix, int row)
        {
            try
            {
                string columnname;
                SAPbouiCOM.BoFormItemTypes itemtype = SAPbouiCOM.BoFormItemTypes.it_EDIT;

                string temp = "";
                oMatrix.AddRow(1, oMatrix.RowCount);
                for (int col = 0; col < oMatrix.Columns.Count; col++)
                {
                    if (oMatrix.Columns.Item(col).Visible)
                    {
                        columnname = oMatrix.Columns.Item(col).UniqueID.ToString();
                        switch (columnname)
                        {
                            case "LINENUM":
                                break;
                            default:
                                if (!columnname.Contains("U_"))
                                    continue;
                                break;
                        }
                        itemtype = oMatrix.Columns.Item(columnname).Type;
                        if (itemtype == SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX)
                        {
                            temp = ((SAPbouiCOM.ComboBox)(oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific)).Selected.Value.ToString();
                            ((SAPbouiCOM.ComboBox)(oMatrix.Columns.Item(columnname).Cells.Item(oMatrix.RowCount).Specific)).Select(temp, SAPbouiCOM.BoSearchKey.psk_ByValue);
                        }
                        else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_CHECK_BOX)
                        {
                            if (((SAPbouiCOM.CheckBox)oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific).Checked)
                            {
                                oMatrix.Columns.Item(columnname).Cells.Item(oMatrix.RowCount).Click(SAPbouiCOM.BoCellClickType.ct_Regular, 0);
                            }
                        }
                        else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_LINKED_BUTTON)
                        {
                            temp = ((SAPbouiCOM.EditText)oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific).String;
                            ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnname).Cells.Item(oMatrix.RowCount).Specific)).String = temp;
                        }
                        else if (itemtype == SAPbouiCOM.BoFormItemTypes.it_EDIT)
                        {
                            temp = ((SAPbouiCOM.EditText)oSMatrix.Columns.Item(columnname).Cells.Item(row).Specific).String;
                            ((SAPbouiCOM.EditText)(oMatrix.Columns.Item(columnname).Cells.Item(oMatrix.RowCount).Specific)).String = temp;
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
