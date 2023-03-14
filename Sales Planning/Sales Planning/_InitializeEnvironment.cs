using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class _InitializeEnvironment
    {
        public _InitializeEnvironment()
        {
            // Get an instantialized application object 
            FT_ADDON.SAP.setApplication();

            // Set Connection Context from UIAPI cookie to DIAPI
            if (!(FT_ADDON.SAP.setConnectionContext() == 0))
            {
                FT_ADDON.SAP.SBOApplication.MessageBox("Failed setting a connection to DIAPI", 1, "Ok", "", "");
                System.Environment.Exit(0); //  Terminating the Add-On Application
            }

            try
            {
                FT_ADDON.SAP.SBOCompany = new SAPbobsCOM.Company();
                FT_ADDON.SAP.SBOCompany = (SAPbobsCOM.Company)FT_ADDON.SAP.SBOApplication.Company.GetDICompany();

            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Failed to connect to company : " + ex.Message, 1, "Ok", "", "");
                System.Environment.Exit(0);
                return;
            }

            // Connect to SBO Database through DIAPI
            //if (!(FT_ADDON.SAP.connectToCompany() == 0))
            //{
            //    FT_ADDON.SAP.SBOApplication.MessageBox(FT_ADDON.SAP.SBOCompany.GetLastErrorCode().ToString() + ": "+FT_ADDON.SAP.SBOCompany.GetLastErrorDescription(),1,"OK","","" );
            //    System.Environment.Exit(0); //  Terminating the Add-On Application
            //}

            // Display status
            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Sales Planning Addon Initializing...", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

            // Add deligates to events
            FT_ADDON.SAP.SBOApplication.AppEvent += new SAPbouiCOM._IApplicationEvents_AppEventEventHandler(SBO_Application_AppEvent);
            FT_ADDON.SAP.SBOApplication.MenuEvent += new SAPbouiCOM._IApplicationEvents_MenuEventEventHandler(SBO_Application_MenuEvent);
            FT_ADDON.SAP.SBOApplication.ItemEvent += new SAPbouiCOM._IApplicationEvents_ItemEventEventHandler(SBO_Application_ItemEvent);
            FT_ADDON.SAP.SBOApplication.ProgressBarEvent += new SAPbouiCOM._IApplicationEvents_ProgressBarEventEventHandler(SBO_Application_ProgressBarEvent);
            FT_ADDON.SAP.SBOApplication.RightClickEvent += new SAPbouiCOM._IApplicationEvents_RightClickEventEventHandler(SBO_Application_RightClickEvent);
            FT_ADDON.SAP.SBOApplication.FormDataEvent += new SAPbouiCOM._IApplicationEvents_FormDataEventEventHandler(ref SBO_Application_FormDataEvent);

            // Add UDT, UDF, Menu Item
            FT_ADDON.SAP.formUID = 0;
            FT_ADDON.SAP.createStatusForm();
            FT_ADDON.SAP.getStatusForm();
            initEnviroment();

            // Display status
            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Sales Planning Addon successfully initialized.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
        }

        private void SBO_Application_RightClickEvent(ref SAPbouiCOM.ContextMenuInfo EventInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            ItemEvent.processRightClickEvent(EventInfo.FormUID, ref EventInfo, ref BubbleEvent);
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            ItemEvent.processItemEvent(FormUID, ref pVal, ref BubbleEvent);
        }

        private void SBO_Application_AppEvent(SAPbouiCOM.BoAppEventTypes EventType)
        {
            switch (EventType)
            {
                case SAPbouiCOM.BoAppEventTypes.aet_ShutDown:
                    //SBO_Application.MessageBox("A Shut Down Event has been caught" + Environment.NewLine + "Terminating Add On...", 1, "Ok", "", "");
                    System.Environment.Exit(0);
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_CompanyChanged:
                    //SBO_Application.MessageBox("A Company Change Event has been caught", 1, "Ok", "", "");
                    System.Environment.Exit(0);
                    break;
                case SAPbouiCOM.BoAppEventTypes.aet_LanguageChanged:
                    //SBO_Application.MessageBox("A Languge Change Event has been caught", 1, "Ok", "", "");
                    break;
            }
        }

        private void SBO_Application_ProgressBarEvent(ref SAPbouiCOM.ProgressBarEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        private void SBO_Application_StatusBarEvent(string Text, SAPbouiCOM.BoStatusBarMessageType MessageType)
        {
            //SBO_Application.MessageBox(@"Status bar event with message: """ + Text + @""" has been sent", 1, "Ok", "", "");
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
            ItemEvent.processDataEvent(ref BusinessObjectInfo, ref BubbleEvent);
        }

        private void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            if (!pVal.BeforeAction) MenuEvent.processMenuEvent(ref pVal);

            MenuEvent.processMenuEvent2(ref pVal, ref BubbleEvent);
        }

        private void initEnviroment()
        {
            // -------------------------------------------------------
            // Add UDT, UDF, Add Menu Item
            // -------------------------------------------------------

            // Column Name cannot more than 10 Charateres
            string errmsg = "";
            FT_ADDON.ApplicationCommon app = new ApplicationCommon();
            FT_ADDON.SAP.SBOCompany.StartTransaction();
            try
            {
                #region emailtable
                if (!app.createTable("FT_SEMAIL", "Sender Email", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_SEMAIL"))
                {
                    if (!app.createField("@FT_SEMAIL", "objcode", "ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_SEMAIL", "emailaddr", "Email Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_SEMAIL", "emailname", "Email Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_SEMAIL", "host", "Email Host", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_SEMAIL", "port", "Email Port", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "0", true)) goto ErrorHandler;
                    if (!app.createField("@FT_SEMAIL", "emailuid", "Email User", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_SEMAIL", "emailpwd", "Email Password", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_SEMAIL", "enablessl"))
                    if (!app.createField("@FT_SEMAIL", "enablessl", "Enable SSL", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", true, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;

                if (!app.createTable("FT_REMAIL", "Receiver Email", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_REMAIL"))
                {
                    if (!app.createField("@FT_REMAIL", "objcode", "ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_REMAIL", "emailaddr", "Email Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_REMAIL", "emailname", "Email Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", true)) goto ErrorHandler;
                }
                #endregion
                #region sapinternalmessage
                if (!app.createTable("FT_RSYSMSG", "Receiver of System Message", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_RSYSMSG"))
                {
                    if (!app.createField("@FT_RSYSMSG", "objcode", "ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_RSYSMSG", "userid", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_RSYSMSG", "trxtype", "Transaction", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "U", true, SAPbobsCOM.BoFldSubTypes.st_None, "A:Add|U:Update|B:Both", "")) goto ErrorHandler;
                }
                if (!app.createTable("FT_RAPPMSG", "Receiver of Approval Message", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_RAPPMSG"))
                {
                    if (!app.createField("@FT_RAPPMSG", "objcode", "ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_RAPPMSG", "userid", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_RAPPMSG", "trxtype", "Transaction", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "U", true, SAPbobsCOM.BoFldSubTypes.st_None, "A:Add|U:Update|B:Both", "")) goto ErrorHandler;
                    if (!app.createField("@FT_RAPPMSG", "isapp", "Is Approve Stage", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", true, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                }
                #endregion
                #region sales planning
                if (!app.createTable("FT_SPLAN", "Sales Planning", SAPbobsCOM.BoUTBTableType.bott_Document)) goto ErrorHandler;
                if (!app.tableGotField("@FT_SPLAN"))
                {
                    if (!app.createField("@FT_SPLAN", "DOCDATE", "Sales Planning Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    //if (!app.createField("@FT_SPLAN", "DODATE", "DO Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN", "REFNO", "Ref No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;

                }
                if (!app.udfExist("@FT_SPLAN", "CARDCODE"))
                    if (!app.createField("@FT_SPLAN", "CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "CARDNAME"))
                    if (!app.createField("@FT_SPLAN", "CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;

                if (!app.udfExist("@FT_SPLAN", "SODOCNUM"))
                    if (!app.createField("@FT_SPLAN", "SODOCNUM", "SO Doc No", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "COMSO"))
                    if (!app.createField("@FT_SPLAN", "COMSO", "Combine SO", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;

                if (!app.udfExist("@FT_SPLAN", "REMARKS"))
                    if (!app.createField("@FT_SPLAN", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "FREMARKS"))
                    if (!app.createField("@FT_SPLAN", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "RELEASE"))
                    if (!app.createField("@FT_SPLAN", "RELEASE", "Released", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;


                if (!app.udfExist("@FT_SPLAN", "APP"))
                    if (!app.createField("@FT_SPLAN", "APP", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "APPRE"))
                    if (!app.createField("@FT_SPLAN", "APPRE", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "APPBY"))
                    if (!app.createField("@FT_SPLAN", "APPBY", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "APPDATE"))
                    if (!app.createField("@FT_SPLAN", "APPDATE", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN", "APPTIME"))
                    if (!app.createField("@FT_SPLAN", "APPTIME", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "")) goto ErrorHandler;

                if (!app.createTable("FT_SPLAN1", "Sales Planning Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_SPLAN1"))
                {
                    if (!app.createField("@FT_SPLAN1", "CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SODOCNUM", "SO Doc No.", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SODATE", "SO Doc Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SOITEMCO", "SO Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SOITEMNA", "SO Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SOWHSCOD", "SO Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "ITEMCODE", "New Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "ITEMNAME", "New Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "WHSCODE", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "ADDRESS", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                    // no use anymore
                    //if (!app.createField("@FT_SPLAN1", "BIN", "Bin", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    // Ori SO Quantity
                    if (!app.createField("@FT_SPLAN1", "ORIQTY", "Original SO QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    // Reserve Existing Stock Quantity
                    if (!app.createField("@FT_SPLAN1", "QUANTITY", "Reserve Existing Stock QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "TPQTY", "Transport Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SOENTRY", "SO Docentry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "SOLINE", "SO Linenum", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "BASEOBJ", "Base Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "BASEENT", "Base Docentry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_SPLAN1", "BASELINE", "Base Linenum", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_SPLAN1", "SODOCQTY"))
                    if (!app.createField("@FT_SPLAN1", "SODOCQTY", "SO Document QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "ADDRESS"))
                    if (!app.createField("@FT_SPLAN1", "ADDRESS", "ADDRESS", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("@FT_SPLAN1", "SOQTY"))
                    if (!app.createField("@FT_SPLAN1", "SOQTY", "Open SO QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "SOPQTY"))
                    if (!app.createField("@FT_SPLAN1", "SOPQTY", "Open SO Planning QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;

                if (!app.udfExist("@FT_SPLAN1", "CUSTPO"))
                    if (!app.createField("@FT_SPLAN1", "CUSTPO", "Customer PO No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "VSNAME"))
                    if (!app.createField("@FT_SPLAN1", "VSNAME", "Vessel Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;

                // Reserved Blanket PO
                if (!app.udfExist("@FT_SPLAN1", "RBPONO"))
                    if (!app.createField("@FT_SPLAN1", "RBPONO", "Reserved Blanket PO", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "0")) goto ErrorHandler;
                // Reserved Blanket PO QTY
                if (!app.udfExist("@FT_SPLAN1", "RBPOQTY"))
                    if (!app.createField("@FT_SPLAN1", "RBPOQTY", "Reserved Blanket PO QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "RBPOOPEN"))
                    if (!app.createField("@FT_SPLAN1", "RBPOOPEN", "Reserved Blanket Open QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;

                if (!app.udfExist("@FT_SPLAN1", "ALLITEM"))
                    if (!app.createField("@FT_SPLAN1", "ALLITEM", "All Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "UNITMSR"))
                    if (!app.createField("@FT_SPLAN1", "UNITMSR", "UOM Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "LREMARKS"))
                    if (!app.createField("@FT_SPLAN1", "LREMARKS", "Internal Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_SPLAN1", "EREMARKS"))
                    if (!app.createField("@FT_SPLAN1", "EREMARKS", "External Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("@FT_SPLAN1", "LSTATUS"))
                    if (!app.createField("@FT_SPLAN1", "LSTATUS", "Line Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false)) goto ErrorHandler;

                //if (!app.createField("@FT_SPLAN1", "RPONO", "Reserved PO No.", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                //if (!app.createField("@FT_SPLAN1", "RPOQTY", "Reserved Stock Qty", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;

                #endregion
                #region transport planning
                if (!app.createTable("FT_TPPLAN", "Transport Planning", SAPbobsCOM.BoUTBTableType.bott_Document)) goto ErrorHandler;
                if (!app.tableGotField("@FT_TPPLAN"))
                {
                    if (!app.createField("@FT_TPPLAN", "CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "DOCDATE", "TP Planning Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "DODATE", "DO Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "TPCODE", "Transporter", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "LORRY", "Lorry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "DRIVER", "Driver", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "DRIVERIC", "Driver IC", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "AREA", "Area Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN", "BILLCUST", "Bill to Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;

                }
                if (!app.udfExist("@FT_TPPLAN", "DRIVERNA"))
                    if (!app.createField("@FT_TPPLAN", "DRIVERNA", "Driver Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "REFNO"))
                    if (!app.createField("@FT_TPPLAN", "REFNO", "Ref No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "APRICE"))
                    if (!app.createField("@FT_TPPLAN", "APRICE", "Area Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "DOCTOTAL"))
                    if (!app.createField("@FT_TPPLAN", "DOCTOTAL", "Document Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "WHSCODE"))
                    if (!app.createField("@FT_TPPLAN", "WHSCODE", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "REMARKS"))
                    if (!app.createField("@FT_TPPLAN", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "FREMARKS"))
                    if (!app.createField("@FT_TPPLAN", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "ADDRESS"))
                    if (!app.createField("@FT_TPPLAN", "ADDRESS", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "REAPRICE"))
                    if (!app.createField("@FT_TPPLAN", "REAPRICE", "Require Price", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "Y", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;

                if (!app.udfExist("@FT_TPPLAN", "APP"))
                    if (!app.createField("@FT_TPPLAN", "APP", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "APPRE"))
                    if (!app.createField("@FT_TPPLAN", "APPRE", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "APPBY"))
                    if (!app.createField("@FT_TPPLAN", "APPBY", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "APPDATE"))
                    if (!app.createField("@FT_TPPLAN", "APPDATE", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "APPTIME"))
                    if (!app.createField("@FT_TPPLAN", "APPTIME", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "RELEASE"))
                    if (!app.createField("@FT_TPPLAN", "RELEASE", "Released", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN", "SPNO"))
                    if (!app.createField("@FT_TPPLAN", "SPNO", "SP Number", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;

                if (!app.createTable("FT_TPPLAN1", "Transport Planning Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_TPPLAN1"))
                {
                    if (!app.createField("@FT_TPPLAN1", "CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SODOCNUM", "SO Doc No.", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SODATE", "SO Doc Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SOITEMCO", "SO Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SOITEMNA", "SO Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SOWHSCOD", "SO Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "ITEMCODE", "New Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "ITEMNAME", "New Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "WHSCODE", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    //if (!app.createField("@FT_TPPLAN1", "BIN", "Bin", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SOQTY", "SO Open Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "ORIQTY", "Original Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "QUANTITY", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "CMQTY", "Charge Module Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SOENTRY", "SO Docentry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "SOLINE", "SO Linenum", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "BASEOBJ", "Base Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "BASEENT", "Base Docentry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "BASELINE", "Base Linenum", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "RELEASE", "Release to DO (Y/N)", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "Y", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                    if (!app.createField("@FT_TPPLAN1", "ADDRESS", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                }

                if (!app.udfExist("@FT_TPPLAN1", "BASEOBJ"))
                    if (!app.createField("@FT_TPPLAN1", "BASEOBJ", "Base Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "SOBASEOB"))
                    if (!app.createField("@FT_TPPLAN1", "SOBASEOB", "SO Base Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "WEIGHT"))
                    if (!app.createField("@FT_TPPLAN1", "WEIGHT", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "LTOTAL"))
                    if (!app.createField("@FT_TPPLAN1", "LTOTAL", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) goto ErrorHandler;

                if (!app.udfExist("@FT_TPPLAN1", "CUSTPO"))
                    if (!app.createField("@FT_TPPLAN1", "CUSTPO", "Customer PO No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "VSNAME"))
                    if (!app.createField("@FT_TPPLAN1", "VSNAME", "Vessel Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "ADDRESS"))
                    if (!app.createField("@FT_TPPLAN1", "ADDRESS", "ADDRESS", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("@FT_TPPLAN1", "UNITMSR"))
                    if (!app.createField("@FT_TPPLAN1", "UNITMSR", "UOM Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "SPQTY"))
                    if (!app.createField("@FT_TPPLAN1", "SPQTY", "Sales Plan QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "SPDOCNUM"))
                    if (!app.createField("@FT_TPPLAN1", "SPDOCNUM", "Sales Plan No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "LREMARKS"))
                    if (!app.createField("@FT_TPPLAN1", "LREMARKS", "Internal Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "EREMARKS"))
                    if (!app.createField("@FT_TPPLAN1", "EREMARKS", "External Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                //if (!app.udfExist("@FT_TPPLAN1", "REMARKS"))
                //    if (!app.createField("@FT_TPPLAN1", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                //if (!app.udfExist("@FT_TPPLAN1", "FREMARKS"))
                //    if (!app.createField("@FT_TPPLAN1", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_TPPLAN1", "LSTATUS"))
                    if (!app.createField("@FT_TPPLAN1", "LSTATUS", "Line Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false)) goto ErrorHandler;

                #endregion
                #region charge module
                if (!app.createTable("FT_CHARGE", "Charge Module", SAPbobsCOM.BoUTBTableType.bott_Document)) goto ErrorHandler;
                if (!app.tableGotField("@FT_CHARGE"))
                {
                    if (!app.createField("@FT_CHARGE", "CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "DOCDATE", "CM Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "DODATE", "DO Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "TPDOCNUM", "Transport Loading No", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "TPCODE", "Transporter", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "LORRY", "Lorry No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "DRIVER", "Driver", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "DRIVERIC", "Driver IC", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "AREA", "Area Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE", "BILLCUST", "Bill to Customer", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_CHARGE", "DRIVERNA"))
                    if (!app.createField("@FT_CHARGE", "DRIVERNA", "Driver Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "REFNO"))
                    if (!app.createField("@FT_CHARGE", "REFNO", "Ref No.", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "APRICE"))
                    if (!app.createField("@FT_CHARGE", "APRICE", "Area Price", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "DOCTOTAL"))
                    if (!app.createField("@FT_CHARGE", "DOCTOTAL", "Document Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;

                if (!app.udfExist("@FT_CHARGE", "REMARKS"))
                    if (!app.createField("@FT_CHARGE", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "FREMARKS"))
                    if (!app.createField("@FT_CHARGE", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "ADDRESS"))
                    if (!app.createField("@FT_CHARGE", "ADDRESS", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("@FT_CHARGE", "APP"))
                    if (!app.createField("@FT_CHARGE", "APP", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "APPRE"))
                    if (!app.createField("@FT_CHARGE", "APPRE", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "APPBY"))
                    if (!app.createField("@FT_CHARGE", "APPBY", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "APPDATE"))
                    if (!app.createField("@FT_CHARGE", "APPDATE", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE", "APPTIME"))
                    if (!app.createField("@FT_CHARGE", "APPTIME", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "")) goto ErrorHandler;

                if (!app.createTable("FT_CHARGE1", "Charge Module Detail", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_CHARGE1"))
                {
                    if (!app.createField("@FT_CHARGE1", "CARDCODE", "Customer Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "CARDNAME", "Customer Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SODOCNUM", "SO Doc No.", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SODATE", "SO Doc Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SOITEMCO", "SO Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SOITEMNA", "SO Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SOWHSCOD", "SO Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "ITEMCODE", "New Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "ITEMNAME", "New Item Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "UOM", "UOM", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "WHSCODE", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    //if (!app.createField("@FT_CHARGE1", "BIN", "Bin", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SOQTY", "SO Open Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "ORIQTY", "Original Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "QUANTITY", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "DOQTY", "Charge Module Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SOENTRY", "SO Docentry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "SOLINE", "SO Linenum", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "BASEOBJ", "Base Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "BASEENT", "Base Docentry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "BASELINE", "Base Linenum", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE1", "ADDRESS", "Address", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_CHARGE1", "SOBASEOB"))
                    if (!app.createField("@FT_CHARGE1", "SOBASEOB", "SO Base Object", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "WEIGHT"))
                    if (!app.createField("@FT_CHARGE1", "WEIGHT", "Weight", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "LTOTAL"))
                    if (!app.createField("@FT_CHARGE1", "LTOTAL", "Line Total", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Measurement)) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "BQTY"))
                    if (!app.createField("@FT_CHARGE1", "BQTY", "Batch QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;

                if (!app.udfExist("@FT_CHARGE1", "CUSTPO"))
                    if (!app.createField("@FT_CHARGE1", "CUSTPO", "Customer PO No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "VSNAME"))
                    if (!app.createField("@FT_CHARGE1", "VSNAME", "Vessel Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "ADDRESS"))
                    if (!app.createField("@FT_CHARGE1", "ADDRESS", "ADDRESS", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("@FT_CHARGE1", "ALLITEM"))
                    if (!app.createField("@FT_CHARGE1", "ALLITEM", "All Item", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "UNITMSR"))
                    if (!app.createField("@FT_CHARGE1", "UNITMSR", "UOM Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "TPQTY"))
                    if (!app.createField("@FT_CHARGE1", "TPQTY", "Transport Plan QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "LREMARKS"))
                    if (!app.createField("@FT_CHARGE1", "LREMARKS", "Internal Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "EREMARKS"))
                    if (!app.createField("@FT_CHARGE1", "EREMARKS", "External Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                //if (!app.udfExist("@FT_CHARGE1", "REMARKS"))
                //    if (!app.createField("@FT_CHARGE1", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                //if (!app.udfExist("@FT_CHARGE1", "FREMARKS"))
                //    if (!app.createField("@FT_CHARGE1", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE1", "LSTATUS"))
                    if (!app.createField("@FT_CHARGE1", "LSTATUS", "Line Status", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false)) goto ErrorHandler;

                if (!app.createTable("FT_CHARGE2", "Charge Module Batch", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_CHARGE2"))
                {
                    if (!app.createField("@FT_CHARGE2", "ITEMCODE", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE2", "WHSCODE", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE2", "BATCHNUM", "Batch No", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE2", "QUANTITY", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_CHARGE2", "BASEVIS", "Base VisOrder No", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "0")) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_CHARGE2", "BIN"))
                    if (!app.createField("@FT_CHARGE2", "BIN", "Bin", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CHARGE2", "BINABS"))
                    if (!app.createField("@FT_CHARGE2", "BINABS", "Bin Entry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                #endregion

                if (!app.createTable("FT_SPERRLOG", "SP Error Log", SAPbobsCOM.BoUTBTableType.bott_Document)) goto ErrorHandler;
                if (!app.tableGotField("@FT_SPERRLOG"))
                {
                    if (!app.createField("@FT_SPERRLOG", "ErrDT", "Error DT", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "ErrMsg", "Error Msg", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "RefDate", "Ref Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "Memo", "Memo", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "DocType", "Doc Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "RInvNo", "U AR RIV", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "ARInvNo", "U AR IV", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "DelNo", "U AR DO", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG", "APINVNo", "U AP IV", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                }
                if (!app.createTable("FT_SPERRLOG1", "SP Error Log", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_SPERRLOG1"))
                {
                    if (!app.createField("@FT_SPERRLOG1", "AccountCode", "Account", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG1", "CostingCode", "Costing Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG1", "Debit", "Debit", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;
                    if (!app.createField("@FT_SPERRLOG1", "Credit", "Credit", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;
                }
                #region APDO Details
                if (!app.createTable("FT_APDOSP", "GRN Sales Planning Doc", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_APDOSP"))
                {
                    //if (!app.createField("@FT_APSOC", "DOCNO", "DOCENTRY", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "0")) goto ErrorHandler;
                    //if (!app.createField("@FT_APSOC", "LINENO", "LINENUM", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_APDOSP", "LINENO", "Actual LINENUM", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_APDOSP", "SPDOC", "Sales Planning Doc", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "0")) goto ErrorHandler;
                    if (!app.createField("@FT_APDOSP", "SPDOCQTY", "Sales Planning QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    //if (!app.createField("@FT_APSOC", "REF", "Reference", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "")) goto ErrorHandler;
                }

                #endregion
                if (!app.createTable("FT_BATCH", "Batch No", SAPbobsCOM.BoUTBTableType.bott_MasterData)) goto ErrorHandler;
                if (!app.tableGotField("@FT_BATCH"))
                {
                    if (!app.createField("@FT_BATCH", "ITEMCODE", "Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_BATCH", "WHSCODE", "Warehouse", SAPbobsCOM.BoFieldTypes.db_Alpha, 30, "")) goto ErrorHandler;
                    if (!app.createField("@FT_BATCH", "BATCHNUM", "Batch No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_BATCH", "ONHAND", "On Hand", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                    if (!app.createField("@FT_BATCH", "QUANTITY", "Quantity", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_BATCH", "BIN"))
                    if (!app.createField("@FT_BATCH", "BIN", "Bin", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_BATCH", "BINABS"))
                    if (!app.createField("@FT_BATCH", "BINABS", "Bin Entry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;

                if (!app.udfExist("@TRANSPORTER_AREA_D", "EXPIRED"))
                    if (!app.createField("@TRANSPORTER_AREA_D", "EXPIRED", "Expired", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;

                if (!app.udfExist("ODLN", "SPLAN"))
                    if (!app.createField("ODLN", "SPLAN", "Sales Planning", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "Y", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;

                if (!app.udfExist("DLN1", "SPLANQTY"))
                    if (!app.createField("DLN1", "SPLANQTY", "Sales Planning QTY", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Quantity)) goto ErrorHandler;

                if (!app.udfExist("DLN1", "CMEntry"))
                    if (!app.createField("DLN1", "CMEntry", "CM DocEntry", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0", false)) goto ErrorHandler;

                if (!app.udfExist("DLN1", "CMLine"))
                    if (!app.createField("DLN1", "CMLine", "CM LineId", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0", false)) goto ErrorHandler;

                if (!app.udfExist("DLN1", "RevisedItemCode"))
                    if (!app.createField("DLN1", "RevisedItemCode", " Revised Item Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("DLN1", "RevisedItemDesc"))
                    if (!app.createField("DLN1", "RevisedItemDesc", "Revised Item Desc", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "", false)) goto ErrorHandler;

                if (!app.udfExist("DLN1", "SPDOC"))
                    if (!app.createField("DLN1", "SPDOC", "Sales Planning Doc", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "", false)) goto ErrorHandler;
                //if (!app.udfExist("DLN1", "LREMARKS")) 
                //    if (!app.createField("DLN1", "LREMARKS", "Internal Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("DLN1", "EREMARKS"))
                    if (!app.createField("DLN1", "EREMARKS", "External Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                //if (!app.udfExist("DLN1", "REMARKS"))
                //    if (!app.createField("DLN1", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                //if (!app.udfExist("DLN1", "FREMARKS"))
                //    if (!app.createField("DLN1", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("ODLN", "ChargeNo"))
                    if (!app.createField("ODLN", "ChargeNo", "Charge Module No", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", false)) goto ErrorHandler;
                if (!app.udfExist("ODLN", "Diff"))
                    if (!app.createField("ODLN", "Diff", "Difference", SAPbobsCOM.BoFieldTypes.db_Float, 0, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;

                if (!app.udfExist("ODLN", "REMARKS"))
                    if (!app.createField("ODLN", "REMARKS", "Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                if (!app.udfExist("ODLN", "FREMARKS"))
                    if (!app.createField("ODLN", "FREMARKS", "Full Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;

                if (!app.udfExist("ODLN", "ADDONIND"))
                    if (!app.createField("ODLN", "ADDONIND", "Add on Indicator", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N")) goto ErrorHandler;

                if (!app.udfExist("OJDT", "RInvNo"))
                    if (!app.createField("OJDT", "RInvNo", "Reserved Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")) goto ErrorHandler;

                if (!app.udfExist("OJDT", "DelNo"))
                    if (!app.createField("OJDT", "DelNo", "Delivery No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")) goto ErrorHandler;

                if (!app.udfExist("OJDT", "ARInvNo"))
                    if (!app.createField("OJDT", "ARInvNo", "A/R Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")) goto ErrorHandler;

                if (!app.udfExist("OJDT", "APINVNo"))
                    if (!app.createField("OJDT", "APINVNo", "A/P Invoice No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")) goto ErrorHandler;

                if (!app.udfExist("OIGE", "DelNo"))
                    if (!app.createField("OIGE", "DelNo", "Delivery No", SAPbobsCOM.BoFieldTypes.db_Alpha, 10, "")) goto ErrorHandler;

                if (!app.udfExist("OPOR", "PRebPer"))
                    if (!app.createField("OPOR", "PRebPer", "Purchase Rebate (%)", SAPbobsCOM.BoFieldTypes.db_Float, 10, "0", false, SAPbobsCOM.BoFldSubTypes.st_Percentage)) goto ErrorHandler;

                if (!app.udfExist("OPOR", "PRebAmt"))
                    if (!app.createField("OPOR", "PRebAmt", "Purchase Rebate (RM)", SAPbobsCOM.BoFieldTypes.db_Float, 10, "0", false, SAPbobsCOM.BoFldSubTypes.st_Percentage)) goto ErrorHandler;

                if (!app.udfExist("OADM", "MDBName"))
                    if (!app.createField("OADM", "MDBName", "Master Database Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "")) goto ErrorHandler;

                if (!app.udfExist("OCRD", "Grace"))
                    if (!app.createField("OCRD", "Grace", "Grace", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;

                //Payment Terms
                if (!app.udfExist("OCTG", "Grace"))
                    if (!app.createField("OCTG", "Grace", "Grace", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;
                if (!app.udfExist("OCTG", "Desc"))
                    if (!app.createField("OCTG", "Desc", "Description", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;


                if (!app.udfExist("OQUT", "Grace"))
                    if (!app.createField("OQUT", "Grace", "Grace", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;


                if (!app.udfExist("ORDR", "Grace"))
                    if (!app.createField("ORDR", "Grace", "Grace", SAPbobsCOM.BoFieldTypes.db_Numeric, 11, "0")) goto ErrorHandler;


                if (!app.udfExist("ORDR", "CUsage"))
                    if (!app.createField("ORDR", "CUsage", "Current Usage", SAPbobsCOM.BoFieldTypes.db_Float, 10, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;

                if (!app.udfExist("ORDR", "TLimit"))
                    if (!app.createField("ORDR", "TLimit", "Temporary Limit", SAPbobsCOM.BoFieldTypes.db_Float, 10, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;
                if (!app.udfExist("ORDR", "CLimit"))
                    if (!app.createField("ORDR", "CLimit", "Customer Limit", SAPbobsCOM.BoFieldTypes.db_Float, 10, "0", false, SAPbobsCOM.BoFldSubTypes.st_Sum)) goto ErrorHandler;

                if (!app.udfExist("ORDR", "APP"))
                    if (!app.createField("ORDR", "APP", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;
                if (!app.udfExist("ORDR", "APPRE"))
                    if (!app.createField("ORDR", "APPRE", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 100, "")) goto ErrorHandler;
                if (!app.udfExist("ORDR", "APPBY"))
                    if (!app.createField("ORDR", "APPBY", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                if (!app.udfExist("ORDR", "APPDATE"))
                    if (!app.createField("ORDR", "APPDATE", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                if (!app.udfExist("ORDR", "APPTIME"))
                    if (!app.createField("ORDR", "APPTIME", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "")) goto ErrorHandler;
                if (!app.udfExist("ORDR", "DSAPP"))
                    if (!app.createField("ORDR", "DSAPP", "Drop Ship Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;

                if (!app.udfExist("ORDR", "CTERM"))
                    if (!app.createField("ORDR", "CTERM", "Credit Term", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N")) goto ErrorHandler;

                if (!app.udfExist("INV1", "InvCost"))
                    if (!app.createField("INV1", "InvCost", "Invoice Cost", SAPbobsCOM.BoFieldTypes.db_Float, 10, "0", false, SAPbobsCOM.BoFldSubTypes.st_Price)) goto ErrorHandler;

                #region Custom Form Setting table
                if (!app.createTable("FT_CFS", "Custom Form Setting", SAPbobsCOM.BoUTBTableType.bott_MasterData)) goto ErrorHandler;
                if (!app.tableGotField("@FT_CFS"))
                {
                    if (!app.createField("@FT_CFS", "FNAME", "Form Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CFS", "USRID", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "")) goto ErrorHandler;
                }
                if (!app.udfExist("@FT_CFS", "MATRIX"))
                    if (!app.createField("@FT_CFS", "MATRIX", "Matrix Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "")) goto ErrorHandler;
                if (!app.udfExist("@FT_CFS", "DSNAME"))
                    if (!app.createField("@FT_CFS", "DSNAME", "Table Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "")) goto ErrorHandler;

                if (!app.createTable("FT_CFSDL", "Custom Form Setting Detail", SAPbobsCOM.BoUTBTableType.bott_MasterDataLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_CFSDL"))
                {
                    if (!app.createField("@FT_CFSDL", "CNAME", "Column Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 20, "")) goto ErrorHandler;
                    if (!app.createField("@FT_CFSDL", "NONVIEW", "Cannot View", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "1")) goto ErrorHandler;
                    if (!app.createField("@FT_CFSDL", "NONEDIT", "Cannot Edit", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "1")) goto ErrorHandler;
                }


                if (!app.createTable("FT_SPCFSQL", "Sales Planning Copy From SQL", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_SPCFSQL"))
                {
                    if (!app.createField("@FT_SPCFSQL", "UDO", "UDO", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPCFSQL", "Header", "Header", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "N", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Yes|N:No", "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPCFSQL", "HColumn", "Header Column Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPCFSQL", "Btn", "Button Code", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPCFSQL", "BtnName", "Button Name", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_SPCFSQL", "BtnSQL", "Button Copy from SQL", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "")) goto ErrorHandler;
                }
                if (!app.createTable("FT_APPSQL", "Approval SQL", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_APPSQL"))
                {
                    if (!app.createField("@FT_APPSQL", "AppSQL", "Approval SQL", SAPbobsCOM.BoFieldTypes.db_Memo, 0, "")) goto ErrorHandler;
                }

                if (!app.createTable("FT_APPUSER", "Custom Approval User", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@FT_APPUSER"))
                {
                    if (!app.createField("@FT_APPUSER", "objcode", "ObjType", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                    if (!app.createField("@FT_APPUSER", "userid", "User ID", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "", true)) goto ErrorHandler;
                }

                if (!app.createTable("FT_APPTPLOG", "TP Approval Log", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_APPTPLOG"))
                {
                    if (!app.createField("@FT_APPTPLOG", "APP", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPTPLOG", "APPRE", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPTPLOG", "APPBY", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPTPLOG", "APPDATE", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPTPLOG", "APPTIME", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPTPLOG", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                }
                if (!app.createTable("FT_APPSOLOG", "SO Approval Log", SAPbobsCOM.BoUTBTableType.bott_DocumentLines)) goto ErrorHandler;
                if (!app.tableGotField("@FT_APPSOLOG"))
                {
                    if (!app.createField("@FT_APPSOLOG", "APP", "Approval", SAPbobsCOM.BoFieldTypes.db_Alpha, 1, "O", false, SAPbobsCOM.BoFldSubTypes.st_None, "Y:Approved|W:Pending|N:Reject|O:OPEN", "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPSOLOG", "APPRE", "Approval Remarks", SAPbobsCOM.BoFieldTypes.db_Alpha, 254, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPSOLOG", "APPBY", "Approved by", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPSOLOG", "APPDATE", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Date, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPSOLOG", "APPTIME", "Approved Date", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "")) goto ErrorHandler;
                    if (!app.createField("@FT_APPSOLOG", "ObjType", "Object Type", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                }

                #endregion
                if (!app.createTable("COUNTRY", "COUNTRY", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.createTable("REGION", "Region", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.createTable("AREA", "Area", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@AREA"))
                {
                    if (!app.createField("@AREA", "REGION", "Region", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                    if (!app.createField("@AREA", "COUNTRY", "Country", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "")) goto ErrorHandler;
                }

                //if (!app.createUDO("FT_SPLAN", "Sales Planning", SAPbobsCOM.BoUDOObjType.boud_Document, "FT_SPLAN", "FT_SPLAN1", "U_DOCDATE|U_DODATE|U_REFNO|Canceled", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, "AFT_SPLAN")) goto ErrorHandler; ;
                if (!app.createUDO("FT_SPLAN", "Sales Planning", SAPbobsCOM.BoUDOObjType.boud_Document, "FT_SPLAN", "FT_SPLAN1", "U_DOCDATE|U_REFNO|Canceled", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, "AFT_SPLAN")) goto ErrorHandler; ;
                if (!app.createUDO("FT_TPPLAN", "Transport Planning", SAPbobsCOM.BoUDOObjType.boud_Document, "FT_TPPLAN", "FT_TPPLAN1", "U_DOCDATE|U_DODATE|U_REFNO|U_CARDCODE|Canceled", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, "AFT_TPPLAN")) goto ErrorHandler; ;
                if (!app.createUDO("FT_CHARGE", "Charge Module", SAPbobsCOM.BoUDOObjType.boud_Document, "FT_CHARGE", "FT_CHARGE1|FT_CHARGE2", "U_DOCDATE|U_DODATE|U_REFNO|U_CARDCODE|Canceled", SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, "AFT_CHARGE")) goto ErrorHandler; ;

                if (!app.createUDO("FT_SPERRLOG", "SP Error Log", SAPbobsCOM.BoUDOObjType.boud_Document, "FT_SPERRLOG", "FT_SPERRLOG1", "U_ErrDT", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, "")) goto ErrorHandler; ;

                if (!app.udfExist("CUFD", "seq"))
                    if (!app.createField("CUFD", "seq", "Seq No", SAPbobsCOM.BoFieldTypes.db_Numeric, 0, "0")) goto ErrorHandler;

                if (!app.createUDO("FT_CFS", "Custom Form Setting", SAPbobsCOM.BoUDOObjType.boud_MasterData, "FT_CFS", "FT_CFSDL", "", SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tNO, SAPbobsCOM.BoYesNoEnum.tYES, SAPbobsCOM.BoYesNoEnum.tNO, "")) goto ErrorHandler; ;

                string tablename = "FT_SPCLCTRL";
                if (!app.createTable(tablename, "SP Credit Limit Control", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@" + tablename))
                {
                    if (!app.createField("@" + tablename, "NotifyV", "Notification", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "NA", true, SAPbobsCOM.BoFldSubTypes.st_None, "MSG_NOBLOCK:Message with No Block|MSG_BLOCK:Message with Block|NA:Not Applicable", "")) goto ErrorHandler;
                    
                    SAPbobsCOM.UserTable oUserTable = SAP.SBOCompany.UserTables.Item(tablename);
                    oUserTable.Code = "149";
                    oUserTable.Name = "Sales Quotation";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_NOBLOCK";
                    oUserTable.Add();

                    oUserTable.Code = "1250000100";
                    oUserTable.Name = "Sales Blanket Agreement";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_NOBLOCK";
                    oUserTable.Add();

                    oUserTable.Code = "139";
                    oUserTable.Name = "Sales Order";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_BLOCK";
                    oUserTable.Add();

                    oUserTable.Code = "FT_SPLAN";
                    oUserTable.Name = "Sales Planning";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "NA";
                    oUserTable.Add();

                    oUserTable.Code = "FT_TPPLAN";
                    oUserTable.Name = "Transport Planning";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "NA";
                    oUserTable.Add();
                }
                tablename = "FT_SPODCTRL";
                if (!app.createTable(tablename, "SP Overdue Control", SAPbobsCOM.BoUTBTableType.bott_NoObject)) goto ErrorHandler;
                if (!app.tableGotField("@" + tablename))
                {
                    if (!app.createField("@" + tablename, "NotifyV", "Notification", SAPbobsCOM.BoFieldTypes.db_Alpha, 50, "NA", true, SAPbobsCOM.BoFldSubTypes.st_None, "MSG_NOBLOCK:Message with No Block|MSG_BLOCK:Message with Block|NA:Not Applicable", "")) goto ErrorHandler;

                    SAPbobsCOM.UserTable oUserTable = SAP.SBOCompany.UserTables.Item(tablename);
                    oUserTable.Code = "149";
                    oUserTable.Name = "Sales Quotation";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_NOBLOCK";
                    oUserTable.Add();

                    oUserTable.Code = "1250000100";
                    oUserTable.Name = "Sales Blanket Agreement";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_NOBLOCK";
                    oUserTable.Add();

                    oUserTable.Code = "139";
                    oUserTable.Name = "Sales Order";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_BLOCK";
                    oUserTable.Add();

                    oUserTable.Code = "FT_SPLAN";
                    oUserTable.Name = "Sales Planning";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "NA";
                    oUserTable.Add();

                    oUserTable.Code = "FT_TPPLAN";
                    oUserTable.Name = "Transport Planning";
                    oUserTable.UserFields.Fields.Item("U_NotifyV").Value = "MSG_BLOCK";
                    oUserTable.Add();
                }


                if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
            }
            catch(Exception ex)
            {
                errmsg = ex.Message;
                goto ErrorHandler;
            }
            FT_ADDON.SAP.hideStatus();

            //app.createMenuItem("FT_GENSPERR", "Generate SPERRLOG", "43523", "8202", false, SAPbouiCOM.BoMenuType.mt_STRING);
            string username = SAP.SBOCompany.UserName;

            app.createMenuItem("FT_SPLAN", "Sales Planning", "2048", "2050", false, SAPbouiCOM.BoMenuType.mt_STRING);
            app.createMenuItem("FT_TPPLAN", "Transport Planning", "2048", "FT_SPLAN", false, SAPbouiCOM.BoMenuType.mt_STRING);
            app.createMenuItem("FT_CHARGE", "Charge Module", "2048", "FT_TPPLAN", false, SAPbouiCOM.BoMenuType.mt_STRING);

            if (ObjectFunctions.Approval("23", username))
                app.createMenuItem("SQAPPLIST", "Sales Quotation Approval", "2048", "FT_CHARGE", false, SAPbouiCOM.BoMenuType.mt_STRING);

            if (ObjectFunctions.Approval("17", username))
                app.createMenuItem("SOAPPLIST", "Sales Order Approval", "2048", "FT_CHARGE", false, SAPbouiCOM.BoMenuType.mt_STRING);

            if (ObjectFunctions.Approval("FT_SPLAN", username))
                app.createMenuItem("SPAPPLIST", "Sales Planning Approval", "2048", "FT_CHARGE", false, SAPbouiCOM.BoMenuType.mt_STRING);

            if (ObjectFunctions.Approval("FT_TPPLAN", username))
                app.createMenuItem("TPAPPLIST", "Transport Planning Approval", "2048", "FT_CHARGE", false, SAPbouiCOM.BoMenuType.mt_STRING);
            //GC.WaitForPendingFinalizers();
            return;

        ErrorHandler:
            if (FT_ADDON.SAP.SBOCompany.InTransaction) FT_ADDON.SAP.SBOCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
            FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Addon was teminated.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);
            FT_ADDON.SAP.hideStatus();
            if (errmsg != "") SAP.SBOApplication.MessageBox(errmsg, 1, "&Ok", "", "");

            System.Environment.Exit(0);
        }
        private bool checkTables()
        {
            try
            {
                FT_ADDON.SAP.setStatus("Checking Table...");

                string ls_sql = "IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[@FT_APSOC]') AND type in (N'U')) select 1";
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(ls_sql);
                if (rs.RecordCount > 0)
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                return false;
            }
       }
        private bool createTables()
        {
            try
            {
                FT_ADDON.SAP.setStatus("Creating Table : FT_APSOC");

                string ls_sql = "CREATE TABLE [dbo].[@FT_APSOC] ( DOCENTRY numeric(10,0) NOT NULL, LINENUM int NOT NULL, ITEMNO int NOT NULL, CONNO varchar(50) NULL , REF varchar(255) NULL , PRIMARY KEY (DOCENTRY,LINENUM,ITEMNO))";

                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)FT_ADDON.SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                rs.DoQuery(ls_sql);
                return true;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
                return false;
            }
        }
    }
}
