using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class MenuEvent
    {
        public static void processMenuEvent(ref SAPbouiCOM.MenuEvent pVal)
        {
            switch (pVal.MenuUID)
            {
                case "FT_GENSPERR":
                    InitForm.genFT_SPERRLOG();
                    break;
                case "FT_SPLAN":
                    InitForm.FT_SPLANscreenpainter("FT_SPLAN", "FT_SPLAN1", "");
                    break;
                case "FT_TPPLAN":
                    InitForm.FT_SPLANscreenpainter("FT_TPPLAN", "FT_TPPLAN1", "");
                    break;
                case "FT_CHARGE":
                    InitForm.FT_SPLANscreenpainter("FT_CHARGE", "FT_CHARGE1", "FT_CHARGE2");
                    break;
                //case "TPAPPLIST":
                //    InitForm.approval(pVal.MenuUID, "002", "17", "ORDR", "RDR1", "Sales Order (TP) Approval", true, "FT_TPPLAN");
                //    break;
                case "SQAPPLIST":
                    InitForm.approval(pVal.MenuUID, "005", "23", "OQUT", "QUT1", "Sales Quotation Approval", true, "");
                    //InitForm.approval(pVal.MenuUID, "001", "112", "ODRF", "DRF1", "Draft Sales Order Approval", true, "");
                    break;
                case "SOAPPLIST":
                    InitForm.approval(pVal.MenuUID, "002", "17", "ORDR", "RDR1", "Sales Order Approval", true, "");
                    //InitForm.approval(pVal.MenuUID, "001", "112", "ODRF", "DRF1", "Draft Sales Order Approval", true, "");
                    break;
                case "SPAPPLIST":
                    InitForm.approval(pVal.MenuUID, "003", "FT_SPLAN", "@FT_SPLAN", "FT_SPLAN1", "Sales Planning Approval", false, "17");
                    break;
                case "TPAPPLIST":
                    InitForm.approval(pVal.MenuUID, "004", "FT_TPPLAN", "@FT_TPPLAN", "FT_TPPLAN1", "Transport Planning Approval", false, "17");
                    break;
                default:
                    break;
            }

        }
        public static void processMenuEvent2(ref SAPbouiCOM.MenuEvent pVal, ref bool BubbleEvent)
        {
            switch (pVal.MenuUID)
            {
                case "1281": //find
                case "1282": //add
                case "1283": //remove
                case "1284": //cancel
                case "1285": //restore
                case "1286": //close
                case "1287": //duplicate
                case "1288": //next record
                case "1289": //previous record
                case "1290": //first record
                case "1291": //last record
                case "1292": //add row
                case "1293": //delete row
                case "1294": //duplicate row
                case "1295": //copy cell above
                case "1296": //copy cell below
                case "1299": //close row
                    {
                        SAPbouiCOM.Form oForm = null;
                        try
                        {
                            oForm = FT_ADDON.SAP.SBOApplication.Forms.ActiveForm;
                        }
                        catch { }
                        if (oForm == null) return;

                        if (pVal.BeforeAction)
                        {
                            if (oForm.TypeEx == "FT_CONM")
                                UserForm_CONmodified.processMenuEventbefore(oForm, ref pVal, ref BubbleEvent);
                            else if (oForm.TypeEx == "FT_DOPTM")
                                UserForm_CONmodified.processMenuEventbefore(oForm, ref pVal, ref BubbleEvent);
                            else if (oForm.TypeEx == "FT_SHIPD")
                                UserForm_ShipDoc.processMenuEventbefore(oForm, ref pVal, ref BubbleEvent);
                            else if (oForm.TypeEx == "FT_APSA")
                                UserForm_APSA.processMenuEventbefore(oForm, ref pVal, ref BubbleEvent);
                            else if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN" || oForm.TypeEx == "FT_CHARGE")
                                UserForm_SalesPlanning.processMenuEventbefore(oForm, ref pVal, ref BubbleEvent);
                            else if (oForm.TypeEx == "133")
                                Sysform_SalesInvoice.processMenuEventbefore(oForm, ref pVal, ref BubbleEvent);
                        }
                        else
                        {
                            if (oForm.TypeEx == "FT_CONM")
                                UserForm_CONmodified.processMenuEventafter(oForm, ref pVal);
                            else if (oForm.TypeEx == "FT_DOPTM")
                                UserForm_CONmodified.processMenuEventafter(oForm, ref pVal);
                            else if (oForm.TypeEx == "FT_SHIPD")
                                UserForm_ShipDoc.processMenuEventafter(oForm, ref pVal);
                            else if (oForm.TypeEx == "FT_APSA")
                                UserForm_APSA.processMenuEventafter(oForm, ref pVal);
                            else if (oForm.TypeEx == "FT_SPLAN" || oForm.TypeEx == "FT_TPPLAN" || oForm.TypeEx == "FT_CHARGE")
                                UserForm_SalesPlanning.processMenuEventafter(oForm, ref pVal);
                            else if (oForm.TypeEx == "133")
                                Sysform_SalesInvoice.processMenuEventafter(oForm, ref pVal);
                        }
                    }
                    break;
                default:
                    break;
            }

        }
    }
}
