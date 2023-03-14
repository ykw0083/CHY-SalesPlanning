using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{

    class UserForm_ShipList
    {
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED:
                        if (pVal.ItemUID == "2")
                        {
                            SAPbouiCOM.Form oSForm = SAP.SBOApplication.Forms.Item(oForm.DataSources.UserDataSources.Item("fuid").Value.ToString());
                            //SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).String);
                                
                            oSForm.Freeze(false);
                            oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                            oSForm.Select();
                        }
                        break;

                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }

        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                SAPbouiCOM.Grid oGrid = null;

                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_CLICK:
                        if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
                            oGrid.Rows.SelectedRows.Add(pVal.Row);
                        }
                        break;
                    case SAPbouiCOM.BoEventTypes.et_FORM_CLOSE:
                        SAPbouiCOM.Form oSForm = SAP.SBOApplication.Forms.Item(oForm.DataSources.UserDataSources.Item("fuid").Value.ToString());
                        //SAPbouiCOM.Form oSForm = FT_ADDON.SAP.SBOApplication.Forms.Item(((SAPbouiCOM.EditText)oForm.Items.Item("FUID").Specific).String);

                        oSForm.Freeze(false);
                        oSForm.State = SAPbouiCOM.BoFormStateEnum.fs_Restore;
                        oSForm.Select();
                        break;

                    case SAPbouiCOM.BoEventTypes.et_DOUBLE_CLICK:
                        if (pVal.ItemUID == "grid" && pVal.Row >= 0)
                        {
                            oGrid = (SAPbouiCOM.Grid)oForm.Items.Item("grid").Specific;
                            if (oGrid.Rows.Count > 0)
                            {
                                long docentry = long.Parse(oForm.DataSources.DataTables.Item("list").GetValue(0, pVal.Row).ToString());
                                long sdoc = long.Parse(oForm.DataSources.DataTables.Item("list").GetValue(1, pVal.Row).ToString());
                                string docnum = oForm.DataSources.DataTables.Item("list").GetValue(2, pVal.Row).ToString();
                                if (docentry > 0)
                                    InitForm.shipdoc(oForm.UniqueID, sdoc, docentry, docnum);
                            }
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox(ex.Message, 1, "Ok", "", "");
            }
        }
    }
}