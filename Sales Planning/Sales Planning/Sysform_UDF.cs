using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Data;

namespace FT_ADDON.CHY
{
    class Sysform_UDF
    {
        public static void processRightClickEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal, ref bool BubbleEvent)
        {
        }
        public static void processRightClickEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ContextMenuInfo pVal)
        {
        }
        public static void processItemEventbefore(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal, ref bool BubbleEvent)
        {
        }
        public static void processItemEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.ItemEvent pVal)
        {
            try
            {
                switch (pVal.EventType)
                {
                    case SAPbouiCOM.BoEventTypes.et_FORM_LOAD:
                        SAPbouiCOM.Item oItemS = null;
                        SAPbouiCOM.Item oItem = oForm.Items.Add("U_seq", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oItem.Specific;
                        oEditText.DataBind.SetBound(true, "CUFD", "U_seq");
                        oItemS = oForm.Items.Item("7");
                        oItem.Top = oItemS.Top;
                        oItemS = oForm.Items.Item("8");
                        oItem.Left = oItemS.Left;
                        oItem.Width = oItemS.Width;

                        oItem = oForm.Items.Add("U_seq_t", SAPbouiCOM.BoFormItemTypes.it_STATIC);
                        SAPbouiCOM.StaticText oStaticText = (SAPbouiCOM.StaticText)oItem.Specific;
                        oStaticText.Caption = "Sequence";
                        oItemS = oForm.Items.Item("U_seq");
                        oItem.Top = oItemS.Top;
                        oItemS = oForm.Items.Item("10");
                        oItem.Left = oItemS.Left;
                        oItem.LinkTo = "U_seq";

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
    }
}
