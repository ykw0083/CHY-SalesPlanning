using System;
using System.Collections.Generic;
using System.Text;

namespace FT_ADDON.CHY
{
    class Sysform_Approval50103
    {
        public static void processDataEventafter(SAPbouiCOM.Form oForm, ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo)
        {
            try
            {
                //switch (BusinessObjectInfo.EventType)
                //{
                //    case SAPbouiCOM.BoEventTypes.et_FORM_DATA_UPDATE:
                //        if (!BusinessObjectInfo.ActionSuccess) return;
                //        SAPbouiCOM.DBDataSource ods = oForm.DataSources.DBDataSources.Item("OWDD");
                //        string docentry = ods.GetValue("DraftEntry", 0);


                //        SAPbouiCOM.ComboBox oComboBox = (SAPbouiCOM.ComboBox)oForm.Items.Item("28").Specific;
                //        SAPbouiCOM.EditText oEditText = (SAPbouiCOM.EditText)oForm.Items.Item("23").Specific;
                //        string remarks = oEditText.Value;
                //        string status = oComboBox.Selected.Value;


                //        string approval = "";
                //        switch (status)
                //        {
                //            case "Y":
                //                approval = "Y";
                //                break;
                //            case "W":
                //                approval = "W";
                //                break;
                //            case "N":
                //                approval = "N";
                //                break;
                //        }
                        
                //        string sql = "Update [ODRF] set U_DSAPP = '" + approval + "' " +
                //                    " where DocEntry = " + docentry;

                //        SAPbobsCOM.Recordset oRC = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                //        oRC.DoQuery(sql);

                //        break;
                //}
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Data Event After " + ex.Message, 1, "Ok", "", "");
            }
        }

    }
}
