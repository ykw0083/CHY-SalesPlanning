using System;
using System.Collections.Generic;
using System.Text;
using System.Net;
using System.Net.Mail;
using SAPbobsCOM;

namespace FT_ADDON.CHY
{
    class EmailClass
    {
        public string ObjType { get; set; }
        public string EmailSubject { get; set; }
        public string EmailMsg { get; set; }

        public EmailClass()
        {
        }

        public bool SendEmail()
        {
            try
            {
                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Sending Email...", SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Warning);

                SAPbobsCOM.Recordset oRS = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);
                oRS.DoQuery("select U_emailaddr, U_emailname, U_host, U_port, U_emailuid, isnull(U_emailpwd,''), isnull(U_enablessl,'N') from [@FT_SEMAIL] where U_objcode = '" + ObjType + "'");

                oRS.MoveFirst();
                string emailaddr = oRS.Fields.Item(0).Value.ToString();
                string emailname = oRS.Fields.Item(1).Value.ToString();
                string host = oRS.Fields.Item(2).Value.ToString(); // smtp.gmail.com
                int port = int.Parse(oRS.Fields.Item(3).Value.ToString()); // 587
                string emailuid = oRS.Fields.Item(4).Value.ToString(); // email user id
                string emailpwd = oRS.Fields.Item(5).Value.ToString(); // email password
                bool enablessl = oRS.Fields.Item(6).Value.ToString() == "Y" ? true : false;
                if (port <= 0)
                {
                    FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Port is invalid.", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    return false;
                }

                MailMessage oMail = new MailMessage();
                MailAddress addFr = new MailAddress(emailaddr, emailname);

                oMail.From = addFr;
                oMail.Subject = EmailSubject;
                oMail.Body = EmailMsg;

                oRS.DoQuery("select U_emailaddr, U_emailname from [@FT_REMAIL] where U_objcode = '" + ObjType + "'");

                oRS.MoveFirst();
                while (!oRS.EoF)
                {
                    emailaddr = oRS.Fields.Item(0).Value.ToString();
                    emailname = oRS.Fields.Item(1).Value.ToString();
                    oMail.To.Add(new MailAddress(emailaddr, emailname));

                    oRS.MoveNext();
                }


                SmtpClient client = new SmtpClient();
                client.Port = port;
                client.Host = host;
                client.Timeout = 20000;
                if (enablessl)
                    client.EnableSsl = true;
                else
                    client.EnableSsl = false;
                client.DeliveryMethod = SmtpDeliveryMethod.Network;
                client.UseDefaultCredentials = false;
                client.Credentials = new System.Net.NetworkCredential(emailuid, emailpwd);

                client.Send(oMail);

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Email Sent", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                return true;
            }
            catch (Exception ex)
            {
                SAP.SBOApplication.MessageBox("Send email error " + ex.Message, 1, "Ok", "", "");
                return false;
            }
        }
    }

    class SystemMsgClass
    {
        public string ObjType { get; set; }
        public string TrxType { get; set; }
        public string MsgSubject { get; set; }
        public string Msg { get; set; }
        public string ColumnName { get; set; }
        public string LinkObj { get; set; }
        public string LinkKey { get; set; }
        public string LineValue { get; set; }
        public List<string> UserList = new List<string>();

        public SystemMsgClass()
        {
            LinkObj = "";
        }

        public void SendMsg()
        {

            if (TrxType == "A" || TrxType == "U")
            { }
            else
                return;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            oRecordSet.DoQuery("select T1.USER_CODE from [@FT_RSYSMSG] T0 inner join OUSR T1 on T0.U_userid = T1.USER_CODE and T1.Locked = 'N' where T0.U_objcode = '" + ObjType + "' and T0.U_trxtype in ('B', '" + TrxType + "')");

            while (!oRecordSet.EoF)
            {
                if (oRecordSet.Fields.Item(0).Value.ToString().Trim() != "")
                {
                    UserList.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim());
                }
                oRecordSet.MoveNext();
            }

            if (UserList.Count == 0) return;

            SAPbobsCOM.CompanyService oService = SAP.SBOCompany.GetCompanyService();
            SAPbobsCOM.MessagesService oMessageService = (SAPbobsCOM.MessagesService)oService.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService);
            SAPbobsCOM.Message oMessage = null;
            SAPbobsCOM.MessageDataColumns pMessageDataColumns = null;
            SAPbobsCOM.MessageDataColumn pMessageDataColumn = null;
            SAPbobsCOM.MessageDataLines oLines = null;
            SAPbobsCOM.MessageDataLine oLine = null;
            SAPbobsCOM.RecipientCollection oRecipientCollection = null;

            try
            {
                // get the data interface for the new message
                oMessage = ((SAPbobsCOM.Message)(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)));

                // fill subject
                oMessage.Subject = MsgSubject;

                oMessage.Text = Msg;

                // Add Recipient 
                oRecipientCollection = oMessage.RecipientCollection;
                int cnt = 0;
                foreach (string user in UserList)
                {
                    oRecipientCollection.Add();

                    // send internal message
                    oRecipientCollection.Item(cnt).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;

                    // add existing user name
                    //oRecipientCollection.Item(0).UserCode = GetUserCode(user);
                    oRecipientCollection.Item(cnt).UserCode = user;

                    cnt++;
                }

                // get columns data
                pMessageDataColumns = oMessage.MessageDataColumns;
                // get column
                pMessageDataColumn = pMessageDataColumns.Add();
                // set column name
                pMessageDataColumn.ColumnName = ColumnName;

                if (LinkObj.Trim() != "")
                {
                    // set link to a real object in the application
                    pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES;

                    // get lines
                    oLines = pMessageDataColumn.MessageDataLines;
                    // add new line
                    oLine = oLines.Add();
                    // set the line value
                    oLine.Value = LineValue;

                    // set the link to BusinessPartner (the object type for Bp is 2)
                    oLine.Object = LinkObj;
                    // set the Bp code
                    oLine.ObjectKey = LinkKey;
                }
                else
                {
                    // get lines
                    oLines = pMessageDataColumn.MessageDataLines;
                    // add new line
                    oLine = oLines.Add();
                    // set the line value
                    oLine.Value = LineValue;
                }

                // send the message
                oMessageService.SendMessage(oMessage);

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Sent System Message " + MsgSubject, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {

                SAP.SBOApplication.MessageBox("SendMsg " + ex.Message, 1, "Ok", "", "");

            }

        }
        private string GetUserCode(string sUserName)
        {

            SAPbobsCOM.Users oUsers;
            SAPbobsCOM.Recordset oRecordSet;

            try
            {
                //get users object
                oUsers = (SAPbobsCOM.Users)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                // Get a new Recordset object
                oRecordSet = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Perform the SELECT statement.
                // The query result will be loaded
                // into the Recordset object
                oRecordSet.DoQuery("Select USER_CODE from OUSR");

                // Asign (link) the Recordset object
                // to the Browser.Recordset property
                oUsers.Browser.Recordset = oRecordSet;

                //find the username
                while (oUsers.Browser.EoF == false)
                {
                    if (oUsers.UserName == sUserName)
                    {
                        break;
                    }
                    oUsers.Browser.MoveNext();
                }

                //return the User Code
                return oUsers.UserCode;

            }
            catch (Exception ex)
            {

                SAP.SBOApplication.MessageBox("GetUserCode " + ex.Message, 1, "Ok", "", "");

            }

            return "";
        }

    }
    class ApprovalMsgClass
    {
        public string ObjType { get; set; }
        public string TrxType { get; set; }
        public string MsgSubject { get; set; }
        public string Msg { get; set; }
        public string ColumnName { get; set; }
        public string LinkObj { get; set; }
        public string LinkKey { get; set; }
        public string LineValue { get; set; }
        public bool IsApprove { get; set; }
        public List<string> UserList = new List<string>();

        public ApprovalMsgClass()
        {
            IsApprove = false;
            LinkObj = "";
        }

        public void SendMsg()
        {

            if (TrxType == "A" || TrxType == "U")
            { }
            else
                return;

            SAPbobsCOM.Recordset oRecordSet = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            if (IsApprove)
                oRecordSet.DoQuery("select T1.USER_CODE from [@FT_RAPPMSG] T0 inner join OUSR T1 on T0.U_userid = T1.USER_CODE and T1.Locked = 'N' where T0.U_objcode = '" + ObjType + "' and T0.U_trxtype in ('B', '" + TrxType + "') and isnull(U_isapp,'N') = 'Y'");
            else
                oRecordSet.DoQuery("select T1.USER_CODE from [@FT_RAPPMSG] T0 inner join OUSR T1 on T0.U_userid = T1.USER_CODE and T1.Locked = 'N' where T0.U_objcode = '" + ObjType + "' and T0.U_trxtype in ('B', '" + TrxType + "') and isnull(U_isapp,'N') = 'N'");

            while (!oRecordSet.EoF)
            {
                if (oRecordSet.Fields.Item(0).Value.ToString().Trim() != "")
                {
                    UserList.Add(oRecordSet.Fields.Item(0).Value.ToString().Trim());
                }
                oRecordSet.MoveNext();
            }

            if (UserList.Count == 0) return;

            SAPbobsCOM.CompanyService oService = SAP.SBOCompany.GetCompanyService();
            SAPbobsCOM.MessagesService oMessageService = (SAPbobsCOM.MessagesService)oService.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService);
            SAPbobsCOM.Message oMessage = null;
            SAPbobsCOM.MessageDataColumns pMessageDataColumns = null;
            SAPbobsCOM.MessageDataColumn pMessageDataColumn = null;
            SAPbobsCOM.MessageDataLines oLines = null;
            SAPbobsCOM.MessageDataLine oLine = null;
            SAPbobsCOM.RecipientCollection oRecipientCollection = null;

            try
            {
                // get the data interface for the new message
                oMessage = ((SAPbobsCOM.Message)(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage)));

                // fill subject
                oMessage.Subject = MsgSubject;

                oMessage.Text = Msg;

                // Add Recipient 
                oRecipientCollection = oMessage.RecipientCollection;
                int cnt = 0;
                foreach (string user in UserList)
                {
                    oRecipientCollection.Add();

                    // send internal message
                    oRecipientCollection.Item(cnt).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES;

                    // add existing user name
                    //oRecipientCollection.Item(0).UserCode = GetUserCode(user);
                    oRecipientCollection.Item(cnt).UserCode = user;

                    cnt++;
                }

                // get columns data
                pMessageDataColumns = oMessage.MessageDataColumns;
                // get column
                pMessageDataColumn = pMessageDataColumns.Add();
                // set column name
                pMessageDataColumn.ColumnName = ColumnName;

                if (LinkObj.Trim() != "")
                {
                    // set link to a real object in the application
                    pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES;

                    // get lines
                    oLines = pMessageDataColumn.MessageDataLines;
                    // add new line
                    oLine = oLines.Add();
                    // set the line value
                    oLine.Value = LineValue;

                    // set the link to BusinessPartner (the object type for Bp is 2)
                    oLine.Object = LinkObj;
                    // set the Bp code
                    oLine.ObjectKey = LinkKey;
                }
                else
                {
                    // get lines
                    oLines = pMessageDataColumn.MessageDataLines;
                    // add new line
                    oLine = oLines.Add();
                    // set the line value
                    oLine.Value = LineValue;
                }

                // send the message
                oMessageService.SendMessage(oMessage);

                FT_ADDON.SAP.SBOApplication.StatusBar.SetText("Sent System Message " + MsgSubject, SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Success);

            }
            catch (Exception ex)
            {

                SAP.SBOApplication.MessageBox("SendMsg " + ex.Message, 1, "Ok", "", "");

            }

        }
        private string GetUserCode(string sUserName)
        {

            SAPbobsCOM.Users oUsers;
            SAPbobsCOM.Recordset oRecordSet;

            try
            {
                //get users object
                oUsers = (SAPbobsCOM.Users)SAP.SBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oUsers);

                // Get a new Recordset object
                oRecordSet = (SAPbobsCOM.Recordset)SAP.SBOCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // Perform the SELECT statement.
                // The query result will be loaded
                // into the Recordset object
                oRecordSet.DoQuery("Select USER_CODE from OUSR");

                // Asign (link) the Recordset object
                // to the Browser.Recordset property
                oUsers.Browser.Recordset = oRecordSet;

                //find the username
                while (oUsers.Browser.EoF == false)
                {
                    if (oUsers.UserName == sUserName)
                    {
                        break;
                    }
                    oUsers.Browser.MoveNext();
                }

                //return the User Code
                return oUsers.UserCode;

            }
            catch (Exception ex)
            {

                SAP.SBOApplication.MessageBox("GetUserCode " + ex.Message, 1, "Ok", "", "");

            }

            return "";
        }

    }
}
