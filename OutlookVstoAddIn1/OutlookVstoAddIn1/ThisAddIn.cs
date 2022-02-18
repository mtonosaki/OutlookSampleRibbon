using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookVstoAddIn1
{
    public partial class ThisAddIn
    {
        private Outlook.Inspectors inspectors;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            inspectors = Application.Inspectors;
            inspectors.NewInspector += new Outlook.InspectorsEvents_NewInspectorEventHandler(Inspectors_NewInspector);
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        // When add a new email
        void Inspectors_NewInspector(Outlook.Inspector Inspector)
        {
            var newMailItem = Inspector.CurrentItem as Outlook.MailItem;
            if (newMailItem != null)
            {
                if (newMailItem.EntryID == null)
                {
                    newMailItem.Subject = "AUTO MAIL TITLE IS HERE !!";
                    var body = new StringBuilder();
                    body.AppendLine($"Hello! {inspectors.Session.CurrentUser.Name} - san");
                    body.AppendLine($"CurrentUser.AddressEntry.Name = {inspectors.Session.CurrentUser.AddressEntry.Name}");
                    body.AppendLine($"CurrentUser.AddressEntry.Address = {inspectors.Session.CurrentUser.AddressEntry.Address}");
                    body.AppendLine($"CurrentUser.AddressEntry.ID = {inspectors.Session.CurrentUser.AddressEntry.ID}");
                    body.AppendLine($"CurrentUser.AddressEntry.Members?.GetFirst()?.Name = {inspectors.Session.CurrentUser.AddressEntry.Members?.GetFirst()?.Name}");
                    body.AppendLine($"CurrentUser.AddressEntry.DisplayType = {inspectors.Session.CurrentUser.AddressEntry.DisplayType}");
                    body.AppendLine($"----------");
                    body.AppendLine($"CurrentUser.Application = {inspectors.Session.CurrentUser.Application}");
                    body.AppendLine($"CurrentUser.AutoResponse = {inspectors.Session.CurrentUser.AutoResponse}");
                    body.AppendLine($"CurrentUser.Class = {inspectors.Session.CurrentUser.Class}");
                    body.AppendLine($"CurrentUser.DisplayType = {inspectors.Session.CurrentUser.DisplayType}");
                    body.AppendLine($"CurrentUser.EntryID = {inspectors.Session.CurrentUser.EntryID}");
                    body.AppendLine($"CurrentUser.Index = {inspectors.Session.CurrentUser.Index}");
                    body.AppendLine($"CurrentUser.MeetingResponseStatus = {inspectors.Session.CurrentUser.MeetingResponseStatus}");
                    body.AppendLine($"CurrentUser.Parent = {inspectors.Session.CurrentUser.Parent}");
                    body.AppendLine($"CurrentUser.Resolved = {inspectors.Session.CurrentUser.Resolved}");
                    body.AppendLine($"CurrentUser.Sendable = {inspectors.Session.CurrentUser.Sendable}");
                    body.AppendLine($"CurrentUser.TrackingStatus = {inspectors.Session.CurrentUser.TrackingStatus}");
                    body.AppendLine($"CurrentUser.TrackingStatusTime = {inspectors.Session.CurrentUser.TrackingStatusTime}");
                    body.AppendLine($"CurrentUser.Type = {inspectors.Session.CurrentUser.Type}");
                    body.AppendLine($"----------");
                    body.AppendLine($"Session.DefaultStore.DisplayName = {inspectors.Session.DefaultStore.DisplayName}");
                    body.AppendLine($"Session.DefaultStore.ExchangeStoreType = {inspectors.Session.DefaultStore.ExchangeStoreType}");
                    newMailItem.Body = body.ToString();

                    // Hello! Manabu Tonosaki -san
                    // CurrentUser.AddressEntry.Name = Manabu Tonosaki
                    // CurrentUser.AddressEntry.Address = / o = ExchangeLabs / ou = Exchange Administrative Group(AAAAAAAA11AAAAA)/ cn = Recipients / cn = 1111aa11111a111aa11a11aaa11a1111 - mailbox1
                    // CurrentUser.AddressEntry.ID = 00000000AAA111A1AaaaaaaAAaA111111A1AA1111111111111111111...
                    // CurrentUser.AddressEntry.Members?.GetFirst()?.Name =
                    // CurrentUser.AddressEntry.DisplayType = olUser
                    // ----------
                    // CurrentUser.Application = Microsoft.Office.Interop.Outlook.ApplicationClass
                    // CurrentUser.AutoResponse =
                    // CurrentUser.Class = olRecipient
                    // CurrentUser.DisplayType = olUser
                    // CurrentUser.EntryID = 00000000AAA111A1AaaaaaaAAaA111111A1AA1111111111111111111...
                    // CurrentUser.Index = 0
                    // CurrentUser.MeetingResponseStatus = olResponseNone
                    // CurrentUser.Parent =
                    // CurrentUser.Resolved = True
                    // CurrentUser.Sendable = True
                    // CurrentUser.TrackingStatus = olTrackingNone
                    // CurrentUser.TrackingStatusTime = 01 / 01 / 4501 00:00:00
                    // CurrentUser.Type = 0
                    // ----------
                    // Session.DefaultStore.DisplayName = manabu@tomarika.com
                    // Session.DefaultStore.ExchangeStoreType = olPrimaryExchangeMailbox
                }

            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
