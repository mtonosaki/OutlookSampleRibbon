using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using Outlook = Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;

namespace OutlookVstoAddIn1
{
    public partial class Ribbon1
    {
        private Outlook.Inspectors inspectors;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            var Application = Globals.ThisAddIn.Application;
            var Inspector = Application.ActiveInspector();
            Outlook.AppointmentItem newAppointment = Inspector.CurrentItem;
            newAppointment.Start = DateTime.Now.AddHours(2);
            newAppointment.End = DateTime.Now.AddHours(3);
            newAppointment.AllDayEvent = false;
            newAppointment.Subject = "DEMO Meeting";
            var body = new StringBuilder();
            body.AppendLine($"Let us DEMO meeting later !");
            body.AppendLine($"https://aaa.example.com/agenda/new/{Guid.NewGuid()}&owner={Inspector.Session.CurrentUser.Name}");
            newAppointment.Body = body.ToString();
            //newAppointment.Location = "ConferenceRoom #999";
            //var sentTo = newAppointment.Recipients;
            //sentTo.Add("manabu").Type = (int)Outlook.OlMeetingRecipientType.olOptional;
            //sentTo.Add("riwa").Type = (int)Outlook.OlMeetingRecipientType.olRequired;
            //sentTo.ResolveAll();
            //newAppointment.Save();
            newAppointment.Display(true);
        }

        // HACK: 新規アイテムを作成する場合
        // var newAppointment = Application.CreateItem(OlItemType.olAppointmentItem);
        // Application.CreateItem(OlItemType.olAppointmentItem);
        //    ...
        // newAppointment.Save();

    }
}
