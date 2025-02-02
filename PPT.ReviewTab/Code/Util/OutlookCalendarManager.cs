using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;


namespace PPT.ReviewTab.Code.Util
{


    public static class OutlookCalendarManager
    {
        public static Outlook.AppointmentItem GetActiveCalendarItem()
        {
            Outlook.Application outlookApp = new Outlook.Application();

            // Get the currently open Outlook item
            Outlook.Inspector activeInspector = outlookApp.ActiveInspector();

            if (activeInspector != null && activeInspector.CurrentItem is Outlook.AppointmentItem appointment)
            {
                return appointment; // Return the opened calendar event
            }

            return null; // No calendar item is open
        }

        public static List<string> GetAttendees()
        {
            Outlook.AppointmentItem appointment = GetActiveCalendarItem();

            if (appointment == null) 
                return null;

            List<string> names = GetAttendees(appointment);

            return names;
        }


        private static List<string> GetAttendees(Outlook.AppointmentItem appointment)
        {
            
            List<string> attendees = new List<string>();

            if (appointment == null) return attendees;


            foreach (Outlook.Recipient recipient in appointment.Recipients)
            {
                attendees.Add($"{recipient.Name}");
            }

            return attendees;
        }


    }
}
