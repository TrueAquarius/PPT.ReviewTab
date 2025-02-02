using System;
using System.Collections.Generic;
using Outlook = Microsoft.Office.Interop.Outlook;



namespace PPT.ReviewTab.Code.Util
{
    public static class OutlookCalendarManager
    {
        public static Outlook.AppointmentItem GetActiveCalendarItem()
        {
            try
            {
                Outlook.Application outlookApp = new Outlook.Application();

                if(outlookApp == null) return null; // Cannot reach Outlook

                // Get the currently open Outlook item
                Outlook.Inspector activeInspector = outlookApp.ActiveInspector();

                if (activeInspector != null && activeInspector.CurrentItem is Outlook.AppointmentItem appointment)
                {
                    return appointment; // Return the opened calendar event
                }

                return null; // No calendar item is open
            }
            catch (Exception)
            {
                return null;      
            }
        }



        public static List<string> GetAttendeesFromActiveOutlookCalendarItem()
        {
            Outlook.AppointmentItem appointment = GetActiveCalendarItem();

            if (appointment == null) 
                return null;

            List<string> names = GetAttendeesFromOutlookCalendarItem(appointment);

            return names;
        }


        private static List<string> GetAttendeesFromOutlookCalendarItem(Outlook.AppointmentItem appointment)
        {
            if (appointment == null) return null;

            List<string> attendees = new List<string>();

            try
            {
                foreach (Outlook.Recipient recipient in appointment.Recipients)
                {
                    attendees.Add($"{recipient.Name}");
                }
            }
            catch (Exception)
            {
                return null;
            }

            return attendees;
        }
    }
}
