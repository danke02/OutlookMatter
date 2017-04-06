using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMatterConsole
{
    public class Job
    {
        public string jobType;
        public DateTime startDateTime;
        public DateTime endDateTime;
        public string jobName;
        public string body;
        public bool isAllDayEvent;
        public string GetContent()
        {
            StringBuilder sb = new StringBuilder();
            if(isAllDayEvent == false)
                sb.AppendLine(string.Format("* [{0}] {1}~{4} : {2}({3})", jobType, startDateTime, jobName, body,endDateTime));
            else
                sb.AppendLine(string.Format("* [{0}] {1} : {2}({3})", jobType, startDateTime, jobName, body));
            return sb.ToString();
        }
    }
    class OutlookUtil
    {
        public static List<Job> GetAllCalendarItems()
        {
            List<Job> jobList = new List<Job>();
            Microsoft.Office.Interop.Outlook.Application oApp = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNamespace = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder CalendarFolder = null;
            Microsoft.Office.Interop.Outlook.Items outlookCalendarItems = null;

            oApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = oApp.GetNamespace("MAPI"); ;
            CalendarFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderCalendar);
            outlookCalendarItems = CalendarFolder.Items;
            outlookCalendarItems.IncludeRecurrences = true;
            
            foreach (Microsoft.Office.Interop.Outlook.AppointmentItem item in outlookCalendarItems)
            {
                if (item.IsRecurring)
                {
                    Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                    DateTime first = item.Start;
                    DateTime last = DateTime.Today.AddDays(30);
                    Microsoft.Office.Interop.Outlook.AppointmentItem recur = null;

                    for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                    {
                        try
                        {
                            recur = rp.GetOccurrence(cur);
                            //Console.WriteLine(recur.Subject + " -> " + cur.ToLongDateString());
                            jobList.Add(new OutlookMatterConsole.Job { startDateTime = cur, jobName =recur.Subject, jobType = "일정", body = recur.Body, endDateTime = recur.End, isAllDayEvent = recur.AllDayEvent });
                        }
                        catch(Exception e)
                        { }
                    }
                }
                else
                {
                    jobList.Add(new OutlookMatterConsole.Job { startDateTime = item.Start, jobName = item.Subject, jobType = "일정", body = item.Body, endDateTime = item.End, isAllDayEvent = item.AllDayEvent });
                    
                    //Console.WriteLine(item.Subject + " -> " + item.Start.ToLongDateString());
                }
            }
            return jobList;

        }
        public static List<Job>  GetAllTaskItems()
        {
            List<Job> jobList = new List<Job>();
            Microsoft.Office.Interop.Outlook.Application oApp = null;
            Microsoft.Office.Interop.Outlook.NameSpace mapiNamespace = null;
            Microsoft.Office.Interop.Outlook.MAPIFolder TasksFolder = null;
            Microsoft.Office.Interop.Outlook.Items outlookTaskItems = null;

            oApp = new Microsoft.Office.Interop.Outlook.Application();
            mapiNamespace = oApp.GetNamespace("MAPI"); ;
            TasksFolder = mapiNamespace.GetDefaultFolder(Microsoft.Office.Interop.Outlook.OlDefaultFolders.olFolderTasks);
            outlookTaskItems = TasksFolder.Items;
            outlookTaskItems.IncludeRecurrences = true;

            
            foreach (Microsoft.Office.Interop.Outlook.TaskItem item in outlookTaskItems)
            {
                if (item.IsRecurring)
                {
                    Microsoft.Office.Interop.Outlook.RecurrencePattern rp = item.GetRecurrencePattern();
                    DateTime first = item.StartDate;
                    DateTime last = DateTime.Today.AddDays(30);
                    Microsoft.Office.Interop.Outlook.AppointmentItem recur = null;



                    for (DateTime cur = first; cur <= last; cur = cur.AddDays(1))
                    {
                        try
                        {
                            recur = rp.GetOccurrence(cur);
                            jobList.Add(new OutlookMatterConsole.Job { startDateTime = cur, jobName = recur.Subject, jobType = "작업", body = recur.Body, endDateTime = recur.End, isAllDayEvent = recur.AllDayEvent });
                        }
                        catch
                        { }
                    }
                }
                else
                {
                    jobList.Add(new OutlookMatterConsole.Job { startDateTime = item.StartDate, jobName = item.Subject, jobType = "작업", body = item.Body, endDateTime = item.DueDate, isAllDayEvent = false });
                }
            }
            return jobList;
        }
    }
}
