using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutlookMatterConsole
{
    class Program
    {

        static void Main(string[] args)
        {
            StringBuilder todaysJob = new StringBuilder(), weeksJob = new StringBuilder();
            List<Job> todaysJobList1 = new List<Job>(OutlookUtil.GetAllCalendarItems().Where(job => job.startDateTime >= DateTime.Today && job.startDateTime < DateTime.Today.AddDays(1)));
            todaysJob.AppendLine(string.Format("### 오늘 일정 {0}개", todaysJobList1.Count));
            foreach (Job job in todaysJobList1)
            {
                todaysJob.Append(job.GetContent());
            }
            todaysJob.AppendLine();
            List <Job> todaysJobList2 = new List<Job>(OutlookUtil.GetAllTaskItems().Where(job => job.startDateTime >= DateTime.Today && job.startDateTime < DateTime.Today.AddDays(1)));
            todaysJob.AppendLine(string.Format("### 오늘 작업 {0}개", todaysJobList2.Count));
            foreach (Job job in todaysJobList2)
            {
                todaysJob.Append(job.GetContent());
            }
            todaysJob.AppendLine();

            List<Job> weeksJobList1 = new List<Job>(OutlookUtil.GetAllCalendarItems().Where(job => job.startDateTime >= DateTime.Today && job.startDateTime < DateTime.Today.AddDays(8)));
            weeksJob.AppendLine(string.Format("### 금주 일정 {0}개", weeksJobList1.Count));
            foreach (Job job in weeksJobList1)
            {
                weeksJob.Append(job.GetContent());
            }
            weeksJob.AppendLine();

            List <Job> weeksJobList2 = new List<Job>(OutlookUtil.GetAllTaskItems().Where(job => job.startDateTime >= DateTime.Today && job.startDateTime < DateTime.Today.AddDays(8)));
            weeksJob.AppendLine(string.Format("### 금주 작업 {0}개", weeksJobList2.Count));
            foreach (Job job in weeksJobList2)
            {
                weeksJob.Append(job.GetContent());
            }
            weeksJob.AppendLine();

            Console.WriteLine(todaysJob.ToString());
            Console.WriteLine(weeksJob.ToString());

            SlackClient mmClient = new SlackClient("http://10.239.203.30:3080/hooks/ig6xdi1bgbgijff8tttcgjmunr");
            mmClient.PostMessage(todaysJob.ToString(),"Outlook Matter");
            mmClient.PostMessage(weeksJob.ToString(), "Outlook Matter");

            Console.Read();
        }
        
    }

}
