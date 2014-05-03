using Outlook = Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace OutlookTasks
{
    public class OutlookGrab
    {
        static List<TaskItem> tasks = new List<TaskItem>();

        # region methods
        public static List<ListViewItem> GetListItems(string Range)
        {
            int daysFilter = 1; //TODO check this is resetting correctly.
            var outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            Folder folder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderTasks) as Folder;

            # region adjust for day of the week
            switch (DateTime.Today.DayOfWeek)
            {
                // go back to Sunday for start of week
                //need to insert date modifiers

                case DayOfWeek.Sunday:
                    daysFilter += 1;
                    break;

                case DayOfWeek.Monday:
                    daysFilter += 2;
                    break;

                case DayOfWeek.Tuesday:
                    daysFilter += 3;
                    break;

                case DayOfWeek.Wednesday:
                    daysFilter += 4;
                    break;

                case DayOfWeek.Thursday:
                    daysFilter += 5;
                    break;

                case DayOfWeek.Friday:
                    daysFilter += 6;
                    break;

                case DayOfWeek.Saturday:
                    daysFilter += 7;
                    break;
            }
            # endregion

            # region set number of days to go back
            switch (Range)
            {
                case "Today":
                    daysFilter = 1;
                    break;

                case "This Week":
                    daysFilter += 0;
                    break;

                case "Last Week":
                    daysFilter += 7;
                    break;
            }
            # endregion

            List<TaskItem> sortedItems = sortItems(folder);
            List<ListViewItem> ListItems = new List<ListViewItem>();

            # region Filter tasks completed this week
            foreach (TaskItem tItem in sortedItems)
            {
                int daysPassed = (DateTime.Today.Date - tItem.DateCompleted.Date).Days;
                if (tItem.Complete && daysPassed < daysFilter)
                {
                    ListViewItem lItem = new ListViewItem();

                    lItem.Text = tItem.Subject;
                    lItem.SubItems.Add(tItem.Complete.ToString());
                    lItem.SubItems.Add(tItem.DateCompleted.DayOfWeek.ToString());
                    lItem.SubItems.Add(tItem.DateCompleted.ToShortDateString());
                    lItem.SubItems.Add(tItem.DueDate.ToShortDateString());

                    ListItems.Add(lItem);
                    tasks.Add(tItem);
                }
            }
            # endregion

            Debug.WriteLine("{0}, {1}", "daysFilter", daysFilter);

            return ListItems;
        }

        public static List<TaskItem> getTasks(string Range)
        {
            int daysFilter = 1; //TODO check this is resetting correctly.
            var outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            Folder folder = outlookApp.Session.GetDefaultFolder(OlDefaultFolders.olFolderTasks) as Folder;

            # region adjust for day of the week
            switch (DateTime.Today.DayOfWeek)
            {
                // go back to Sunday for start of week
                //need to insert date modifiers

                case DayOfWeek.Sunday:
                    daysFilter += 1;
                    break;

                case DayOfWeek.Monday:
                    daysFilter += 2;
                    break;

                case DayOfWeek.Tuesday:
                    daysFilter += 3;
                    break;

                case DayOfWeek.Wednesday:
                    daysFilter += 4;
                    break;

                case DayOfWeek.Thursday:
                    daysFilter += 5;
                    break;

                case DayOfWeek.Friday:
                    daysFilter += 6;
                    break;

                case DayOfWeek.Saturday:
                    daysFilter += 7;
                    break;
            }
            # endregion

            # region set number of days to go back
            switch (Range)
            {
                case "Today":
                    daysFilter = 1;
                    break;

                case "This Week":
                    daysFilter += 0;
                    break;

                case "Last Week":
                    daysFilter += 7;
                    break;
            }
            # endregion

            List<TaskItem> sortedItems = sortItems(folder);

            return sortedItems;
        }

        public static RichTextBox DisplayResults(Font f)
        {
            List<TaskItem>[] tsk = new List<TaskItem>[7];
            string[] names = new string[7];
            RichTextBox RTBox = new RichTextBox();
            RTBox.Font = f;

            # region create tasks for days of the week
            for (int i = 0; i < tsk.Length; i++)
            {
                tsk[i] = new List<TaskItem>();
            }
            # endregion

            # region switch (day of week)
            foreach (TaskItem t in tasks)
            {
                switch (t.DateCompleted.DayOfWeek)
                {
                    case DayOfWeek.Sunday:
                        tsk[0].Add(t);
                        names[0] = "Sunday";
                        break;

                    case DayOfWeek.Monday:
                        names[1] = "Monday";
                        tsk[1].Add(t);
                        break;

                    case DayOfWeek.Tuesday:
                        names[2] = "Tuesday";
                        tsk[2].Add(t);
                        break;

                    case DayOfWeek.Wednesday:
                        names[3] = "Wednesday";
                        tsk[3].Add(t);
                        break;

                    case DayOfWeek.Thursday:
                        names[4] = "Thursday";
                        tsk[4].Add(t);
                        break;

                    case DayOfWeek.Friday:
                        names[5] = "Friday";
                        tsk[5].Add(t);
                        break;

                    case DayOfWeek.Saturday:
                        names[6] = "Saturday";
                        tsk[6].Add(t);
                        break;
                }
            }
            # endregion

            List<TaskItem> list;

            # region insert tasks
            for (int i = 0; i < tsk.Length; i++)
            {
                list = tsk[i];

                if (list.Count > 0)
                {
                    RTBox.SelectionBullet = false;
                    RTBox.AppendText("\r\n" + names[i] + "\r\n");
                    RTBox.SelectionBullet = true;

                    foreach (TaskItem t in list)
                    {
                        RTBox.AppendText(t.Subject + "\r\n");
                    }
                }
            }
            # endregion

            return RTBox;
        }

        public static void clear()
        {
            tasks.Clear();
        }

        private static List<TaskItem> sortItems(Folder f)
        {
            List<TaskItem> sItems = new List<TaskItem>();

            foreach (TaskItem t in f.Items)
            {
                sItems.Add(t);
            }

            sItems.Sort((x, y) => x.DateCompleted.CompareTo(y.DateCompleted));

            return sItems;
        }
        # endregion
    }

    public class PassToExcel
    {
        public static string ParseListView()
        {
            var outlookApp = Marshal.GetActiveObject("Outlook.Application") as Outlook.Application;
            var selection = outlookApp.ActiveExplorer().Selection;

            string tasks = "";


            foreach (object o in selection)
            {
                if (o is TaskItem)
                {
                    TaskItem t = o as TaskItem;
                    if ( t.StartDate > (DateTime.Today + new TimeSpan(360,0,0,0,0)))
                        t.StartDate = DateTime.Today.Date;

                    if (t.DueDate > (DateTime.Today + new TimeSpan(360, 0, 0, 0, 0)))
                        t.DueDate = DateTime.Today.Date;

                    tasks += (t.Subject + "\t" + t.StartDate.ToShortDateString() + "\t" + t.DueDate.ToShortDateString() + "\r\n");
                    Debug.WriteLine(t.Subject + "\t" + t.StartDate.ToShortDateString() + "\t" + t.DueDate.ToShortDateString() + "\r\n");
                }
            }

            return tasks;
        }

        public static string ParseTaskString(string s)
        {
            # region setup
            string result = "";
            string[] sArr = s.Split(new char[]{'\n'},StringSplitOptions.RemoveEmptyEntries);
            # endregion

            # region main work
            for (int i = 0; i < sArr.Length; i++) //i = 1 to skip titles
            {
                //split into sections
                //0 subject, 1 startdate, 2 duedate, 3 organiser, 4 status, 5 % complete, 6 contacts, 7 categories
                string[] tempArr = sArr[i].Split('\t');

                # region skip titles but process everything else
                if (i > 0)
                {
                    for (int tr = 1; tr < 3; tr++)
                    {
                        string[] ts = tempArr[tr].Split(' ');   //remove leading day name

                        //split into parts of date
                        ts = ts[1].Split('/');
                        int day = int.Parse(ts[1]);
                        int month = int.Parse(ts[0]);
                        int year = int.Parse(ts[2]);

                        DateTime dt = new DateTime(year, month, day);

                        tempArr[tr] = dt.ToShortDateString();
                    }
                }
                # endregion

                # region update sArr[i] now that everything has been processed
                sArr[i] = "";
                for (int rst = 0; rst < 3; rst++)
                    sArr[i] += tempArr[rst] + "\t";

                sArr[i] = sArr[i].TrimEnd('\t');
                # endregion
            }


            # endregion

            # region end
            for (int i = 1; i < sArr.Length; i++)
                result += sArr[i] + "\r\n";

            return result;
            # endregion
        }
    }
}
