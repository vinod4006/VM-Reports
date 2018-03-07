using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using Newtonsoft;
using Newtonsoft.Json;
using JIRASPRINTREPORT;
using System.Reflection;
using Newtonsoft.Json.Linq;
using Microsoft.Office.Interop.Excel;
using System.Drawing;

namespace JIRASPRINTREPORT
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello and welcome to VM application!");

           // Console.Write("Username: ");
            string username = "";

            //Console.Write("Password: ");
            string password = "";

            JiraManager manager = new JiraManager(username, password);
            manager.RunQuery();

            //Console.Read();
        }
    }
}


public class JiraManager
{
    private const string m_BaseUrl = "";
    private string m_Username;
    private string m_Password;
    private string m_url;

    public JiraManager(string username, string password)
    {
        m_Username = username;
        m_Password = password;
    }

    public VMData RunQuery()
    {
        VMData vmData = new VMData();
        HttpWebRequest request = WebRequest.Create(m_BaseUrl) as HttpWebRequest;
        request.ContentType = "application/json";
        request.Method = "GET";

        string base64Credentials = GetEncodedCredentials();
        request.Headers.Add("Authorization", "Basic " + base64Credentials);

        HttpWebResponse response = request.GetResponse() as HttpWebResponse;

        string result = string.Empty;
        using (StreamReader reader = new StreamReader(response.GetResponseStream()))
        {
            result = reader.ReadToEnd();
            Sprint obj = JsonConvert.DeserializeObject<Sprint>(result);
            obj.iIssuesAddedDuringSprint = new List<IssuesAddedDuringSprint>();
            foreach (JProperty property in obj.contents.issueKeysAddedDuringSprint.Properties())
            {
                IssuesAddedDuringSprint issueadded = new IssuesAddedDuringSprint(property.Name, (bool)property.Value);
                obj.iIssuesAddedDuringSprint.Add(issueadded);
            }
            
            foreach (Issue issue in obj.contents.completedIssues)
            {
                Double initialStoryPoint = 0;
                Double finalStoryPoint = 0;
                if (issue.estimateStatistic!= null && issue.estimateStatistic.statFieldValue.value != null)
                {
                    initialStoryPoint = Double.Parse(issue.estimateStatistic.statFieldValue.value);
                }
                if (issue.currentEstimateStatistic != null && issue.currentEstimateStatistic.statFieldValue.value != null)
                {
                    finalStoryPoint = Double.Parse(issue.currentEstimateStatistic.statFieldValue.value);
                }
                if (issue.typeName == "New Feature" || issue.typeName == "Story")
                {
                    if (obj.iIssuesAddedDuringSprint.Find(x => x.key == issue.key) != null)
                    {
                        vmData.unplannedAndCompletedStories.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                    }
                    else
                    {
                        vmData.plannedAndCompletedStories.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                    }
                }
                else if (issue.typeName == "Bug")
                {
                    if (obj.iIssuesAddedDuringSprint.Find(x => x.key == issue.key) != null)
                    {
                        vmData.unplannedAndCompletedIssues.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                    }
                    else
                    {
                        vmData.plannedAndCompletedIssues.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                    }
                    vmData.completedIssues.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                }

                else if (issue.typeName == "Task")
                {
                    vmData.completedTasks.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                }
                else if (issue.typeName == "Improvement")
                {
                    vmData.completedImprovements.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                }
                else if (issue.typeName == "Technical Debt")
                {
                    vmData.completedTDs.Add(new VMIssue(issue.key, issue.summary, initialStoryPoint, finalStoryPoint));
                }
            }

            foreach (Issue issue1 in obj.contents.issuesNotCompletedInCurrentSprint)
            {
                Double initialStoryPoint = 0;
                if (issue1.estimateStatistic!= null && issue1.estimateStatistic.statFieldValue.value != null)
                {
                    initialStoryPoint = Double.Parse(issue1.estimateStatistic.statFieldValue.value);
                }
                if (issue1.typeName == "New Feature" || issue1.typeName == "Story")
                {
                    vmData.plannedAndNotCompletedStories.Add(new VMIssue(issue1.key, issue1.summary, initialStoryPoint, 0));
                }
            }
            foreach (Issue issue2 in obj.contents.puntedIssues)
            {
                Double initialStoryPoint = 0;
                if (issue2.estimateStatistic != null && issue2.estimateStatistic.statFieldValue.value != null)
                {
                    initialStoryPoint = Double.Parse(issue2.estimateStatistic.statFieldValue.value);
                }
                if (issue2.typeName == "New Feature" || issue2.typeName == "Story")
                {
                    vmData.pushedOutStories.Add(new VMIssue(issue2.key, issue2.summary, initialStoryPoint, 0));
                }
            }
            CreateExceclFile(obj, vmData);
        }
      
        return vmData;
    }

    private string GetEncodedCredentials()
    {
        string mergedCredentials = string.Format("{0}:{1}", m_Username, m_Password);
        byte[] byteCredentials = UTF8Encoding.UTF8.GetBytes(mergedCredentials);
        return Convert.ToBase64String(byteCredentials);
    }

    private void CreateExceclFile(Sprint sp, VMData vmData)
    {
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;
        object misvalue = System.Reflection.Missing.Value;
        //  try

        //Start Excel and get Application object.
        oXL = new Microsoft.Office.Interop.Excel.Application();
       // oXL.Visible = true;

        //Get a new workbook.
        oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
        oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;

        WriteSprintDetails(sp, oSheet);
        createIssuesHeader(oSheet);
        oSheet.Cells[11, 1] = "New Feature stories delivered";
        Range r = oSheet.get_Range("A11", "J11");
        r.Interior.Color = Color.LightGreen;
        r.BorderAround();
        r = null;
        int i = 0;
        for (i = 0; i < vmData.plannedAndCompletedStories.Count; ++i)
        {
            oSheet.Cells[i + 12, 2] = vmData.plannedAndCompletedStories[i].key;
            oSheet.Cells[i + 12, 3] = vmData.plannedAndCompletedStories[i].summary;
            oSheet.Cells[i + 12, 4] = vmData.plannedAndCompletedStories[i].initialStoryPoint == 0 ? 
                vmData.plannedAndCompletedStories[i].finalStoryPoint : vmData.plannedAndCompletedStories[i].initialStoryPoint;
            oSheet.Cells[i + 12, 5] = vmData.plannedAndCompletedStories[i].finalStoryPoint;
            r = oSheet.get_Range("A" + (i +12), "J"+ (i +12));
            r.Interior.Color = Color.LightGreen;
            r.BorderAround();
            r = null;
        }
        if (i != 0)
        {
            Range rng = oSheet.Range["I11"];
            rng.Formula = "=SUM(E12:E" + (i + 11);
        }

        oSheet.Cells[i + 12, 1] = "Major Improvements over past user stories delivered";
        r = oSheet.get_Range("A" + (i + 12), "J" + (i + 12));
        r.Interior.Color = Color.LightYellow;
        r.BorderAround();
        r = null;
        i++;
        int j=0;
        for (j = 0; j < vmData.completedImprovements.Count; ++j)
        {
            oSheet.Cells[j + i + 12, 2] = vmData.completedImprovements[j].key;
            oSheet.Cells[j + i + 12, 3] = vmData.completedImprovements[j].summary;
            oSheet.Cells[j + i + 12, 4] = vmData.completedImprovements[j].initialStoryPoint == 0 ?
                vmData.completedImprovements[j].finalStoryPoint : vmData.completedImprovements[j].initialStoryPoint;
            oSheet.Cells[j + i + 12, 5] = vmData.completedImprovements[j].finalStoryPoint;
            r = oSheet.get_Range("A" + (j + i + 12), "J" + (j + i + 12));
            r.Interior.Color = Color.LightYellow;
            r.BorderAround();
            r = null;
        }

        if (j != 0)
        {
            Range rng1 = oSheet.Range["I" + (i + 11)];
            rng1.Formula = "=SUM(E" + (i + 12) + ":E" + (j + i + 11);
        }

        int impTotalIndex = i + 11; ;


        int k = j + i +12;
        oSheet.Cells[k, 1] = "Miscellaneous delivered";
        r = oSheet.get_Range("A" + (k ), "J" + (k));
        r.Interior.Color = Color.LimeGreen;
        r.BorderAround();
        r = null;
        k++;
        int l = 0;
        for (l = 0; l < vmData.completedTasks.Count; ++l)
        {
            oSheet.Cells[k + l, 2] = vmData.completedTasks[l].key;
            oSheet.Cells[k + l, 3] = vmData.completedTasks[l].summary;
            oSheet.Cells[k + l, 4] = vmData.completedTasks[l].initialStoryPoint == 0 ?
                vmData.completedTasks[l].finalStoryPoint : vmData.completedTasks[l].initialStoryPoint;
            oSheet.Cells[k + l, 5] = vmData.completedTasks[l].finalStoryPoint;
            r = oSheet.get_Range("A" + (k+l), "J" + (k +l));
            r.Interior.Color = Color.LimeGreen;
            r.BorderAround();
            r = null;
        }

        if (l != 0)
        {
            Range rng2 = oSheet.Range["I" + (k - 1)];
            rng2.Formula = "=SUM(E" + (k) + ":E" + (k + l - 1);
        }

        k += l;
        oSheet.Cells[k, 1] = "Technical Debt delivered";
        r = oSheet.get_Range("A" + (k), "J" + (k));
        r.Interior.Color = Color.LightSkyBlue;
        r.BorderAround();
        r = null;
        k++;
        l = 0;
        for (l = 0; l < vmData.completedTDs.Count; ++l)
        {
            oSheet.Cells[k + l, 2] = vmData.completedTDs[l].key;
            oSheet.Cells[k + l, 3] = vmData.completedTDs[l].summary;
            oSheet.Cells[k + l, 4] = vmData.completedTDs[l].initialStoryPoint == 0 ?
                vmData.completedTDs[l].finalStoryPoint : vmData.completedTDs[l].initialStoryPoint;
            oSheet.Cells[k + l, 5] = vmData.completedTDs[l].finalStoryPoint;
            r = oSheet.get_Range("A" + (k + l), "J" + (k + l));
            r.Interior.Color = Color.LightSkyBlue;
            r.BorderAround();
            r = null;
        }
        if (l != 0)
        {
            Range rng3 = oSheet.Range["I" + (k - 1)];
            rng3.Formula = "=SUM(E" + (k) + ":E" + (k + l - 1);
        }
        int tdTotalIndex = k - 1;
        k += l;
        int indexForScopeCreepCal = k - 1;
        oSheet.Cells[k, 1] = "Unplanned Stories committed, and delivered";
        r = oSheet.get_Range("A" + (k), "J" + (k));
        r.Interior.Color = Color.LightYellow;
        r.BorderAround();
        r = null;
        k++;
        l = 0;
        for (l = 0; l < vmData.unplannedAndCompletedStories.Count; ++l)
        {
            oSheet.Cells[k + l, 2] = vmData.unplannedAndCompletedStories[l].key;
            oSheet.Cells[k + l, 3] = vmData.unplannedAndCompletedStories[l].summary;
            oSheet.Cells[k + l, 4] = vmData.unplannedAndCompletedStories[l].initialStoryPoint == 0 ?
                vmData.unplannedAndCompletedStories[l].finalStoryPoint : vmData.unplannedAndCompletedStories[l].initialStoryPoint;
            oSheet.Cells[k + l, 5] = vmData.unplannedAndCompletedStories[l].finalStoryPoint;
            r = oSheet.get_Range("A" + (k + l), "J" + (k + l));
            r.Interior.Color = Color.LightYellow;
            r.BorderAround();
            r = null;
        }

        if (l != 0)
        {
            Range rng4 = oSheet.Range["I" + (k - 1)];
            rng4.Formula = "=SUM(E" + (k) + ":E" + (k + l - 1);
        }

        int totalUnplannedSPindex = k-1;

        k += l;
        oSheet.Cells[k, 1] = "Past Bugs fixed";
        r = oSheet.get_Range("A" + (k), "J" + (k));
        r.Interior.Color = Color.LightGreen;
        r.BorderAround();
        r = null;
        k++;
        l = 0;
        for (l = 0; l < vmData.completedIssues.Count; ++l)
        {
            oSheet.Cells[k + l, 2] = vmData.completedIssues[l].key;
            oSheet.Cells[k + l, 3] = vmData.completedIssues[l].summary;
            //oSheet.Cells[k + l, 4] = vmData.unplannedAndCompletedStories[l].initialStoryPoint == 0 ?
            //    vmData.unplannedAndCompletedStories[l].finalStoryPoint : vmData.unplannedAndCompletedStories[l].initialStoryPoint;
            oSheet.Cells[k + l, 7] = vmData.completedIssues[l].finalStoryPoint;
            r = oSheet.get_Range("A" + (k + l), "J" + (k + l));
            r.Interior.Color = Color.LightGreen;
            r.BorderAround();
            r = null;
        }

        if (l != 0)
        {
            Range rng5 = oSheet.Range["I" + (k - 1)];
            rng5.Formula = "=SUM(G" + (k) + ":G" + (k + l - 1);
        }
        int bugstotalIndex = k - 1;
        k += l;
        oSheet.Cells[k, 1] = "Other work committed, and delivered";
        r = oSheet.get_Range("A" + (k), "J" + (k));
        r.Interior.Color = Color.LightYellow;
        r.BorderAround();
        r = null;
        k++;
        l = 0;

        k += l;
        oSheet.Cells[k, 1] = "Stories committed, but Pushed out of sprint";
        r = oSheet.get_Range("A" + (k), "J" + (k));
        r.Interior.Color = Color.LightPink;
        r.BorderAround();
        r = null;
        k++;
        l = 0;
        for (l = 0; l < vmData.plannedAndNotCompletedStories.Count; ++l)
        {
            oSheet.Cells[k + l, 2] = vmData.plannedAndNotCompletedStories[l].key;
            oSheet.Cells[k + l, 3] = vmData.plannedAndNotCompletedStories[l].summary;
            oSheet.Cells[k + l, 4] = vmData.plannedAndNotCompletedStories[l].initialStoryPoint == 0 ?
                vmData.plannedAndNotCompletedStories[l].finalStoryPoint : vmData.plannedAndNotCompletedStories[l].initialStoryPoint;
            oSheet.Cells[k + l, 5] = vmData.plannedAndNotCompletedStories[l].finalStoryPoint;
            r = oSheet.get_Range("A" + (k + l), "J" + (k + l));
            r.Interior.Color = Color.LightPink;
            r.BorderAround();
            r = null;
        }

        if (l != 0)
        {
            Range rng6 = oSheet.Range["I" + (k - 1)];
            rng6.Formula = "=SUM(D" + (k) + ":D" + (k + l - 1);
        }

        int storiesPushedOutSPTotalIndex = k -1;

        k += l;
        oSheet.Cells[k, 1] = "Stories committed, but Undelivered";
        r = oSheet.get_Range("A" + (k), "J" + (k));
        r.Interior.Color = Color.Pink;
        r.Font.Color = Color.Red;
        r.BorderAround();
        r = null;
        k++;
        l = 0;
        for (l = 0; l < vmData.pushedOutStories.Count; ++l)
        {
            oSheet.Cells[k + l, 2] = vmData.pushedOutStories[l].key;
            oSheet.Cells[k + l, 3] = vmData.pushedOutStories[l].summary;
            oSheet.Cells[k + l, 4] = vmData.pushedOutStories[l].initialStoryPoint == 0 ?
                vmData.pushedOutStories[l].finalStoryPoint : vmData.pushedOutStories[l].initialStoryPoint;
            oSheet.Cells[k + l, 5] = vmData.pushedOutStories[l].finalStoryPoint;
            r = oSheet.get_Range("A" + (k + l), "J" + (k + l));
            r.Interior.Color = Color.Pink;
            r.BorderAround();
            r = null;
        }

        if (l != 0)
        {
            Range rng7 = oSheet.Range["I" + (k - 1)];
            rng7.Formula = "=SUM(D" + (k) + ":D" + (k + l - 1);
        }

        int storiesNotDelTotalIndex = k-1;

        k += l;
        int totalRowIndex = k;
        Range rTotal = oSheet.get_Range("A" +k, "J" +k);
        rTotal.Interior.Color = Color.LightGray;

        oSheet.Cells[k, 1] = "Total";

        Range rng8 = oSheet.Range["D" + (k)];
        rng8.Formula = "=SUM(D11:D" + (k - 1);

        Range rng9 = oSheet.Range["E" + (k)];
        rng9.Formula = "=SUM(E11:E" + (k - 1);

        Range rng10 = oSheet.Range["F" + (k)];
        rng10.Formula = "=SUM(F11:F" + (k - 1);

        Range rng11 = oSheet.Range["G" + (k)];
        rng11.Formula = "=SUM(G11:G" + (k - 1);


        k++;
        oSheet.Cells[k, 1] = "VALUE MEASURES";
        Range rVM = oSheet.Cells[k, 1];
        rVM.Font.Color = Color.Blue;

        k++;
        oSheet.Cells[k, 1] = "Value Measure";
        oSheet.Cells[k, 2] = "Final Result";
        oSheet.Cells[k, 3] = "Formula with Data";
        oSheet.Cells[k, 4] = "Formula";
        oSheet.Cells[k, 5] = "Comment";
        Range rVM1 = oSheet.get_Range("A" + k, "J"+k);
        rVM1.Interior.Color = Color.LightGray;
        k++;
        int vmSummaryStartIndex = k;
        oSheet.Cells[k, 1] = "Sprint Delivery Ratio";
        Range rng12 = oSheet.Range["B" + (k)];
        rng12.Formula = "=E" + totalRowIndex + "/D" + totalRowIndex;
        Range rng13 = oSheet.Range["E" + (totalRowIndex)];
        Range rng14 = oSheet.Range["D" + (totalRowIndex)];
        oSheet.Cells[k, 3] = rng13.Value + "/" + rng14.Value;
        oSheet.Cells[k, 4] = "Total story points completed/Total planned story points";
        oSheet.Cells[k, 5] = "excluding bugs, pushed out and undelivered stories";

        k++;
        oSheet.Cells[k, 1] = "Velocity (story points/day)";
        Range rng15 = oSheet.Range["B" + (k)];
        rng15.Formula = "=E" + totalRowIndex + "/B6";
        Range rng16 = oSheet.Range["E" + (totalRowIndex)];
        Range rng17 = oSheet.Range["B6"];
        oSheet.Cells[k, 3] = rng16.Value + "/" + rng17.Value;
        oSheet.Cells[k, 4] = "Number of story points completed / Sprint Capacity available in Man Days";
        oSheet.Cells[k, 5] = "excluding bugs, pushed out and undelivered stories";

        k++;
        oSheet.Cells[k, 1] = "First Time Right Ratio";
        Range rng18 = oSheet.Range["B" + (k)];
        rng18.Formula = "=1-(G" + totalRowIndex + "/E" + totalRowIndex + ")";
        Range rng19 = oSheet.Range["G" + (totalRowIndex)];
        Range rng20 = oSheet.Range["E" + totalRowIndex];
        oSheet.Cells[k, 3] = "1-(" + rng19.Value + "/" + rng20.Value + ")";
        oSheet.Cells[k, 4] = "1 - (Story points of bugs fixed / Total story points of completed stories)";

        k++;
        oSheet.Cells[k, 1] = "Scope Creep Ratio (changes to planned stories)";
        Range rng21 = oSheet.Range["B" + (k)];
        rng21.Formula = "=(SUM(E11:E" + indexForScopeCreepCal + ")-SUM(D11:D" + indexForScopeCreepCal + "))/D" + totalRowIndex;
        oSheet.Cells[k, 4] = "Story points added & delivered / Story points planned in Sprint";

        k++;
        oSheet.Cells[k, 1] = "Agility for Unplanned Tasks";
        Range rng22 = oSheet.Range["B" + (k)];
        rng22.Formula = "=(I"+totalUnplannedSPindex+"-I"+storiesPushedOutSPTotalIndex+"-I" + storiesNotDelTotalIndex + ")/(D" + totalRowIndex + ")";
        oSheet.Cells[k, 4] = "(Unplanned Story points - Story points pushed out - SP undelivered) / Total Planned Story points in the Sprint";

        k++;
        oSheet.Cells[k, 1] = "Technical Debt Work Ratio";
        Range rng23 = oSheet.Range["B" + (k)];
        rng23.Formula = "=I" + tdTotalIndex + "/E" + totalRowIndex;
        oSheet.Cells[k, 4] = "Total debt story points completed / Total story points completed";
        oSheet.Cells[k, 5] = "code refactoring, design changes, performance issue fix for product improvement";

        
        Range vmSummary = oSheet.get_Range("A" + vmSummaryStartIndex, "A" + k);
        vmSummary.Interior.Color = Color.LightGray;
        vmSummary.BorderAround();

        Range rng24 = oSheet.Range["J11"];
        rng24.Formula = "=I11/E" + totalRowIndex;
        rng24.NumberFormat = "###,##%";
        oSheet.Cells[11, 11] = "% of new features against delivered work";   


        Range rng25 = oSheet.Range["J" + impTotalIndex];
        rng25.Formula = "=I"+impTotalIndex + "/E" + totalRowIndex;
        rng25.NumberFormat = "###,##%";
        oSheet.Cells[impTotalIndex, 11] = "% of feature improvements against delivered work";   


        Range rng26 = oSheet.Range["J" + tdTotalIndex];
        rng26.Formula = "=I" + tdTotalIndex + "/E" + totalRowIndex;
        rng26.NumberFormat = "###,##%";
        oSheet.Cells[tdTotalIndex, 11] = "% of tech debt work against delivered work";   


        Range rng27 = oSheet.Range["J" + totalUnplannedSPindex];
        rng27.Formula = "=I" + totalUnplannedSPindex + "/E" + totalRowIndex;
        rng27.NumberFormat = "###,##%";
        oSheet.Cells[totalUnplannedSPindex, 11] = "% of unplanned work against delivered work";   


        Range rng28 = oSheet.Range["J" + bugstotalIndex];
        rng28.Formula = "=I" + bugstotalIndex + "/E" + totalRowIndex;
        rng28.NumberFormat = "###,##%";
        oSheet.Cells[bugstotalIndex, 11] = "% of past bug fixing work against delivered work";   

        Range rng29 = oSheet.Range["J" + storiesPushedOutSPTotalIndex];
        rng29.Formula = "=I" + storiesPushedOutSPTotalIndex + "/E" + totalRowIndex;
        rng29.NumberFormat = "###,##%";
        oSheet.Cells[storiesPushedOutSPTotalIndex, 11] = "% of work pushed out against committed work";   

        Range rng30 = oSheet.Range["J" + storiesNotDelTotalIndex];
        rng30.Formula = "=I" + storiesNotDelTotalIndex + "/E" + totalRowIndex;
        rng30.NumberFormat = "###,##%";
        oSheet.Cells[storiesNotDelTotalIndex, 11] = "% of work undelivered against committed work";        


            SetColumnsProperties(oSheet);
        oXL.Visible = false;
        oXL.UserControl = false;
        oWB.SaveAs("c:\\test\\test505.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
            false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlShared,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

        oWB.Close();
    }

    private static void SetColumnsProperties(Microsoft.Office.Interop.Excel._Worksheet oSheet)
    {
        oSheet.Columns.Font.Name = "Calibri";
        oSheet.Columns.Font.Size = 10;
        oSheet.Columns.AutoFit();
        Range rc = oSheet.Columns[3];
        rc.ColumnWidth = 35;
        rc.WrapText = true;
        rc.BorderAround();
        rc = null;
        rc = oSheet.Columns[1];
        rc.BorderAround();
        rc.ColumnWidth = 31;
        rc.WrapText = true;
        rc = null;
        rc = oSheet.Columns[2];
        rc.BorderAround();
        rc.ColumnWidth = 15;
        rc = null;
        rc = oSheet.Columns[4];
        rc.ColumnWidth = 20;
        rc.WrapText = true;
        rc.BorderAround();
        rc = null;
        rc = oSheet.Columns[5];
        rc.ColumnWidth = 12;
        rc.WrapText = true;
        rc.Font.Color = Color.Blue;
        rc.Font.Italic = true;
        rc.BorderAround();
        rc = null;
        rc = oSheet.Columns[6];
        rc.BorderAround();
        rc.ColumnWidth = 12;
        rc = null;
        rc = oSheet.Columns[7];
        rc.ColumnWidth = 12;
        rc.BorderAround();
        rc = null;
        rc = oSheet.Columns[8];
        rc.ColumnWidth = 12;
        rc.BorderAround();
        rc = null;
        rc = oSheet.Columns[9];
        rc.ColumnWidth = 12;
        rc.BorderAround();
        rc = null;
        rc = oSheet.Columns[10];
        rc.ColumnWidth = 12;
        
        rc.BorderAround();
        rc = null;
    }

    private static void WriteSprintDetails(Sprint sp, Microsoft.Office.Interop.Excel._Worksheet oSheet)
    {
        oSheet.Cells[1, 1] = "SPRINT DETAILS";
        Range r = oSheet.Cells[1, 1];
        r.Font.Color = Color.Blue;
        oSheet.Cells[1, 5] = "Additional Notes";
        oSheet.Cells[2, 1] = "Sprint Name";
        oSheet.Cells[2, 2] = sp.sprint.name;

        oSheet.Cells[3, 1] = "Sprint Start Date";
        oSheet.Cells[3, 2] = sp.sprint.startDate;

        oSheet.Cells[4, 1] = "Sprint End Date";
        oSheet.Cells[4, 2] = sp.sprint.endDate;

        oSheet.Cells[5, 1] = "Agile Methodology";
        oSheet.Cells[5, 2] = "Scrum";

        oSheet.Cells[6, 1] = "Sprint Capacity (days)";
        oSheet.Cells[6, 2] = 66;

        oSheet.Cells[9, 1] = "SPRINT REPORT";
        r = oSheet.Cells[9, 1];
        r.Font.Color = Color.Blue;
        oSheet.Cells[9, 4] = "Delivered";
        oSheet.Cells[9, 6] = " Internal/Past";
    }

    private static void createIssuesHeader(Microsoft.Office.Interop.Excel._Worksheet oSheet)
    {
        oSheet.Cells[10, 1] = "User Story Details";
        oSheet.Cells[10, 4] = "Original Story Point Esimate";
        oSheet.Cells[10, 5] = "Final Story Point Estimate";
        oSheet.Cells[10, 6] = "Bugs in hrs (optional)";
        oSheet.Cells[10, 7] = "Bugs in SP";
        oSheet.Cells[10, 8] = "Status";
        oSheet.Cells[10, 9] = "Total SPs";
        oSheet.Cells[10, 10] = "Percentage";
        Range r = oSheet.get_Range("A10", "J10");
        r.Interior.Color = Color.LightGray;
    }
}