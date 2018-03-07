using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace JIRASPRINTREPORT
{
    public class VMIssue
    {
        public string key { get; set; }
        public string summary { get; set; }
        public string typeName { get; set; }
        public string statusName { get; set; }
        public Double initialStoryPoint{ get; set; }
        public Double finalStoryPoint { get; set; }
        public VMIssue(string _Key, string _summary, Double _initialStoryPoint, Double _finalStoryPoint)
        {
            key = _Key;
            summary = _summary;
            initialStoryPoint = _initialStoryPoint;
            finalStoryPoint = _finalStoryPoint;
        }
    }
    public class VMData
    {
        public List<VMIssue> plannedAndCompletedStories;
        public List<VMIssue> plannedAndPushedStories;
        public List<VMIssue> plannedAndNotCompletedStories;
        public List<VMIssue> unplannedAndCompletedStories;
        public List<VMIssue> plannedAndCompletedIssues;
        public List<VMIssue> unplannedAndCompletedIssues;
        public List<VMIssue> completedIssues;
        public List<VMIssue> completedTasks;
        public List<VMIssue> completedImprovements;
        public List<VMIssue> completedTDs;
        public List<VMIssue> pushedOutStories;



        public VMData()
        {
            plannedAndCompletedStories = new List<VMIssue>();
            plannedAndPushedStories = new List<VMIssue>();
            plannedAndNotCompletedStories = new List<VMIssue>();
            unplannedAndCompletedStories = new List<VMIssue>();
            plannedAndCompletedIssues = new List<VMIssue>();
            unplannedAndCompletedIssues = new List<VMIssue>();
            completedIssues = new List<VMIssue>();
            completedTasks = new List<VMIssue>();
            completedImprovements = new List<VMIssue>();
            completedTDs = new List<VMIssue>();
            pushedOutStories = new List<VMIssue>();


        }
    }
}
