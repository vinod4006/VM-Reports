using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace JIRASPRINTREPORT
{
    public class EpicField
    {
        public string id { get; set; }
        public string label { get; set; }
        public bool editable { get; set; }
        public string renderer { get; set; }
        public string epicKey { get; set; }
        public string epicColor { get; set; }
        public string text { get; set; }
        public bool canRemoveEpic { get; set; }
    }

    public class StatFieldValue
    {
        public string value { get; set; }
    }

    public class CurrentEstimateStatistic
    {
        public string statFieldId { get; set; }
        public StatFieldValue statFieldValue { get; set; }
    }

    public class StatFieldValue2
    {
        public string value { get; set; }
    }

    public class EstimateStatistic
    {
        public string statFieldId { get; set; }
        public StatFieldValue2 statFieldValue { get; set; }
    }

    public class StatusCategory
    {
        public string id { get; set; }
        public string key { get; set; }
        public string colorName { get; set; }
    }

    public class Status
    {
        public string id { get; set; }
        public string name { get; set; }
        public string description { get; set; }
        public string iconUrl { get; set; }
        public StatusCategory statusCategory { get; set; }
    }

    public class Issue
    {
        public int id { get; set; }
        public string key { get; set; }
        public bool hidden { get; set; }
        public string typeName { get; set; }
        public string typeId { get; set; }
        public string summary { get; set; }
        public string typeUrl { get; set; }
        public string priorityUrl { get; set; }
        public string priorityName { get; set; }
        public bool done { get; set; }
        public string assignee { get; set; }
        public string assigneeKey { get; set; }
        public string assigneeName { get; set; }
        public string avatarUrl { get; set; }
        public bool hasCustomUserAvatar { get; set; }
        public string color { get; set; }
        public bool flagged { get; set; }
        public string epic { get; set; }
        public EpicField epicField { get; set; }
        public CurrentEstimateStatistic currentEstimateStatistic { get; set; }
        public bool estimateStatisticRequired { get; set; }
        public EstimateStatistic estimateStatistic { get; set; }
        public string statusId { get; set; }
        public string statusName { get; set; }
        public string statusUrl { get; set; }
        public Status status { get; set; }
        public List<int> fixVersions { get; set; }
        public int projectId { get; set; }
        public int linkedPagesCount { get; set; }
    }
    public class IssuesAddedDuringSprint
    {
        public string key { get; set; }
        public bool val { get; set; }
        public IssuesAddedDuringSprint(string _key, bool _val)
        {
            key = _key;
            val = _val;
        }
    }

    public class SprintIssues
    {
        public List<Issue> completedIssues;
        public List<Issue> issuesNotCompletedInCurrentSprint;
        public List<Issue> puntedIssues;
        public JObject issueKeysAddedDuringSprint;
    }

    public class SprintDetails
    {
        public int id { get; set; }
        public int sequence { get; set; }
        public string name { get; set; }
        public string state { get; set; }
        public int linkedPagesCount { get; set; }
        public string goal { get; set; }
        public string startDate { get; set; }
        public string endDate { get; set; }
        public string completeDate { get; set; }
        public bool canUpdateSprint { get; set; }
        public List<object> remoteLinks { get; set; }
        public int daysRemaining { get; set; }
    }

    public class Sprint
    {
        public SprintIssues contents;
        public SprintDetails sprint;
        public List<IssuesAddedDuringSprint> iIssuesAddedDuringSprint;
    }
}
