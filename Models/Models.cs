using System.Collections.Generic;

namespace BtTestViewer.Models
{
    public class TestSummary
    {
        public string DeviceType { get; set; } = "";
        public string DeviceNo { get; set; } = "";
        public string TestPlan { get; set; } = "";
        public string User { get; set; } = "";
        public string Instrument { get; set; } = "";
        public string StartTime { get; set; } = "";
        public string EndTime { get; set; } = "";
        public string TotalTime { get; set; } = "";
        public int ItemsPassed { get; set; }
        public int ItemsFailed { get; set; }
        public int ItemsTotal { get; set; }
        public string FinalVerdict { get; set; } = "";
    }

    public class TestGroup
    {
        public string TpId { get; set; } = "";       // e.g. "BV-01-C"
        public string TpName { get; set; } = "";     // e.g. "Output Power"
        public string FullContext { get; set; } = "";
        public List<TestItem> Items { get; set; } = new();

        public int PassedCount => Items.Count(i => i.Verdict == "Passed");
        public int FailedCount => Items.Count(i => i.Verdict == "Failed");
        public int TotalCount => Items.Count(i => i.Verdict == "Passed" || i.Verdict == "Failed");
        public string GroupVerdict => FailedCount > 0 ? "Failed" : (TotalCount > 0 ? "Passed" : "Info");
    }

    public class TestItem
    {
        public int No { get; set; }
        public string Description { get; set; } = "";
        public string Condition { get; set; } = "";
        public string LowLimit { get; set; } = "";
        public string HighLimit { get; set; } = "";
        public string MeasValue { get; set; } = "";
        public string Unit { get; set; } = "";
        public string VerdictInfo { get; set; } = "";
        public string Verdict { get; set; } = "";
        public string GroupName { get; set; } = "";
    }

    public class ParsedResult
    {
        public TestSummary Summary { get; set; } = new();
        public List<TestGroup> Groups { get; set; } = new();
        public List<TestItem> AllItems { get; set; } = new();
    }
}
