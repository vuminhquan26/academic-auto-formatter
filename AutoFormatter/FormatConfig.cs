using System.Collections.Generic;

namespace AutoFormatter
{
    public class FormatConfig
    {
        public PageSetupConfig PageSetup { get; set; } = new PageSetupConfig();
        public TypographyConfig Typography { get; set; } = new TypographyConfig();
        public AutomationConfig Automation { get; set; } = new AutomationConfig();
        public TocConfig Toc { get; set; } = new TocConfig();
        public CoverGeneratorConfig Cover { get; set; } = new CoverGeneratorConfig();
    }

    public class PageSetupConfig
    {
        public double LeftCm { get; set; } = 3.5;
        public double RightCm { get; set; } = 2.0;
        public double TopCm { get; set; } = 2.0;
        public double BottomCm { get; set; } = 2.0;
        public bool IsTwoSidedPrint { get; set; } = true;
    }

    public class TypographyConfig
    {
        public string FontName { get; set; } = "Times New Roman";
        public int FontSize { get; set; } = 13;
        public bool EnforceLineSpacing { get; set; } = true;
        public double LineSpacingValue { get; set; } = 1.5;
        public bool AutoFormatHeadings { get; set; } = true;
    }

    public class AutomationConfig
    {
        public bool FixTableAndFigureWidth { get; set; } = true;
        public bool FixEquations { get; set; } = true;
        public bool FixReferenceIndentation { get; set; } = true;
    }

    public class TocConfig
    {
        public bool GenerateTOC { get; set; } = true;
        public bool GenerateTOF { get; set; } = true;
        public bool GenerateTOT { get; set; } = true;
    }

    public class CoverGeneratorConfig
    {
        public bool GenerateCover { get; set; } = true;
        public bool AddPageBorder { get; set; } = true;
        public string UniversityName { get; set; } = "";
        public string FacultyName { get; set; } = "";
        public string LogoPath { get; set; } = "";
        public string ReportType { get; set; } = "BÁO CÁO BÀI TẬP LỚN";
        public string TopicName { get; set; } = "";
        public string InstructorName { get; set; } = "";

        public bool IsGroupProject { get; set; } = false;
        public string GroupName { get; set; } = "";
        public List<CustomField> CommonFields { get; set; } = new List<CustomField>();
        public List<StudentInfo> Students { get; set; } = new List<StudentInfo>();
        public string LocationAndDate { get; set; } = "";
    }

    public class CustomField
    {
        public string Label { get; set; } = "";
        public string Value { get; set; } = "";
    }

    public class StudentInfo
    {
        public string Name { get; set; } = "";
        public string StudentId { get; set; } = "";
    }
}