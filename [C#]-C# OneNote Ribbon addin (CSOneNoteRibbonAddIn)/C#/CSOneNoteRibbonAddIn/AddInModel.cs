using System.Collections.Generic;


namespace CSOneNoteRibbonAddIn
{
    public class AddInModel
    {
        public string NotebookId { get; set; } = "";
        public string NotebookName { get; set; } = "";
        public string NotebookColor { get; set; } = "";
        public SectionGroupModel SectionGroup { get; set; }
        public SectionModel Section { get; set; }
        public PageModel Page { get; set; }
    }

    public class SectionGroupModel
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string Color { get; set; } = "";
        public List<SectionModel> Sections { get; set; } = new List<SectionModel>();
    }

    public class SectionModel
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public string Color { get; set; } = "";
        public List<PageModel> Pages { get; set; } = new List<PageModel>();
    }

    public class PageModel
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";
        public List<ParagraphModel> Paragraphs { get; set; } = new List<ParagraphModel>();
    }

    public class ParagraphModel
    {
        public string Id { get; set; } = "";
        public string Name { get; set; } = "";  // paragraph content (text)
    }
}
