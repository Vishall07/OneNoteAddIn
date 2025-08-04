using System;
using System.Windows.Forms;

namespace CSOneNoteRibbonAddIn
{
    public class BookMark_Window : Form
    {
        private Label label;
        private TreeView treeView;

        public BookMark_Window()
        {
            this.Text = "Notebook Details";
            this.Width = 600;
            this.Height = 400;
            this.TopMost = true;

            label = new Label();
            label.Location = new System.Drawing.Point(30, 20);
            label.AutoSize = true;

            treeView = new TreeView();
            treeView.Location = new System.Drawing.Point(30, 50);
            treeView.Width = 520;
            treeView.Height = 300;
            treeView.NodeMouseDoubleClick += TreeView_NodeMouseDoubleClick;

            this.Controls.Add(label);
            this.Controls.Add(treeView);

            this.Activated += BookMark_Window_Activated;
        }

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);
            LoadNotebookHierarchy();
            this.Activate(); // Force focus on load
        }

        private void BookMark_Window_Activated(object sender, EventArgs e)
        {
            this.TopMost = true; // Re-assert TopMost when window activates
        }

        private void LoadNotebookHierarchy()
        {
            treeView.Nodes.Clear();
            try
            {
                var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
                string currentNotebookId = oneNoteApp.Windows.CurrentWindow.CurrentNotebookId;

                string xml;
                oneNoteApp.GetHierarchy(currentNotebookId,
                    Microsoft.Office.Interop.OneNote.HierarchyScope.hsPages, out xml);

                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(xml);
                var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                var notebook = doc.SelectSingleNode("//one:Notebook", nsmgr);
                if (notebook != null)
                {
                    string notebookName = notebook.Attributes["name"]?.Value ?? "Notebook";
                    TreeNode notebookNode = new TreeNode(notebookName);
                    notebookNode.Tag = notebook.Attributes["ID"]?.Value ?? "";
                    treeView.Nodes.Add(notebookNode);

                    var sections = notebook.SelectNodes("one:Section", nsmgr);
                    foreach (System.Xml.XmlNode section in sections)
                    {
                        string sectionName = section.Attributes["name"]?.Value ?? "Section";
                        TreeNode sectionNode = new TreeNode(sectionName);
                        sectionNode.Tag = section.Attributes["ID"]?.Value ?? "";
                        notebookNode.Nodes.Add(sectionNode);

                        var pages = section.SelectNodes("one:Page", nsmgr);
                        foreach (System.Xml.XmlNode page in pages)
                        {
                            string pageName = page.Attributes["name"]?.Value ?? "Page";
                            TreeNode pageNode = new TreeNode(pageName);
                            pageNode.Tag = page.Attributes["ID"]?.Value ?? "";
                            sectionNode.Nodes.Add(pageNode);
                        }
                    }
                    notebookNode.Expand();
                }
                label.Text = "Loaded structure.";
            }
            catch (Exception ex)
            {
                label.Text = "Load error: " + ex.Message;
            }
        }

        private void TreeView_NodeMouseDoubleClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            try
            {
                var id = e.Node.Tag?.ToString();
                if (!string.IsNullOrEmpty(id))
                {
                    var oneNoteApp = new Microsoft.Office.Interop.OneNote.Application();
                    oneNoteApp.NavigateTo(id);
                    label.Text = "Opened: " + e.Node.Text;
                }
            }
            catch (Exception ex)
            {
                label.Text = "Open error: " + ex.Message;
            }
        }
    }
}
