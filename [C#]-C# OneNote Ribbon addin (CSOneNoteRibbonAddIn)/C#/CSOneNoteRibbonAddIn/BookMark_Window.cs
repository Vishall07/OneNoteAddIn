#region NameSpaces
using Microsoft.Office.Interop.OneNote;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices; 
using System.Text;
using System.Web;
using System.Windows.Forms;
using System.Xml;
using OneNote = Microsoft.Office.Interop.OneNote;


#endregion

namespace CSOneNoteRibbonAddIn
{
    public class BookMark_Window : Form
    {
        #region Properties
        private ContextMenuStrip columnHeaderContextMenu;
        private DataGridViewColumn clickedColumnHeader;
        private Label label;
        private ComboBox comboScope;
        private Button btnSave;
        private Button btnDelete;
        private DataGridView grid;
        public string selectedId;
        private string selectedScope;
        private string selectedText;
        private string tablePath;
        private string notebookName, notebookColor, sectionGroupName, sectionName, sectionColor, pageName, paraContent;
        private Label labelNotebook, labelSection, labelPage, labelPara;
        private const int ResizeBorder = 6;
        private List<BookmarkItem> items = new List<BookmarkItem>();
        private Point dragStart;
        private ContextMenuStrip contextMenu;
        private bool sortAscending = true;
        private List<BookmarkItem> cachedList;
        private bool showingAlphabetical = false;
        private bool isTextWrapEnabled = false;
        private ListBox listScope;
        private Point mouseOffset;
        private bool isDragging = false;
        [DllImport("user32.dll")]
        public static extern bool ReleaseCapture();
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        private const int WM_NCLBUTTONDOWN = 0xA1;
        private const int HTCAPTION = 0x2;
        #endregion

        #region Initialization and Grid Building
        public BookMark_Window(
            string onenoteId,
            string onenoteScope,
            string displayText,
            string notebookName,
            string notebookColor,
            string sectionGroupName,
            string sectionName,
            string sectionColor,
            string pageName,
            string paraContent)
        {
            try
            {
                // Basic window setup
                this.FormBorderStyle = FormBorderStyle.None;
                //this.Width = 600;
                //this.Height = 300;
                
                this.BackColor = Color.White;

                labelNotebook = new Label { Location = new Point(20, 30), AutoSize = true };
                labelSection = new Label { Location = new Point(20, 70), AutoSize = true };
                labelPage = new Label { Location = new Point(20, 90), AutoSize = true };
                labelPara = new Label { Location = new Point(20, 110), AutoSize = true };

                listScope = new ListBox()
                {
                    Location = new Point(20, 12),
                    Width = 140,
                    Font = new Font("Segoe UI", 10),
                    Height = 90 // Adjust as needed
                };

                // Add your scope options
                listScope.Items.AddRange(new[]
                {
                    "Current Notebook",
                    "Current Section Group",
                    "Current Section",
                    "Current Page",
                    "Current Paragraph"

                });

                listScope.Click += ListScope_Click; 

                grid = new DataGridView
                {
                    Location = new Point(20, 120),
                    Width = this.ClientSize.Width - 40,
                    Height = this.ClientSize.Height - 10,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    AllowDrop = true,
                    AllowUserToResizeColumns = true,
                    AllowUserToOrderColumns = true,
                    RowHeadersVisible = false,
                    Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
                };
                grid.MouseDown += Grid_MouseDown;
                grid.MouseDown += Grid_MouseDown_StartDrag;
                grid.MouseMove += Grid_MouseMove;
                grid.DragOver += Grid_DragOver;
                grid.DragDrop += Grid_DragDrop;
                grid.CellValueChanged += grid_CellValueChanged;
                grid.CellDoubleClick += Grid_CellDoubleClick;
                grid.ColumnHeaderMouseClick += grid_ColumnHeaderMouseClick;
                grid.KeyDown += Grid_KeyDown;
                grid.CellClick += Grid_CellClick;

                contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("New Folder", null, NewFolder_Click);
                contextMenu.Items.Add("Rename", null, Rename_Click);
                contextMenu.Items.Add("Delete", null, Delete_Click);
                contextMenu.Items.Add("Add URL Bookmark", null, AddUrlBookmark_Click);
                contextMenu.Items.Add("TextWrap On/Off", null, TextWrap_Click);
                contextMenu.Items.Add("Text Wrap Current Row", null, TextWrapCurrentRow_Click);
                contextMenu.Items.Add("Open All Notes", null, Open_All_Notes); 
                contextMenu.Items.Add("Export All Bookmarks", null, Export_All_Bookmarks_Click);
                contextMenu.Items.Add("Settings", null, Settings_Click);
                contextMenu.Items.Add("Show Method Time Logs", null, ShowMethodLogs_Click);

                //Controls.Add(btnDelete);
                Controls.Add(grid);
                Controls.Add(listScope);

                columnHeaderContextMenu = new ContextMenuStrip();
                var textWrapMenuItem = new ToolStripMenuItem("Text Wrap This Column");
                textWrapMenuItem.Click += TextWrapMenuItem_Click;
                columnHeaderContextMenu.Items.Add(textWrapMenuItem);

                selectedId = onenoteId;
                selectedScope = onenoteScope;
                selectedText = displayText;
                tablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bookmarks.txt");

                this.notebookName = notebookName;
                this.notebookColor = notebookColor;
                this.sectionGroupName = sectionGroupName;
                this.sectionName = sectionName;
                this.sectionColor = sectionColor;
                this.pageName = pageName;
                this.paraContent = paraContent;
                this.Font = new Font("Segoe UI", 10);
                this.BackColor = ColorTranslator.FromHtml("#f3f3f3");
                this.MouseDown += Form_MouseDown;
                LoadTable();
                UpdateBookmarkInfo(selectedId, selectedScope, selectedText, notebookName, notebookColor,
                    sectionGroupName, sectionName, sectionColor, pageName, paraContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing Bookmark window: " + ex.Message);
            }
        }
        public void UpdateBookmarkInfo(
            string newSelectedId,
            string newSelectedScope,
            string newSelectedText,
            string newNotebookName,
            string newNotebookColor,
            string newSectionGroupName,
            string newSectionName,
            string newSectionColor,
            string newPageName,
            string newParaContent)
                {
                    selectedId = newSelectedId;
                    selectedScope = newSelectedScope;
                    selectedText = newSelectedText;

                    notebookName = newNotebookName;
                    notebookColor = newNotebookColor;
                    sectionGroupName = newSectionGroupName;
                    sectionName = newSectionName;
                    sectionColor = newSectionColor;
                    pageName = newPageName;
                    paraContent = newParaContent;

                    if (label != null)
                    {
                        label.Text = selectedText ?? "No Selection";
                    }

                    if (comboScope != null)
                    {
                        int index = comboScope.FindStringExact(selectedScope);
                        comboScope.SelectedIndex = (index >= 0) ? index : -1;
                    }

                    if (labelNotebook != null)
                    {
                        labelNotebook.Text = $"Notebook: {notebookName ?? "N/A"} [{notebookColor ?? "No Color"}]";
                    }
                    if (labelSection != null)
                    {
                        labelSection.Text = $"Section: {sectionName ?? "N/A"} [{sectionColor ?? "No Color"}] | Group: {sectionGroupName ?? "N/A"}";
                    }
                    if (labelPage != null)
                    {
                        labelPage.Text = $"Page: {pageName ?? "N/A"}";
                    }
                    if (labelPara != null)
                    {
                        labelPara.Text = $"Paragraph: {paraContent ?? "N/A"}";
                    }

                    // Decide what list to show:
                    if (cachedList == null)
                        RefreshGridDisplay();
                    else
                        RefreshGridDisplay(cachedList);

                    if (btnSave != null)
                    {
                        btnSave.Enabled = !string.IsNullOrEmpty(selectedId);
                    }
                    if (btnDelete != null)
                    {
                        btnDelete.Enabled = grid.SelectedRows.Count > 0;
                    }
                }
        #endregion

        #region CONTEXT MENU HANDLERS   
        private void ShowMethodLogs_Click(object sender, EventArgs e)
        {
            var win = new MethodLogWindow();
            win.ShowDialog(this);
        }
        private void AddUrlBookmark_Click(object sender, EventArgs e)
        {
            using (var form = new UrlBookmarkForm())
            {
                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    var newBookmark = new BookmarkItem
                    {
                        Type = "Bookmark",
                        Scope = "URL",
                        Name = form.BookmarkName,
                        ParentId = null,
                        Id = "URL_" + Guid.NewGuid().ToString(),
                        OriginalId = form.BookmarkUrl, // Store URL here
                        NotebookName = "",
                        NotebookColor = "",
                        SectionGroupName = "",
                        SectionName = "",
                        SectionColor = "",
                        PageName = "",
                        ParaContent = "",
                        Notes = "",
                        SortOrder = items.Count
                    };

                    items.Add(newBookmark);
                    SaveToFile();
                    cachedList = null;
                    RefreshGridDisplay();
                }
            }
        }
        private void Settings_Click(object sender, EventArgs e)
        {
            using (FontSizeDialog dlg = new FontSizeDialog())
            {
                dlg.FontSize = this.Font.Size;
                dlg.StartPosition = FormStartPosition.CenterParent;

                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    float newSize = dlg.FontSize;

                    // Update the main form font
                    this.Font = new Font(this.Font.FontFamily, newSize);

                    // Update DataGridView fonts
                    grid.DefaultCellStyle.Font = new Font(grid.DefaultCellStyle.Font.FontFamily, newSize);
                    grid.ColumnHeadersDefaultCellStyle.Font = new Font(grid.ColumnHeadersDefaultCellStyle.Font.FontFamily, newSize, FontStyle.Bold);

                    // Update ListBox font
                    listScope.Font = new Font(listScope.Font.FontFamily, newSize);

                    // Calculate scale factor based on font size change ratio
                    float scaleFactor = newSize / this.Font.Size;

                    // Resize main window by scaling width and height
                    this.Width = (int)(this.Width * scaleFactor);
                    this.Height = (int)(this.Height * scaleFactor);

                    // Optionally limit minimum and maximum size for main window

                    // Resize listScope (ListBox) proportionally
                    listScope.Width = (int)(listScope.Width * scaleFactor);
                    listScope.Height = (int)(listScope.Height * scaleFactor);

                    // Resize grid (DataGridView) proportionally
                    grid.Width = (int)(grid.Width * scaleFactor);
                    grid.Height = (int)(grid.Height * scaleFactor);

                    // Optionally reposition controls if needed for layout consistency
        
                }
            }
        }
        private void grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                clickedColumnHeader = grid.Columns[e.ColumnIndex];
                columnHeaderContextMenu.Show(Cursor.Position);
            }
            else if (e.ColumnIndex >= 0 && grid.Columns[e.ColumnIndex].Name == "Name" && e.Button == MouseButtons.Left)
            {
                ToggleSort();
            }
        }
        private void Export_All_Bookmarks_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog sfd = new SaveFileDialog())
            {
                sfd.Filter = "CSV files (*.csv)|*.csv";
                sfd.FileName = "Bookmarks.csv";

                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        StringBuilder sb = new StringBuilder();

                        // Write header
                        var headers = grid.Columns
                            .Cast<DataGridViewColumn>()
                            .Select(col => QuoteValue(col.HeaderText));
                        sb.AppendLine(string.Join(",", headers));

                        // Write rows
                        foreach (DataGridViewRow row in grid.Rows)
                        {
                            if (!row.IsNewRow) // skip the new row placeholder
                            {
                                var cells = row.Cells.Cast<DataGridViewCell>()
                                    .Select(cell => QuoteValue(cell.Value));
                                sb.AppendLine(string.Join(",", cells));
                            }
                        }

                        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
                        MessageBox.Show("Export completed!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error exporting: " + ex.Message);
                    }
                }
            }
        }
        private void NewFolder_Click(object sender, EventArgs e)
        {
            var currentRow = GetSelectedItem();
            string parentId = currentRow?.Id;

            if (currentRow != null && currentRow.Type == "Bookmark")
            {
                parentId = currentRow.ParentId;
            }

            string baseFolderName = "New Folder";
            string folderName = baseFolderName;
            int copyNum = 1;
            while (items.Any(i => i.ParentId == parentId && i.Name == folderName && i.Type == "Folder"))
            {
                folderName = $"{baseFolderName} {copyNum++}";
            }

            var newFolder = new BookmarkItem
            {
                Type = "Folder",
                Id = Guid.NewGuid().ToString(),
                ParentId = parentId,
                Name = folderName,
                NotebookName = "",
                NotebookColor = "",
                SectionGroupName = "",
                SectionName = "",
                SectionColor = "",
                PageName = "",
                ParaContent = ""
            };

            items.Add(newFolder);
            SaveToFile();
            cachedList = null; // reset cache on data change
            RefreshGridDisplay();
        }
        private void Rename_Click(object sender, EventArgs e)
        {
            var currentRow = GetSelectedItem();
            if (currentRow == null) return;

            if (grid.SelectedRows.Count == 0) return;
            var rowIndex = grid.SelectedRows[0].Index;

            var nameColIndex = grid.Columns["Name"]?.Index ?? -1;
            if (nameColIndex < 0) return;
            grid.Rows[rowIndex].Cells[nameColIndex].Value = "";
            grid.CurrentCell = grid.Rows[rowIndex].Cells[nameColIndex];
            grid.BeginEdit(true);
        }
        private void Delete_Click(object sender, EventArgs e)
        {
            BtnDelete_Click(sender, e);
        }
        private void Open_All_Notes(object sender, EventArgs e)
        {
            var currentRow = GetSelectedItem();
            if (currentRow == null)
            {
                MessageBox.Show("No row selected.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if(currentRow.Scope == "URL")
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = currentRow.OriginalId,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to open URL: " + ex.Message);
                }
            }
            else if (currentRow.Type == "Bookmark" )
            {
                try
                {
                    var app = new Microsoft.Office.Interop.OneNote.Application();
                    app.NavigateTo(currentRow.OriginalId, "", true); // true opens in new window
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Failed to open OneNote object: " + ex.Message);
                }
                return;
            }
            if (currentRow.Type == "Folder")
            {
                // Get all bookmarks in this folder and its subfolders recursively
                var bookmarks = GetOneNoteObjectsRecursive(currentRow.Id);
                if (bookmarks.Count == 0)
                {
                    MessageBox.Show("No OneNote objects found in this folder or its subfolders.");
                    return;
                }
                var app = new Microsoft.Office.Interop.OneNote.Application();
                int count = 0;
                foreach (var bm in bookmarks)
                {
                    try
                    {
                        if(bm.Scope != "URL")
                        {
                            // Open each in a new window
                            app.NavigateTo(bm.OriginalId, "", true);
                            count++;
                        }

                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Failed to open: {bm.Name}\n{ex.Message}");
                    }
                }
            }
        }
        private List<BookmarkItem> GetOneNoteObjectsRecursive(string folderId)
        {
            var result = new List<BookmarkItem>();
            var children = items.Where(i => i.ParentId == folderId).ToList();

            foreach (var child in children)
            {
                if (child.Type == "Bookmark" && !string.IsNullOrEmpty(child.OriginalId))
                {
                    result.Add(child);
                }
                else if (child.Type == "Folder")
                {
                    result.AddRange(GetOneNoteObjectsRecursive(child.Id));
                }
            }
            return result;
        }
        private void TextWrap_Click(object sender, EventArgs e)
        {
            RefreshGridDisplay();
            isTextWrapEnabled = !isTextWrapEnabled;

            foreach (DataGridViewColumn column in grid.Columns)
            {
                column.DefaultCellStyle.WrapMode = isTextWrapEnabled ? DataGridViewTriState.True : DataGridViewTriState.False;
            }
            grid.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        }
        private void TextWrapCurrentRow_Click(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0) return;

            int rowIndex = grid.SelectedRows[0].Index;
            if (rowIndex < 0) return;

            foreach (DataGridViewCell cell in grid.Rows[rowIndex].Cells)
            {
                cell.Style.WrapMode = cell.Style.WrapMode == DataGridViewTriState.True
                    ? DataGridViewTriState.False
                    : DataGridViewTriState.True;
            }
            grid.AutoResizeRow(rowIndex, DataGridViewAutoSizeRowMode.AllCells);
        }
        private void TextWrapMenuItem_Click(object sender, EventArgs e)
        {
            if (clickedColumnHeader == null)
                return;

            var currentWrap = clickedColumnHeader.DefaultCellStyle.WrapMode;
            if (currentWrap == DataGridViewTriState.True)
                clickedColumnHeader.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
            else
                clickedColumnHeader.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            grid.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCells);
        }
        #endregion

        #region Fill Path Handlers
        private (string notebookName, string sectionGroupName, string sectionName, string pageName) GetPathFromOriginalId(OneNote.Application oneNoteApp, string pageId, string hierarchyXml)
        {
            try
            {
                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(hierarchyXml);

                var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                var pageNode = doc.SelectSingleNode($"//one:Page[@ID='{pageId}']", nsmgr);
                if (pageNode == null) return (null, null, null, null);

                // Section node (parent of page)
                var sectionNode = pageNode.ParentNode;
                if (sectionNode == null || sectionNode.Name != "one:Section") return (null, null, null, null);

                // Notebook or SectionGroup node (parent of section)
                var notebookNode = sectionNode.ParentNode;
                SectionGroupModel sectionGroupModel = null;
                string sectionGroupName = null;

                if (notebookNode != null && notebookNode.Name == "one:SectionGroup")
                {
                    var sectionGroupNode = notebookNode;
                    notebookNode = sectionGroupNode.ParentNode;

                    sectionGroupName = sectionGroupNode.Attributes["name"]?.Value;
                }

                if (notebookNode == null || notebookNode.Name != "one:Notebook") return (null, null, null, null);

                string notebookName = notebookNode.Attributes["name"]?.Value;
                string sectionName = sectionNode.Attributes["name"]?.Value;
                string pageName = pageNode.Attributes["name"]?.Value;

                return (notebookName, sectionGroupName, sectionName, pageName);
            }
            catch
            {
                // Handle exceptions or return null infos
                return (null, null, null, null);
            }
        }

        private string GetBookmarkName(string scope, string notebookName, string sectionGroupName, string sectionName, string pageName, string paraContent)
        {
            switch (scope)
            {
                case "Current Paragraph":
                    return paraContent ?? "Unnamed Paragraph";
                case "Current Page":
                    return pageName ?? "Unnamed Page";
                case "Current Section":
                    return sectionName ?? "Unnamed Section";
                case "Current Section Group":
                    return sectionGroupName ?? "Unnamed Section Group";
                case "Current Notebook":
                    return notebookName ?? "Unnamed Notebook";
                default:
                    return "Unnamed Bookmark";
            }
        }

        #endregion

        #region List Scope Handlers
        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;

                string id = grid.Rows[e.RowIndex].Cells["Id"].Value?.ToString();
                if (string.IsNullOrEmpty(id)) return;

                // Use cachedList if sorting is active
                var sourceList = cachedList ?? items;
                var item = sourceList.FirstOrDefault(i => i.Id == id);
                if (item == null) return;

                string clickedColumn = grid.Columns[e.ColumnIndex].Name;
                if (clickedColumn == "Name" && item.Scope == "URL")
                {

                        try
                        {
                            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                            {
                                FileName = item.OriginalId,
                                UseShellExecute = true
                            });
                            this.Hide();
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Failed to open URL: " + ex.Message);
                        }
                    
                }
                else if (clickedColumn == "Name" && item.Type == "Bookmark")
                {
                    try
                    {
                        // Open the OneNote object (Notebook, Section, Page)
                        var app = new Microsoft.Office.Interop.OneNote.Application();

                        // If your "id" can refer to Notebook, Section, or Page, 
                        // you may use NavigateTo with right id

                        // NavigateTo(string bstrHierarchyID, string bstrObjectID, bool fNewWindow)
                        // (see OneNote Interop docs for alternatives if you want more granularity)
                        app.NavigateTo(item.OriginalId);
                        // Optionally, to open in a new window:
                        // app.NavigateTo(id, "", true);
                        this.Hide();
                    }
                    catch (Exception exNav)
                    {
                        MessageBox.Show("Failed to open OneNote object: " + exNav.Message);
                    }

                }
                else if (clickedColumn == "Name" && item.Type == "Folder")
                {
                    item.IsExpanded = !item.IsExpanded;
                    SaveToFile();
                    RefreshGridDisplay(cachedList ?? null);
                }
                else if (clickedColumn == "Notes")
                {
                    grid.Rows[e.RowIndex].Cells["Notes"].ReadOnly = false;
                    grid.BeginEdit(true);
                }
                

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error handling click: " + ex.Message);
            }
        }
        private void ListScope_Click(object sender, EventArgs e)
        {
            using (MethodTimerLog.Time("ListScope_Click"))
            {
                try
                {
                    var oneNoteApp = new OneNote.Application();
                    Window currentWindow = oneNoteApp.Windows.CurrentWindow;

                    // Use Onenote to get Notebook hyperlink and section hyperlink and using it get notebook name and section name
                    string notebookId = currentWindow.CurrentNotebookId;
                    string sectionId = currentWindow.CurrentSectionId;
                    oneNoteApp.GetHyperlinkToObject(notebookId, null, out string notebookLink);
                    oneNoteApp.GetHyperlinkToObject(sectionId, null, out string sectionLink);
                    string notebookPath = notebookLink.Replace("onenote:///", "");
                    string sectionPath = sectionLink.Replace("onenote:///", "");
                    notebookPath = Uri.UnescapeDataString(notebookPath);
                    sectionPath = Uri.UnescapeDataString(sectionPath);

                    string notebookNames = Path.GetFileName(notebookPath);
                    string sectionFile = Path.GetFileNameWithoutExtension(sectionPath.Split('#')[0]);
                    var sectionParts = sectionPath.Split('#')[0].Split(Path.DirectorySeparatorChar);
                    var groups = sectionParts.SkipWhile(p => p != notebookNames).Skip(1).Take(sectionParts.Length - 2).ToList();
                    string sectionGroupNames = (groups.Count > 1 && groups.Any()) ? string.Join(" > ", groups[0]) : "";

                    var model = GetCurrentNotebookModel(oneNoteApp);
                    if (model == null)
                    {
                        MessageBox.Show("Failed to load the current notebook model.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    // Extract data from your model
                    string selectedId = model.Page?.Id ?? "";
                    string selectedScope = "page";
                    string displayText = model.Page?.Name ?? "";
                    string notebookName = notebookNames ?? "";
                    string notebookColor = model.NotebookColor ?? "";
                    string sectionGroupName = sectionGroupNames ?? "";
                    string sectionName = sectionFile ?? "";
                    string sectionColor = model.Section?.Color ?? "";
                    string pageName = model.Page.Name ?? "";
                    string paraContent = model.Page?.Paragraphs?.FirstOrDefault()?.Name ?? "";

                    if (string.IsNullOrEmpty(selectedId))
                    {
                        MessageBox.Show("No bookmark selected to save.");
                        return;
                    }
                    UpdateBookmarkInfo(
                                selectedId, selectedScope, displayText,
                                notebookName, notebookColor,
                                sectionGroupName, sectionName, sectionColor,
                                pageName, paraContent);


                    string bookmarkName;
                    string selected = listScope.SelectedItem.ToString();

                    switch (selected)
                    {
                        case "Current Paragraph":
                            bookmarkName = paraContent ?? "Unnamed Paragraph";
                            break;
                        case "Current Page":
                            bookmarkName = pageName ?? "Unnamed Page";
                            break;
                        case "Current Section":
                            bookmarkName = sectionName ?? "Unnamed Section";
                            break;
                        case "Current Section Group":
                            bookmarkName = sectionGroupName ?? "Unnamed Section Group";
                            break;
                        case "Current Notebook":
                            bookmarkName = notebookName ?? "Unnamed Notebook";
                            break;
                        default:
                            bookmarkName = "Unnamed Bookmark";
                            break;
                    }

                    var newBookmark = new BookmarkItem
                    {
                        Type = "Bookmark",
                        Scope = selected,
                        Name = bookmarkName,
                        ParentId = null,
                        Id = selectedId + "_" + Guid.NewGuid().ToString(), // unique composite ID
                        OriginalId = selectedId, // store the actual OneNote ID separately
                        NotebookName = notebookName,
                        NotebookColor = notebookColor,
                        SectionGroupName = sectionGroupName,
                        SectionName = sectionName,
                        SectionColor = sectionColor,
                        PageName = pageName,
                        ParaContent = paraContent,
                        Notes = ""
                    };

                    items.Insert(0, newBookmark);
                    SaveToFile();
                    cachedList = null;
                    RefreshGridDisplay();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error saving: " + ex.Message);
                }
            }
        }
        #endregion

        #region SAVE AND DELETE TXT HANDLERS
        private void BtnDelete_Click(object sender, EventArgs e)
        {
            if (grid.SelectedRows.Count == 0)
            {
                MessageBox.Show("Please select a row to delete.");
                return;
            }

            var selectedRow = grid.SelectedRows[0];
            var itemId = selectedRow.Cells["Id"].Value?.ToString();

            if (string.IsNullOrEmpty(itemId))
            {
                MessageBox.Show("Selected row has invalid ID.");
                return;
            }

            // *** Ask for confirmation BEFORE deleting ***
            var item = items.FirstOrDefault(i => i.Id == itemId);
            if (item == null) return;

            string message = $"Are you sure you want to delete the {(item.Type == "Folder" ? "folder" : "bookmark")} \"{item.Name}\"?";
            if (item.Type == "Folder")
            {
                message += "\n\nAll its subfolders and bookmarks will also be deleted.";
            }

            var result = MessageBox.Show(message, "Confirm Delete", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            if (result != DialogResult.Yes)
                return;

            // Proceed with removal if confirmed
            RemoveItemAndChildren(itemId);

            SaveToFile();
            cachedList = null; // reset cache on data change
            RefreshGridDisplay();

        }
        private void RemoveItemAndChildren(string id)
        {
            var toRemove = items.Where(i => i.Id == id).ToList();

            foreach (var item in toRemove)
            {
                items.Remove(item);
                if (item.Type == "Folder")
                {
                    var children = items.Where(c => c.ParentId == item.Id).Select(c => c.Id).ToList();
                    foreach (var childId in children)
                        RemoveItemAndChildren(childId);
                }
            }
        }
        private void LoadTable()
        {
            using (MethodTimerLog.Time("LoadTable"))
            {
                items.Clear();
                if (!File.Exists(tablePath)) return;
                try
                {
                    var lines = File.ReadAllLines(tablePath);
                    foreach (var line in lines)
                    {
                        var parts = line.Split(new[] { ',' }, 16); // Now 16 fields
                        var item = new BookmarkItem
                        {
                            Type = parts[0],
                            Scope = parts[1],
                            Id = parts[2],
                            OriginalId = parts[3], // NEW!
                            ParentId = parts[4] == "null" ? null : parts[4],
                            Name = parts[5],
                            NotebookName = parts[6],
                            NotebookColor = parts[7],
                            SectionGroupName = parts[8],
                            SectionName = parts[9],
                            SectionColor = parts[10],
                            PageName = parts[11],
                            ParaContent = parts[12],
                            Notes = parts[13],
                            IsExpanded = parts[14] == "1",
                            SortOrder = (parts.Length >= 16 && int.TryParse(parts[15], out var so)) ? so : 0
                        };

                        items.Add(item);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error loading bookmarks: " + ex.Message);
                }
            }

        }
        private void SaveToFile()
        {
            using (MethodTimerLog.Time("SaveToFile"))
            {
                try
                {
                    var lines = items.Select(i => string.Join(",", new[]{
                    EscapeCsv(i.Type),
                    EscapeCsv(i.Scope ?? ""),
                    EscapeCsv(i.Id),
                    EscapeCsv(i.OriginalId ?? ""),                 // <--- NEW
                    EscapeCsv(i.ParentId ?? "null"),
                    EscapeCsv(i.Name),
                    EscapeCsv(i.NotebookName),
                    EscapeCsv(i.NotebookColor),
                    EscapeCsv(i.SectionGroupName),
                    EscapeCsv(i.SectionName),
                    EscapeCsv(i.SectionColor),
                    EscapeCsv(i.PageName),
                    EscapeCsv(i.ParaContent),
                    EscapeCsv(i.Notes ?? ""),
                    i.IsExpanded ? "1" : "0",
                    i.SortOrder.ToString()
                })).ToList();
                    File.WriteAllLines(tablePath, lines);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error saving bookmarks: " + ex.Message);
                }
            }
            
        }
        private string EscapeCsv(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            if (input.Contains(",") || input.Contains("\"") || input.Contains("\n"))
                return "\"" + input.Replace("\"", "\"\"") + "\"";
            return input;
        }
        #endregion

        #region HELPERS
        public AddInModel GetCurrentNotebookModel(OneNote.Application oneNoteApp)
        {
            using (MethodTimerLog.Time("GetCurrentNotebookModel"))
            {
                var model = new AddInModel();

                try
                {
                    // Get current page ID from active OneNote window
                    var window = oneNoteApp.Windows.CurrentWindow;
                    string currentPageId = window.CurrentPageId;
                    if (string.IsNullOrEmpty(currentPageId))
                        return null;

                    // Get only the current page XML (no full hierarchy)
                    string pageXml;
                    oneNoteApp.GetPageContent(currentPageId, out pageXml, OneNote.PageInfo.piBasic);

                    var doc = new XmlDocument();
                    doc.LoadXml(pageXml);

                    var nsmgr = new XmlNamespaceManager(doc.NameTable);
                    nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                    // Current page node
                    var pageNode = doc.SelectSingleNode("//one:Page", nsmgr);
                    if (pageNode == null)
                        return null;

                    // Build current page model
                    var pageModel = new PageModel
                    {
                        Id = pageNode.Attributes?["ID"]?.Value,
                        Name = pageNode.Attributes?["name"]?.Value
                    };

                    // Optionally load paragraphs/text inside page
                    LoadParagraphs(oneNoteApp, pageModel);
                    model.Page = pageModel;

                    // Section node (page’s direct parent)
                    var sectionNode = pageNode.ParentNode;
                    if (sectionNode == null || sectionNode.Name != "one:Section")
                        return model; // stop if we can't find section

                    var sectionModel = new SectionModel
                    {
                        Id = sectionNode.Attributes?["ID"]?.Value,
                        Name = sectionNode.Attributes?["name"]?.Value,
                        Color = sectionNode.Attributes?["color"]?.Value
                    };
                    model.Section = sectionModel;

                    // Check if section belongs to SectionGroup
                    var parentNode = sectionNode.ParentNode;
                    if (parentNode != null && parentNode.Name == "one:SectionGroup")
                    {
                        model.SectionGroup = new SectionGroupModel
                        {
                            Id = parentNode.Attributes?["ID"]?.Value,
                            Name = parentNode.Attributes?["name"]?.Value
                        };

                        // Notebook is above SectionGroup
                        parentNode = parentNode.ParentNode;
                    }

                    // Notebook details
                    if (parentNode != null && parentNode.Name == "one:Notebook")
                    {
                        model.NotebookId = parentNode.Attributes?["ID"]?.Value;
                        model.NotebookName = parentNode.Attributes?["name"]?.Value;
                        model.NotebookColor = parentNode.Attributes?["color"]?.Value;
                    }

                    return model;
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error loading OneNote info: {ex.Message}", "Error",
                        MessageBoxButtons.OK, MessageBoxIcon.Error);
                    throw;
                }
            }
        }
        private void LoadParagraphs(OneNote.Application oneNoteApp, PageModel page)
        {
            try
            {
                string pageXml;
                oneNoteApp.GetPageContent(page.Id, out pageXml, PageInfo.piAll);

                var doc = new System.Xml.XmlDocument();
                doc.LoadXml(pageXml);

                var nsmgr = new System.Xml.XmlNamespaceManager(doc.NameTable);
                nsmgr.AddNamespace("one", "http://schemas.microsoft.com/office/onenote/2013/onenote");

                // First try: get the paragraph that was last selected
                var selectedParaNode = doc.SelectSingleNode("//one:OE[@selected]/one:T", nsmgr);

                // Fallback: if nothing selected, take the first paragraph
                if (selectedParaNode == null)
                {
                    selectedParaNode = doc.SelectSingleNode("//one:Outline/one:OEChildren/one:OE/one:T", nsmgr);
                }

                if (selectedParaNode != null)
                {
                    string paraText = selectedParaNode.InnerText?.Trim();
                    if (!string.IsNullOrEmpty(paraText))
                    {
                        page.Paragraphs.Add(new ParagraphModel
                        {
                            Id = page.Id + "_selected",
                            Name = paraText
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error loading paragraphs for page {page.Name}: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                throw;
            }
        }
        private List<BookmarkItem> FlattenForDisplay(string parentId, int depth)
        {
            var result = new List<BookmarkItem>();

            // Get all children (both folders and bookmarks) in SortOrder
            var children = items
                .Where(i => i.ParentId == parentId)
                .OrderBy(i => i.SortOrder)
                .ToList();

            foreach (var child in children)
            {
                result.Add(child);

                // If the child is an expanded folder, recurse into it
                if (child.Type == "Folder" && child.IsExpanded)
                {
                    result.AddRange(FlattenForDisplay(child.Id, depth + 1));
                }
            }

            return result;
        }
        private List<BookmarkItem> FlattenForDisplaySorted(string parentId, int depth, bool ascending)
        {
            var result = new List<BookmarkItem>();

            var folders = ascending ?
                items.Where(i => i.ParentId == parentId && i.Type == "Folder").OrderBy(i => i.Name).ToList() :
                items.Where(i => i.ParentId == parentId && i.Type == "Folder").OrderByDescending(i => i.Name).ToList();

            foreach (var folder in folders)
            {
                result.Add(folder);
                result.AddRange(FlattenForDisplaySorted(folder.Id, depth + 1, ascending));
            }

            var bookmarks = ascending ?
                items.Where(i => i.ParentId == parentId && i.Type == "Bookmark").OrderBy(i => i.Name).ToList() :
                items.Where(i => i.ParentId == parentId && i.Type == "Bookmark").OrderByDescending(i => i.Name).ToList();

            result.AddRange(bookmarks);

            return result;
        }
        private string RemoveIconsFromName(string nameWithIcons)
        {
            if (string.IsNullOrEmpty(nameWithIcons))
                return nameWithIcons;

            // Remove leading indentation spaces
            string cleaned = nameWithIcons.TrimStart();

            // Remove known icons
            string[] icons = { "📂", "📁", "📓", "📙", "📑", "📝", "¶", "📄", "📔", "🔖" };
            foreach (string icon in icons)
            {
                if (cleaned.StartsWith(icon + " "))
                {
                    cleaned = cleaned.Substring(icon.Length + 1);
                    break;
                }
                else if (cleaned.StartsWith(icon))
                {
                    cleaned = cleaned.Substring(icon.Length);
                    break;
                }
            }

            return cleaned.Trim();
        }
        private void ToggleSort()
        {
            if (!showingAlphabetical)
            {
                // First time -> show alphabetical order
                sortAscending = true; // or toggle if you want to alternate between asc/desc
                cachedList = FlattenForDisplaySorted(null, 0, sortAscending);

                RefreshGridDisplay(cachedList);
                showingAlphabetical = true;
            }
            else
            {
                // Going back -> show manual order from items.SortOrder
                cachedList = null; // forces RefreshGridDisplay to use FlattenForDisplay
                RefreshGridDisplay();

                showingAlphabetical = false;
            }
        }
        private int GetDepth(BookmarkItem item)
        {
            int depth = 0;
            string parentId = item.ParentId;
            while (parentId != null)
            {
                depth++;
                var parent = items.FirstOrDefault(i => i.Id == parentId);
                if (parent != null)
                    parentId = parent.ParentId;
                else
                    break;
            }
            return depth;
        }
        private void Form_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(this.Handle, WM_NCLBUTTONDOWN, HTCAPTION, 0);
            }
        }
        #endregion

        #region DRAG AND DROP
        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0) return;

                string id = grid.Rows[e.RowIndex].Cells["Id"].Value?.ToString();
                if (string.IsNullOrEmpty(id)) return;

                // Use cachedList if sorting is active
                var sourceList = cachedList ?? items;
                var item = sourceList.FirstOrDefault(i => i.Id == id);
                if (item == null) return;

                string clickedColumn = grid.Columns[e.ColumnIndex].Name;

                if (clickedColumn == "Name")
                {
                    grid.Rows[e.RowIndex].Cells["Name"].ReadOnly = false;
                    grid.BeginEdit(true);
                }
                else if (clickedColumn == "Notes")
                {
                    grid.Rows[e.RowIndex].Cells["Notes"].ReadOnly = false;
                    grid.BeginEdit(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error handling double-click: " + ex.Message);
            }
        }
        private void Grid_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                var hit = grid.HitTest(e.X, e.Y);
                if (hit.RowIndex >= 0)
                {
                    grid.ClearSelection();
                    grid.Rows[hit.RowIndex].Selected = true;
                    contextMenu.Show(grid, e.Location);
                }
            }
        }
        private void Grid_MouseDown_StartDrag(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                dragStart = new Point(e.X, e.Y);
        }
        private void Grid_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                if (Math.Abs(e.X - dragStart.X) > SystemInformation.DragSize.Width ||
                    Math.Abs(e.Y - dragStart.Y) > SystemInformation.DragSize.Height)
                {
                    if (grid.SelectedRows.Count > 0)
                    {
                        // Collect IDs of all selected rows
                        var selectedIds = grid.SelectedRows
                            .Cast<DataGridViewRow>()
                            .Select(r => r.Cells["Id"].Value?.ToString())
                            .Where(id => !string.IsNullOrEmpty(id))
                            .ToList();

                        if (selectedIds.Count > 0)
                        {
                            // Serialize IDs into a string (e.g., joined by a delimiter)
                            string dragData = string.Join(";", selectedIds);

                            grid.DoDragDrop(dragData, DragDropEffects.Move);
                        }
                    }
                }
            }
        }
        private void Grid_DragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(string)))
            {
                e.Effect = DragDropEffects.Move;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }
        private void Grid_DragDrop(object sender, DragEventArgs e)
        {
            if (!e.Data.GetDataPresent(typeof(string))) return;
            string draggedData = (string)e.Data.GetData(typeof(string));
            var draggedIds = draggedData.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
            if (draggedIds.Length == 0) return;

            Point clientPoint = grid.PointToClient(new Point(e.X, e.Y));
            var hitTest = grid.HitTest(clientPoint.X, clientPoint.Y);
            if (hitTest.RowIndex < 0) return;

            string targetId = grid.Rows[hitTest.RowIndex].Cells["Id"].Value.ToString();
            var targetItem = items.FirstOrDefault(i => i.Id == targetId);
            if (targetItem == null) return;

            // Prevent self-drop and drop into own descendants
            foreach (var did in draggedIds)
            {
                if (targetId == did || IsDescendant(targetId, did))
                    return;
            }

            Rectangle rowRect = grid.GetRowDisplayRectangle(hitTest.RowIndex, false);
            int rowHeight = rowRect.Height;
            int relativeY = clientPoint.Y - rowRect.Top;
            double dropPositionRatio = (double)relativeY / rowHeight;

            string parentId;
            int insertIndex;
            List<BookmarkItem> siblings;

            // ---- TOP ZONE ----
            if (dropPositionRatio <= 0.15)
            {
                parentId = targetItem.ParentId;
                siblings = items.Where(i => i.ParentId == parentId)
                                .OrderBy(i => i.SortOrder)
                                .ToList();
                insertIndex = siblings.FindIndex(i => i.Id == targetId);
                if (insertIndex < 0) insertIndex = siblings.Count;
            }
            // ---- BOTTOM ZONE ----
            else if (dropPositionRatio >= 0.65)
            {
                parentId = targetItem.ParentId;
                siblings = items.Where(i => i.ParentId == parentId)
                                .OrderBy(i => i.SortOrder)
                                .ToList();
                insertIndex = siblings.FindIndex(i => i.Id == targetId);
                if (insertIndex < 0) insertIndex = siblings.Count;
                insertIndex++; // after target
            }
            // ---- MIDDLE ZONE ON FOLDER ----
            else if (targetItem.Type == "Folder")
            {
                parentId = targetItem.Id;
                siblings = items.Where(i => i.ParentId == parentId)
                                .OrderBy(i => i.SortOrder)
                                .ToList();
                insertIndex = 0; // top of folder's children
            }
            // ---- MIDDLE ZONE ON BOOKMARK ----
            else
            {
                parentId = targetItem.ParentId;
                siblings = items.Where(i => i.ParentId == parentId)
                                .OrderBy(i => i.SortOrder)
                                .ToList();
                insertIndex = siblings.FindIndex(i => i.Id == targetId);
                if (insertIndex < 0) insertIndex = siblings.Count;
                insertIndex++; // after bookmark
            }

            // Remove dragged items from old location (avoid duplications)
            foreach (var did in draggedIds)
                siblings.RemoveAll(i => i.Id == did);

            // Insert at the calculated position
            foreach (var did in draggedIds)
            {
                var item = items.FirstOrDefault(i => i.Id == did);
                if (item != null)
                {
                    item.ParentId = parentId;
                    siblings.Insert(insertIndex++, item);
                }
            }

            // Update sort orders
            for (int i = 0; i < siblings.Count; i++)
            {
                siblings[i].SortOrder = i;
            }

            SaveToFile();
            cachedList = null;
            RefreshGridDisplay();
        }
        private bool IsDescendant(string potentialDescendantId, string ancestorId)
        {
            string parentId = items.FirstOrDefault(i => i.Id == potentialDescendantId)?.ParentId;
            while (!string.IsNullOrEmpty(parentId))
            {
                if (parentId == ancestorId)
                    return true;
                parentId = items.FirstOrDefault(i => i.Id == parentId)?.ParentId;
            }
            return false;
        }
        #endregion

        #region GRID RELATED METHODS
        private void Grid_CellPainting(object sender, DataGridViewCellPaintingEventArgs e)
        {
            if (e.ColumnIndex == grid.Columns["Name"].Index && e.RowIndex >= 0)
            {
                e.PaintBackground(e.CellBounds, true);
                var item = grid.Rows[e.RowIndex].Tag as BookmarkItem;
                if (item == null)
                {
                    e.Handled = false;
                    return;
                }

                Image img;

                if (item.Type == "Folder")
                {
                    img = item.IsExpanded
                        ? Properties.Resources.folder_open
                        : Properties.Resources.folder_close;
                }
                else
                {
                    switch (item.Scope)
                    {
                        case "Current Notebook":
                            img = Properties.Resources.note_icon;
                            break;
                        case "Current Section Group":
                            img = Properties.Resources.section_group;
                            break;
                        case "Current Section":
                            img = Properties.Resources.section;
                            break;
                        case "Current Page":
                            img = Properties.Resources.page;
                            break;
                        case "Current Paragraph":
                            img = Properties.Resources.para;
                            break;
                        case "URL":
                            img = Properties.Resources.url;
                            break;
                        default:
                            img = Properties.Resources.note_icon;
                            break;
                    }
                }

                // Calculate indent based on depth
                int spaceWidth = TextRenderer.MeasureText(" ", e.CellStyle.Font).Width;
                int indentPixels = 4 + GetDepth(item) * 2 * spaceWidth;

                int imageHeight = e.CellStyle.Font.Height;
                int imageWidth = (int)(img.Width * ((float)imageHeight / img.Height));
                int imageTop = e.CellBounds.Top + (e.CellBounds.Height - imageHeight) / 2;

                // Draw image
                Rectangle imageRect = new Rectangle(e.CellBounds.Left + indentPixels, imageTop, imageWidth, imageHeight);
                e.Graphics.DrawImage(img, imageRect);

                // Draw text
                string text = grid.Rows[e.RowIndex].Cells["Name"].Value?.ToString();
                if (!string.IsNullOrEmpty(text))
                {
                    Rectangle textRect = new Rectangle(
                        imageRect.Right + 4,
                        e.CellBounds.Top,
                        e.CellBounds.Width - imageRect.Width - indentPixels - 6,
                        e.CellBounds.Height
                    );
                    TextRenderer.DrawText(e.Graphics, text, e.CellStyle.Font, textRect, e.CellStyle.ForeColor, TextFormatFlags.VerticalCenter | TextFormatFlags.Left);
                }

                e.Handled = true;
            }
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.F2 && grid.SelectedCells.Count == 1)
                {
                    var cell = grid.SelectedCells[0];
                    if (cell.OwningColumn.Name == "Name")
                    {

                        var rowIndex = grid.SelectedRows[0].Index;
                        var nameColIndex = grid.Columns["Name"]?.Index ?? -1;
                        if (nameColIndex < 0) return;
                        grid.Rows[rowIndex].Cells[nameColIndex].Value = "";
                        grid.CurrentCell = cell;
                        grid.BeginEdit(true);
                        e.Handled = true;
                    }
                }
                else if (e.KeyCode == Keys.Enter && grid.SelectedRows.Count == 1)
                {
                    string id = grid.SelectedRows[0].Cells["Id"].Value?.ToString();
                    if (string.IsNullOrEmpty(id)) return;

                    // Use cachedList if available
                    var sourceList = cachedList ?? items;
                    var item = sourceList.FirstOrDefault(i => i.Id == id);
                    if (item == null) return;

                    if (item.Type == "Folder")
                    {
                        item.IsExpanded = !item.IsExpanded;
                        SaveToFile();
                        RefreshGridDisplay(cachedList ?? null);
                    }
                    else if (item.Type == "Bookmark")
                    {
                        try
                        {
                            var app = new Microsoft.Office.Interop.OneNote.Application();
                            app.NavigateTo(item.OriginalId);
                        }
                        catch (Exception exNav)
                        {
                            MessageBox.Show("Failed to open OneNote page: " + exNav.Message);
                        }
                    }
                    e.Handled = true;
                }
                else if (e.KeyCode == Keys.Escape)
                {
                    this.Hide();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in key handling: " + ex.Message);
            }
        }
        private void grid_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0)
                    return;

                var colName = grid.Columns[e.ColumnIndex].Name;
                var id = grid.Rows[e.RowIndex].Cells["Id"].Value?.ToString();
                if (id == null) return;

                var item = items.FirstOrDefault(i => i.Id == id);
                if (item == null) return;

                if (colName == "Notes")
                {
                    item.Notes = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    SaveToFile();
                    cachedList = null;
                }
                else if (colName == "Name")
                {
                    var newName = grid.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
                    // Step 1: Remove icons and indent spaces
                    string cleanName = RemoveIconsFromName(newName);

                    // Step 2: Keep only alphanumeric + underscore
                    cleanName = KeepAlphaNumericUnderscore(cleanName);
                    
                    if (!string.IsNullOrEmpty(cleanName) && cleanName != item.Name)
                    {
                        item.Name = cleanName;
                        SaveToFile();
                        cachedList = null;
                    }
                }
                RefreshGridDisplay();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating cell value: {ex.Message}");
            }
        }
        private string KeepAlphaNumericUnderscore(string input)
        {
            // Only allow A-Z, a-z, 0-9, and _
            return new string(input.Where(c => char.IsLetterOrDigit(c) || c == '_').ToArray());
        }
        private void RefreshGridDisplay(List<BookmarkItem> flatList = null)
        {
            using (MethodTimerLog.Time("RefreshGridDisplay"))
            {
                try
                {
                    grid.Columns.Clear();
                    grid.Rows.Clear();

                    // Setup columns without a separate image column
                    grid.Columns.Add("Type", "Type");
                    grid.Columns.Add("Scope", "Scope");
                    grid.Columns.Add("Name", "Name"); // This will display image + text
                    grid.Columns.Add("Id", "Id");
                    grid.Columns.Add("OriginalId", "OriginalId");
                    grid.Columns.Add("NotebookName", "Notebook Name");
                    grid.Columns.Add("NotebookColor", "Notebook Color");
                    grid.Columns.Add("SectionGroupName", "Section Group");
                    grid.Columns.Add("SectionName", "Section Name");
                    grid.Columns.Add("SectionColor", "Section Color");
                    grid.Columns.Add("PageName", "Page Name");
                    grid.Columns.Add("ParaContent", "Paragraph Content");
                    grid.Columns.Add("BookMarkPath", "BookMark Path");
                    grid.Columns.Add("Notes", "Notes");
                    grid.Columns.Add("Depth", "Depth");

                    // Set visibility to false for some columns as before
                    grid.Columns["Type"].Visible = false;
                    grid.Columns["Scope"].Visible = false;
                    grid.Columns["Id"].Visible = false;
                    grid.Columns["NotebookName"].Visible = false;
                    grid.Columns["NotebookColor"].Visible = false;
                    grid.Columns["SectionGroupName"].Visible = false;
                    grid.Columns["SectionName"].Visible = false;
                    grid.Columns["SectionColor"].Visible = false;
                    grid.Columns["ParaContent"].Visible = false;
                    grid.Columns["PageName"].Visible = false;
                    grid.Columns["Depth"].Visible = false;
                    grid.Columns["OriginalId"].Visible = false;
                    grid.Columns["Notes"].ReadOnly = false;
                    grid.Columns["Name"].ReadOnly = false;

                    // Event subscriptions and style settings
                    grid.KeyDown += Grid_KeyDown;
                    grid.BackgroundColor = ColorTranslator.FromHtml("#f3f3f3");
                    grid.BorderStyle = BorderStyle.None;
                    grid.DefaultCellStyle.Font = new Font("Segoe UI", 10);
                    grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);
                    grid.DefaultCellStyle.SelectionBackColor = ColorTranslator.FromHtml("#ddd9ec");
                    grid.DefaultCellStyle.SelectionForeColor = Color.Black;

                    grid.Columns["BookMarkPath"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    grid.Columns["Notes"].SortMode = DataGridViewColumnSortMode.NotSortable;
                    if (grid.Columns.Contains("Name"))
                        grid.Columns["Name"].SortMode = DataGridViewColumnSortMode.NotSortable;

                    // Set row height based on font size plus padding
                    int fontHeight = grid.DefaultCellStyle.Font.Height;
                    grid.RowTemplate.Height = fontHeight + 6; // Padding for aesthetics and image

                    // Subscribe to CellPainting to draw image + text in "Name" column
                    grid.CellPainting -= Grid_CellPainting; // Avoid duplicate subscriptions
                    grid.CellPainting += Grid_CellPainting;
                    grid.ScrollBars = ScrollBars.Both;

                    if (flatList == null)
                        flatList = FlattenForDisplay(null, 0);

                    foreach (var item in flatList)
                    {
                        int depth = GetDepth(item);
                        string bookmarkPath = item.NotebookName;
                        if (!string.IsNullOrWhiteSpace(item.SectionGroupName))
                            bookmarkPath += " - " + item.SectionGroupName;
                        bookmarkPath += " - " + item.SectionName + " - " + item.PageName;
                        string displayName = item.Name;

                        // Add the row with "Name" cell value set to displayName (image will be drawn by CellPainting)
                        int rowIndex = grid.Rows.Add(
                            item.Type,
                            item.Scope,
                            displayName, // Name column value
                            item.Id,
                            item.OriginalId,
                            item.NotebookName,
                            item.NotebookColor,
                            item.SectionGroupName,
                            item.SectionName,
                            item.SectionColor,
                            item.PageName,
                            item.ParaContent,
                            bookmarkPath,
                            item.Notes ?? string.Empty,
                            depth
                        );

                        // Store the item in the row's Tag for use in CellPainting
                        grid.Rows[rowIndex].Tag = item;

                        // Color coding for folders
                        if (item.Type == "Folder")
                        {
                            grid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.DarkBlue;
                            if (!item.IsExpanded) // collapsed folder
                            {
                                grid.Rows[rowIndex].DefaultCellStyle.Font = new Font(grid.Font, FontStyle.Bold);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in RefreshGridDisplay: {ex.Message}");
                }
            }
            
        }
        #endregion
        
        #region Window Size and Position
        private string QuoteValue(object val)
        {
            if (val == null) return "\"\"";
            string output = val.ToString().Replace("\"", "\"\"");
            return $"\"{output}\""; // always wrap in quotes
        }
        private void ApplyRoundedCorners(int radius)
        {
            var path = new GraphicsPath();
            int diameter = radius * 2;
            var rect = new Rectangle(0, 0, this.Width, this.Height);

            // Top-left corner
            path.AddArc(rect.X, rect.Y, diameter, diameter, 180, 90);
            // Top-right corner
            path.AddArc(rect.Right - diameter, rect.Y, diameter, diameter, 270, 90);
            // Bottom-right corner
            path.AddArc(rect.Right - diameter, rect.Bottom - diameter, diameter, diameter, 0, 90);
            // Bottom-left corner
            path.AddArc(rect.X, rect.Bottom - diameter, diameter, diameter, 90, 90);
            path.CloseFigure();

            this.Region = new Region(path);
        }
        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);

        }
        protected override void WndProc(ref Message m)
        {
            const int WM_NCHITTEST = 0x84;
            const int WM_ACTIVATEAPP = 0x001C;

            // --- Handle "click outside" to hide form ---
            if (m.Msg == WM_ACTIVATEAPP)
            {
                bool active = m.WParam != IntPtr.Zero;
                if (!active)
                {
                    this.Hide();
                }
            }

            // --- Keep your resize handling ---
            if (m.Msg == WM_NCHITTEST)
            {
                Point pos = PointToClient(Cursor.Position);
                int resizeDir = 0;
                int ResizeBorder = 8; // Set your custom border size

                if (pos.X < ResizeBorder) resizeDir |= 1;
                else if (pos.X > Width - ResizeBorder) resizeDir |= 2;
                if (pos.Y < ResizeBorder) resizeDir |= 4;
                else if (pos.Y > Height - ResizeBorder) resizeDir |= 8;

                if (resizeDir != 0)
                {
                    switch (resizeDir)
                    {
                        case 5: m.Result = (IntPtr)13; return; // top-left
                        case 6: m.Result = (IntPtr)14; return; // top-right
                        case 9: m.Result = (IntPtr)16; return; // bottom-left
                        case 10: m.Result = (IntPtr)17; return; // bottom-right
                        case 1: m.Result = (IntPtr)10; return; // left
                        case 2: m.Result = (IntPtr)11; return; // right
                        case 4: m.Result = (IntPtr)12; return; // top
                        case 8: m.Result = (IntPtr)15; return; // bottom
                    }
                }
            }

            base.WndProc(ref m);
        }
        #endregion

        #region SOON TO BE REMOVED
        public static class Prompt
        {
            public static string ShowDialog(string text, string caption, string defaultText)
            {
                Form prompt = new Form()
                {
                    Width = 400,
                    Height = 150,
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    Text = caption,
                    StartPosition = FormStartPosition.CenterParent,
                    MinimizeBox = false,
                    MaximizeBox = false
                };
                Label textLabel = new Label() { Left = 20, Top = 20, Text = text, AutoSize = true };
                TextBox inputBox = new TextBox() { Left = 20, Top = 50, Width = 340, Text = defaultText };
                Button confirmation = new Button() { Text = "OK", Left = 280, Width = 80, Top = 80, DialogResult = DialogResult.OK };
                confirmation.Click += (sender, e) => { prompt.Close(); };
                prompt.Controls.Add(textLabel);
                prompt.Controls.Add(inputBox);
                prompt.Controls.Add(confirmation);
                prompt.AcceptButton = confirmation;

                return prompt.ShowDialog() == DialogResult.OK ? inputBox.Text : null;
            }
        }
        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(selectedId))
                {
                    MessageBox.Show("No bookmark selected to save.");
                    return;
                }

                string bookmarkName;
                string selected = selectedScope; // Use selectedScope not selectedId

                switch (selected)
                {
                    case "Current Paragraph":
                        bookmarkName = paraContent ?? "Unnamed Paragraph";
                        break;
                    case "Current Page":
                        bookmarkName = pageName ?? "Unnamed Page";
                        break;
                    case "Current Section":
                        bookmarkName = sectionName ?? "Unnamed Section";
                        break;
                    case "Current Section Group":
                        bookmarkName = sectionGroupName ?? "Unnamed Section Group";
                        break;
                    case "Current Notebook":
                        bookmarkName = notebookName ?? "Unnamed Notebook";
                        break;
                    default:
                        bookmarkName = "Unnamed Bookmark";
                        break;
                }

                var newBookmark = new BookmarkItem
                {
                    Type = "Bookmark",
                    Scope = selected,
                    Name = bookmarkName,
                    ParentId = null,
                    Id = selectedId + "_" + Guid.NewGuid().ToString(),    // Composite unique ID
                    OriginalId = selectedId, // The actual OneNote ID
                    NotebookName = notebookName,
                    NotebookColor = notebookColor,
                    SectionGroupName = sectionGroupName,
                    SectionName = sectionName,
                    SectionColor = sectionColor,
                    PageName = pageName,
                    ParaContent = paraContent,
                    Notes = "",
                    SortOrder = items.Count // append at end
                };

                // Do NOT remove duplicates!
                items.Add(newBookmark);

                SaveToFile();
                cachedList = null;  // reset cached list on data change
                RefreshGridDisplay();
                this.Hide();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving: " + ex.Message);
            }
        }
        private string IndentName(string name, int depth, bool isFolder = false, bool expanded = true, string scope = "")
        {
            string indent = new string(' ', depth * 6);

            if (isFolder)
            {
                // Folder open or closed icon
                return indent + (expanded ? "📂 " : "📁 ") + name;
            }
            else
            {
                string icon;

                switch (scope)
                {
                    case "Current Notebook": icon = "📓 "; break;  // Notebook
                    case "Current Section Group": icon = "📙 "; break;  // Section Group
                    case "Current Section": icon = "📑 "; break;  // Section
                    case "Current Page": icon = "📝 "; break;  // Page
                    case "Current Paragraph": icon = "¶ "; break;  // Paragraph
                    case "File": icon = "📄 "; break;  // File
                    case "NotebookObject": icon = "📔 "; break;  // Notebook object
                    default: icon = "🔖 "; break;  // Generic bookmark
                }

                return indent + icon + name;
            }
        }
        #endregion
        public class FontSizeDialog : Form
        {
            private NumericUpDown numericFontSize;
            private Button btnOK;
            private Button btnCancel;

            public float FontSize { get; set; }

            public FontSizeDialog()
            {
                this.Text = "Settings - Font Size";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;
                this.Width = 250;
                this.Height = 150;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;
                this.TopMost = true;

                numericFontSize = new NumericUpDown()
                {
                    Minimum = 6,
                    Maximum = 72,
                    DecimalPlaces = 1,
                    Increment = 0.5M,
                    Location = new Point(20, 20),
                    Width = 100
                };
                this.Controls.Add(numericFontSize);

                btnOK = new Button()
                {
                    Text = "OK",
                    DialogResult = DialogResult.OK,
                    Location = new Point(20, 60),
                    Width = 80
                };
                btnOK.Click += (s, e) => { FontSize = (float)numericFontSize.Value; this.Close(); };
                this.Controls.Add(btnOK);

                btnCancel = new Button()
                {
                    Text = "Cancel",
                    DialogResult = DialogResult.Cancel,
                    Location = new Point(120, 60),
                    Width = 80
                };
                this.Controls.Add(btnCancel);
            }

            protected override void OnShown(EventArgs e)
            {
                base.OnShown(e);
                numericFontSize.Value = (decimal)FontSize;
            }
        }
        private class BookmarkItem
        {
            public string Type { get; set; } // "Folder" or "Bookmark"
            public string Scope { get; set; }
            public string Name { get; set; }
            public string ParentId { get; set; } // null means root
            public string Id { get; set; }       // unique row ID: pageId_GUID
            public string OriginalId { get; set; } // actual OneNote object ID for navigation
            public string NotebookName { get; set; }
            public string NotebookColor { get; set; }
            public string SectionGroupName { get; set; }
            public string SectionName { get; set; }
            public string SectionColor { get; set; }
            public string PageName { get; set; }
            public string ParaContent { get; set; }
            public string Notes { get; set; }
            public bool IsExpanded { get; set; } = true;
            public int SortOrder { get; set; }
        }
        private BookmarkItem GetSelectedItem()
        {
            if (grid.SelectedRows.Count == 0)
                return null;

            var selectedRow = grid.SelectedRows[0];
            var id = selectedRow.Cells["Id"].Value?.ToString();
            return items.FirstOrDefault(i => i.Id == id);
        }
        public class UrlBookmarkForm : Form
        {
            public string BookmarkUrl { get; private set; }
            public string BookmarkName { get; private set; }

            private TextBox txtUrl;
            private TextBox txtName;
            private Button btnOk;
            private Button btnCancel;

            public UrlBookmarkForm()
            {
                this.Text = "Add URL Bookmark";
                this.FormBorderStyle = FormBorderStyle.FixedDialog;
                this.StartPosition = FormStartPosition.CenterParent;
                this.Width = 400;
                this.Height = 160;
                this.MaximizeBox = false;
                this.MinimizeBox = false;
                this.ShowInTaskbar = false;

                Label lblUrl = new Label() { Text = "URL:", Left = 10, Top = 20, Width = 50 };
                txtUrl = new TextBox() { Left = 70, Top = 18, Width = 300 };
                txtUrl.TextChanged += TxtUrl_TextChanged;

                Label lblName = new Label() { Text = "Name:", Left = 10, Top = 50, Width = 50 };
                txtName = new TextBox() { Left = 70, Top = 48, Width = 300 };

                btnOk = new Button() { Text = "OK", Left = 210, Width = 70, Top = 85, DialogResult = DialogResult.OK };
                btnOk.Click += BtnOk_Click;

                btnCancel = new Button() { Text = "Cancel", Left = 290, Width = 70, Top = 85, DialogResult = DialogResult.Cancel };

                this.Controls.Add(lblUrl);
                this.Controls.Add(txtUrl);
                this.Controls.Add(lblName);
                this.Controls.Add(txtName);
                this.Controls.Add(btnOk);
                this.Controls.Add(btnCancel);

                this.AcceptButton = btnOk;
                this.CancelButton = btnCancel;
            }

            private void TxtUrl_TextChanged(object sender, EventArgs e)
            {
                try
                {
                    Uri uri = new Uri(txtUrl.Text);
                    txtName.Text = uri.Host;
                }
                catch
                {
                    // Invalid URL or empty, don't update name
                }
            }

            private void BtnOk_Click(object sender, EventArgs e)
            {
                if (string.IsNullOrWhiteSpace(txtUrl.Text))
                {
                    MessageBox.Show("Please enter a valid URL.", "Invalid Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    this.DialogResult = DialogResult.None;
                    return;
                }
                BookmarkUrl = txtUrl.Text.Trim();
                BookmarkName = string.IsNullOrWhiteSpace(txtName.Text) ? BookmarkUrl : txtName.Text.Trim();
            }
        }
        public static class MethodTimerLog
        {
            private static readonly List<string> Logs = new List<string>();
            private static readonly object LockObj = new object();

            public static IDisposable Time(string methodName)
            {
                var sw = System.Diagnostics.Stopwatch.StartNew();
                return new DisposableAction(() =>
                {
                    sw.Stop();
                    string logEntry = $"{DateTime.Now:HH:mm:ss} {methodName} took {sw.ElapsedMilliseconds}ms";
                    lock (LockObj)
                    {
                        Logs.Add(logEntry);
                    }
                });
            }

            public static List<string> GetLogs()
            {
                lock (LockObj)
                {
                    return new List<string>(Logs);
                }
            }

            public static void ClearLogs()
            {
                lock (LockObj)
                {
                    Logs.Clear();
                }
            }

            private class DisposableAction : IDisposable
            {
                private readonly Action _onDispose;
                public DisposableAction(Action onDispose) => _onDispose = onDispose;
                public void Dispose() => _onDispose?.Invoke();
            }
        }
        public class MethodLogWindow : Form
        {
            private ListBox logsListBox;
            private Button btnClear;

            public MethodLogWindow()
            {
                this.Text = "Method Time Logs";
                this.Width = 400;
                this.Height = 350;

                logsListBox = new ListBox()
                {
                    Dock = DockStyle.Top,
                    Height = 270
                };
                btnClear = new Button()
                {
                    Text = "Clear Logs",
                    Dock = DockStyle.Bottom
                };
                btnClear.Click += (s, e) =>
                {
                    MethodTimerLog.ClearLogs();
                    RefreshLogs();
                };

                Controls.Add(logsListBox);
                Controls.Add(btnClear);
                RefreshLogs();
            }

            private void RefreshLogs()
            {
                logsListBox.Items.Clear();
                foreach (var log in MethodTimerLog.GetLogs())
                    logsListBox.Items.Add(log);
            }
        }
    }
}