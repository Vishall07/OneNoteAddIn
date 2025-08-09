using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace CSOneNoteRibbonAddIn
{
    public class BookMark_Window : Form
    {
        private Label label;
        private ComboBox comboScope;
        private Button btnSave;
        private Button btnDelete;
        private DataGridView grid;
        private string selectedId;
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
                this.Width = 600;
                this.Height = 300;
                this.TopMost = true;
                this.BackColor = Color.White;

                labelNotebook = new Label { Location = new Point(20, 30), AutoSize = true };
                labelSection = new Label { Location = new Point(20, 70), AutoSize = true };
                labelPage = new Label { Location = new Point(20, 90), AutoSize = true };
                labelPara = new Label { Location = new Point(20, 110), AutoSize = true };

                comboScope = new ComboBox() { Location = new Point(20, 12), Width = 150 };
                comboScope.Items.AddRange(new[] {
                    "Current Paragraph", "Current Section Group", "Current Section", "Current Page", "Current Notebook"
                });

                btnSave = new Button { Location = new Point(180, 12), Text = "Save", Width = 80 };
                btnSave.Click += BtnSave_Click;

                grid = new DataGridView
                {
                    Location = new Point(20, 40),
                    Width = this.ClientSize.Width - 20,
                    Height = this.ClientSize.Height - 10,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    AllowDrop = true,
                    AllowUserToResizeColumns = true,
                    AllowUserToOrderColumns = true,
                    RowHeadersVisible = false
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

                contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("New Folder", null, NewFolder_Click);
                contextMenu.Items.Add("Rename", null, Rename_Click);
                contextMenu.Items.Add("Delete", null, Delete_Click);
                Controls.Add(comboScope);
                Controls.Add(btnSave);
                Controls.Add(btnDelete);
                Controls.Add(grid);
                //Controls.Add(labelNotebook);
                //Controls.Add(labelSection);
                //Controls.Add(labelPage);
                //Controls.Add(labelPara);

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

                LoadTable();
                UpdateBookmarkInfo(selectedId, selectedScope, selectedText, notebookName, notebookColor,
                    sectionGroupName, sectionName, sectionColor, pageName, paraContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing Bookmark window: " + ex.Message);
            }
        }

        // BookmarkItem class
        private class BookmarkItem
        {
            public string Type { get; set; } // "Folder" or "Bookmark"
            public string Name { get; set; }
            public string ParentId { get; set; } // null means root
            public string Id { get; set; }
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
                string selected = comboScope.SelectedItem as string;
                MessageBox.Show(selected);

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
                    Name = bookmarkName,
                    ParentId = null,
                    Id = selectedId,
                    NotebookName = notebookName,
                    NotebookColor = notebookColor,
                    SectionGroupName = sectionGroupName,
                    SectionName = sectionName,
                    SectionColor = sectionColor,
                    PageName = pageName,
                    ParaContent = paraContent,
                    Notes = ""
                };

                items.RemoveAll(i => i.Type == "Bookmark" && i.Id == newBookmark.Id);
                items.Add(newBookmark);

                SaveToFile();
                cachedList = null;  // reset cached list on data change
                RefreshGridDisplay();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving: " + ex.Message);
            }
        }
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
            items.Clear();
            if (!File.Exists(tablePath)) return;
            try
            {
                var lines = File.ReadAllLines(tablePath);
                foreach (var line in lines)
                {
                    var parts = line.Split(new[] { ',' }, 14);
                    var item = new BookmarkItem
                    {
                        Type = parts[0],
                        Id = parts[1],
                        ParentId = parts[2] == "null" ? null : parts[2],
                        Name = parts[3],
                        NotebookName = parts[4],
                        NotebookColor = parts[5],
                        SectionGroupName = parts[6],
                        SectionName = parts[7],
                        SectionColor = parts[8],
                        PageName = parts[9],
                        ParaContent = parts[10],
                        Notes = parts[11],
                        IsExpanded = parts[12] == "1",
                        SortOrder = (parts.Length >= 14 && int.TryParse(parts[13], out var so)) ? so : 0
                    };
                    items.Add(item);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading bookmarks: " + ex.Message);
            }
        }
        private void SaveToFile()
        {
            try
            {
                var lines = items.Select(i => string.Join(",", new[]{
                    EscapeCsv(i.Type),
                    EscapeCsv(i.Id),
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
                    i.SortOrder.ToString() // NEW
                })).ToList();
                File.WriteAllLines(tablePath, lines);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error saving bookmarks: " + ex.Message);
            }
        }
        private string EscapeCsv(string input)
        {
            if (string.IsNullOrEmpty(input)) return "";
            if (input.Contains(",") || input.Contains("\"") || input.Contains("\n"))
                return "\"" + input.Replace("\"", "\"\"") + "\"";
            return input;
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
                    if (!string.IsNullOrEmpty(newName) && newName != item.Name)
                    {
                        item.Name = newName;
                        SaveToFile();
                        cachedList = null;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating cell value: {ex.Message}");
            }
        }
        private void RefreshGridDisplay(List<BookmarkItem> flatList = null)
        {
            try
            {
                grid.Columns.Clear();
                grid.Rows.Clear();

                // same column setup as before...
                grid.Columns.Add("Type", "Type");
                grid.Columns.Add("Name", "Name");
                grid.Columns.Add("Id", "Id");
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

                grid.Columns["Type"].Visible = false;
                grid.Columns["Id"].Visible = false;
                grid.Columns["NotebookName"].Visible = false;
                grid.Columns["NotebookColor"].Visible = false;
                grid.Columns["SectionGroupName"].Visible = false;
                grid.Columns["SectionName"].Visible = false;
                grid.Columns["SectionColor"].Visible = false;
                grid.Columns["ParaContent"].Visible = false;
                grid.Columns["PageName"].Visible = false;
                grid.Columns["Depth"].Visible = false;
                grid.Columns["Notes"].ReadOnly = false;

                grid.Columns["Name"].ReadOnly = false;
                grid.KeyDown += Grid_KeyDown;

                if (grid.Columns.Contains("Name"))
                    grid.Columns["Name"].SortMode = DataGridViewColumnSortMode.NotSortable;

                if (flatList == null)
                    flatList = FlattenForDisplay(null, 0);

                foreach (var item in flatList)
                {
                    int depth = GetDepth(item);
                    string bookmarkPath = item.NotebookName;
                    if (!string.IsNullOrWhiteSpace(item.SectionGroupName))
                        bookmarkPath += " - " + item.SectionGroupName;
                    bookmarkPath += " - " + item.SectionName + " - " + item.PageName;

                    string displayName = IndentName(item.Name, depth, item.Type == "Folder", item.IsExpanded);

                    int rowIndex = grid.Rows.Add(
                        item.Type,
                        displayName,
                        item.Id,
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

                    // *** color coding folders
                    if (item.Type == "Folder")
                    {
                        grid.Rows[rowIndex].DefaultCellStyle.ForeColor = Color.DarkBlue;
                        if (!item.IsExpanded) // collapsed
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
        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            try
            {
                if (e.KeyCode == Keys.F2 && grid.SelectedCells.Count == 1)
                {
                    var cell = grid.SelectedCells[0];
                    if (cell.OwningColumn.Name == "Name")
                    {
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
                            app.NavigateTo(id);
                        }
                        catch (Exception exNav)
                        {
                            MessageBox.Show("Failed to open OneNote page: " + exNav.Message);
                        }
                    }
                    e.Handled = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error in key handling: " + ex.Message);
            }
        }

        private void grid_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (grid.Columns[e.ColumnIndex].Name == "Name")
            {
                ToggleSort();
            }
        }
        private List<BookmarkItem> FlattenForDisplay(string parentId, int depth)
        {
            var result = new List<BookmarkItem>();
            var folders = items.Where(i => i.ParentId == parentId && i.Type == "Folder")
                               .OrderBy(i => i.SortOrder).ToList();
            foreach (var folder in folders)
            {
                result.Add(folder);
                if (folder.IsExpanded)
                    result.AddRange(FlattenForDisplay(folder.Id, depth + 1));
            }
            var bookmarks = items.Where(i => i.ParentId == parentId && i.Type == "Bookmark")
                                 .OrderBy(i => i.SortOrder).ToList();
            result.AddRange(bookmarks);
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
        private string IndentName(string name, int depth, bool isFolder = false, bool expanded = true)
        {
            string icon = "";
            if (isFolder)
            {
                icon = expanded ? "▼ " : "► "; // *** visual arrow for expand/collapse
            }
            return new string(' ', depth * 4) + icon + name;
        }
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

                if (clickedColumn == "Name" && item.Type == "Folder")
                {
                    item.IsExpanded = !item.IsExpanded;
                    SaveToFile();
                    RefreshGridDisplay(cachedList ?? null);
                }
                else if (clickedColumn == "Name" && item.Type == "Bookmark")
                {
                    try
                    {
                        var app = new Microsoft.Office.Interop.OneNote.Application();
                        app.NavigateTo(id);
                    }
                    catch (Exception exNav)
                    {
                        MessageBox.Show("Failed to open OneNote page: " + exNav.Message);
                    }
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

            string parentId = targetItem.Type == "Folder" ? targetItem.Id : targetItem.ParentId;
            var siblings = items.Where(i => i.ParentId == parentId).OrderBy(i => i.SortOrder).ToList();
            foreach (var did in draggedIds)
                siblings.RemoveAll(i => i.Id == did);

            int insertIndex = siblings.FindIndex(i => i.Id == targetId);
            if (insertIndex < 0) insertIndex = siblings.Count;
            foreach (var did in draggedIds)
            {
                var item = items.FirstOrDefault(i => i.Id == did);
                if (item != null)
                {
                    item.ParentId = parentId;
                    siblings.Insert(insertIndex++, item);
                }
            }
            for (int i = 0; i < siblings.Count; i++) siblings[i].SortOrder = i;

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

            grid.CurrentCell = grid.Rows[rowIndex].Cells[nameColIndex];
            grid.BeginEdit(true);
        }
        private void Delete_Click(object sender, EventArgs e)
        {
            BtnDelete_Click(sender, e);
        }
        private BookmarkItem GetSelectedItem()
        {
            if (grid.SelectedRows.Count == 0)
                return null;

            var selectedRow = grid.SelectedRows[0];
            var id = selectedRow.Cells["Id"].Value?.ToString();
            return items.FirstOrDefault(i => i.Id == id);
        }
        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == 0x84) // WM_NCHITTEST
            {
                Point pos = PointToClient(Cursor.Position);
                int resizeDir = 0;
                if (pos.X < ResizeBorder) resizeDir |= 1;
                else if (pos.X > Width - ResizeBorder) resizeDir |= 2;
                if (pos.Y < ResizeBorder) resizeDir |= 4;
                else if (pos.Y > Height - ResizeBorder) resizeDir |= 8;

                if (resizeDir != 0)
                {
                    switch (resizeDir)
                    {
                        case 5: m.Result = (IntPtr)13; break; // top-left
                        case 6: m.Result = (IntPtr)14; break; // top-right
                        case 9: m.Result = (IntPtr)16; break; // bottom-left
                        case 10: m.Result = (IntPtr)17; break; // bottom-right
                        case 1: m.Result = (IntPtr)10; break; // left
                        case 2: m.Result = (IntPtr)11; break; // right
                        case 4: m.Result = (IntPtr)12; break; // top
                        case 8: m.Result = (IntPtr)15; break; // bottom
                        default: m.Result = (IntPtr)0; break;
                    }
                }
            }
        }
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
    }
}