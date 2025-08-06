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
        private string notebookName;
        private string notebookColor;
        private string sectionGroupName;
        private string sectionName;
        private string sectionColor;
        private string pageName;
        private string paraContent;
        private Label labelNotebook, labelSection, labelPage, labelPara;

        // New: list to hold both folders and bookmarks
        private List<BookmarkItem> items = new List<BookmarkItem>();

        // New: To track unique ids / folders/bookmarks - simple guid generator usage
        private int idCounter = 1;

        // ContextMenu
        private ContextMenuStrip contextMenu;

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
                // Initialize all labels (you should create and position them as you wish, example below)
                label = new Label { Location = new Point(20, 50), AutoSize = true };
                labelNotebook = new Label { Location = new Point(20, 30), AutoSize = true };
                labelSection = new Label { Location = new Point(20, 70), AutoSize = true };
                labelPage = new Label { Location = new Point(20, 90), AutoSize = true };
                labelPara = new Label { Location = new Point(20, 110), AutoSize = true };

                // Window base settings
                this.FormBorderStyle = FormBorderStyle.None;
                this.Width = 600;
                this.Height = 400;
                this.TopMost = true;
                this.BackColor = Color.White;
                this.Padding = new Padding(1);

                // Rounded corners and border
                this.Paint += (s, e) =>
                {
                    using (Pen pen = new Pen(Color.Black, 1))
                    using (Graphics g = e.Graphics)
                    {
                        g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                        g.DrawRoundedRectangle(pen, 0, 0, this.Width - 1, this.Height - 1, 12);
                    }
                };

                comboScope = new ComboBox() { Location = new Point(90, 12), Width = 120 };
                comboScope.Items.AddRange(new string[] { "Paragraph", "Page", "Section", "Notebook" });
                comboScope.SelectedItem = onenoteScope;
                btnSave = new Button() { Location = new Point(220, 11), Text = "Save", Width = 90 };
                btnSave.Click += BtnSave_Click;
                btnDelete = new Button() { Location = new Point(320, 11), Text = "Delete", Width = 90 };
                btnDelete.Click += BtnDelete_Click;
                grid = new DataGridView()
                {
                    Location = new Point(20, 130),
                    Width = 550,
                    Height = 220,
                    ReadOnly = true,
                    AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                    SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                    MultiSelect = false
                };
                grid.CellDoubleClick += Grid_CellDoubleClick;

                // Add context menu for folder/bookmark operations
                contextMenu = new ContextMenuStrip();
                contextMenu.Items.Add("New Folder", null, NewFolder_Click);
                contextMenu.Items.Add("Rename", null, Rename_Click);
                contextMenu.Items.Add("Delete", null, Delete_Click);

                grid.MouseDown += Grid_MouseDown;

                Controls.Add(label);
                Controls.Add(comboScope);
                Controls.Add(btnSave);
                Controls.Add(btnDelete);
                Controls.Add(grid);
                Controls.Add(labelNotebook);
                Controls.Add(labelSection);
                Controls.Add(labelPage);
                Controls.Add(labelPara);

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

                // Enable drag-and-drop
                grid.AllowDrop = true;
                grid.MouseMove += Grid_MouseMove;
                grid.DragOver += Grid_DragOver;
                grid.DragDrop += Grid_DragDrop;
                grid.MouseDown += Grid_MouseDown_StartDrag;

                Application.AddMessageFilter(new CustomMessageFilter(this));

                LoadTable();

                UpdateBookmarkInfo(
                    selectedId,
                    selectedScope,
                    selectedText,
                    notebookName,
                    notebookColor,
                    sectionGroupName,
                    sectionName,
                    sectionColor,
                    pageName,
                    paraContent);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error initializing Bookmark window: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #region
        private bool IsDescendant(string sourceId, string potentialAncestorId)
        {
            var item = items.FirstOrDefault(i => i.Id == sourceId);
            while (item != null && item.ParentId != null)
            {
                if (item.ParentId == potentialAncestorId)
                    return true;
                item = items.FirstOrDefault(i => i.Id == item.ParentId);
            }
            return false;
        }

        private void Grid_DragDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(typeof(string)))
            {
                string draggedId = (string)e.Data.GetData(typeof(string));

                // Figure out what was dropped on
                Point clientPoint = grid.PointToClient(new Point(e.X, e.Y));
                var hitTest = grid.HitTest(clientPoint.X, clientPoint.Y);

                if (hitTest.RowIndex >= 0)
                {
                    string targetId = grid.Rows[hitTest.RowIndex].Cells["Id"].Value?.ToString();
                    var draggedItem = items.FirstOrDefault(i => i.Id == draggedId);
                    var targetItem = items.FirstOrDefault(i => i.Id == targetId);

                    if (draggedItem == null || targetItem == null)
                        return;

                    // Only allow dropping onto folders.
                    if (targetItem.Type != "Folder")
                    {
                        MessageBox.Show("You may only move items into folders.");
                        return;
                    }

                    // Prevent invalid moves
                    if (draggedItem.Id == targetItem.Id ||
                        IsDescendant(draggedItem.Id, targetItem.Id))
                    {
                        MessageBox.Show("Cannot move a folder into itself or its descendant.");
                        return;
                    }

                    // Perform move
                    draggedItem.ParentId = targetItem.Id;
                    SaveToFile();
                    RefreshGridDisplay();
                }
                else
                {
                    // Drop to empty space = move to root
                    var draggedItem = items.FirstOrDefault(i => i.Id == draggedId);
                    if (draggedItem != null)
                    {
                        draggedItem.ParentId = null;
                        SaveToFile();
                        RefreshGridDisplay();
                    }
                }
            }
        }

        private void Grid_DragOver(object sender, DragEventArgs e)
        {
            e.Effect = DragDropEffects.Move;
        }

        private void Grid_MouseMove(object sender, MouseEventArgs e)
        {
            if ((e.Button & MouseButtons.Left) == MouseButtons.Left)
            {
                // Only start on significant movement
                if (Math.Abs(e.X - dragStart.X) > SystemInformation.DragSize.Width ||
                    Math.Abs(e.Y - dragStart.Y) > SystemInformation.DragSize.Height)
                {
                    if (grid.SelectedRows.Count > 0)
                    {
                        var row = grid.SelectedRows[0];
                        var dragData = row.Cells["Id"].Value?.ToString();
                        if (!string.IsNullOrEmpty(dragData))
                        {
                            grid.DoDragDrop(dragData, DragDropEffects.Move);
                        }
                    }
                }
            }
        }
        private Point dragStart;

        private void Grid_MouseDown_StartDrag(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
                dragStart = new Point(e.X, e.Y);
        }



        #endregion

        // Data class representing folder or bookmark
        private class BookmarkItem
        {
            public string Type { get; set; } // "Folder" or "Bookmark"
            public string Name { get; set; }
            public string ParentId { get; set; } // null or string for root level
            public string Id { get; set; }
            public string NotebookName { get; set; }
            public string NotebookColor { get; set; }
            public string SectionGroupName { get; set; }
            public string SectionName { get; set; }
            public string SectionColor { get; set; }
            public string PageName { get; set; }
            public string ParaContent { get; set; }
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

            RefreshGridDisplay();

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

                // Add as Bookmark under root (or you may select a folder - here we put under root for simplicity)
                var newBookmark = new BookmarkItem
                {
                    Type = "Bookmark",
                    Name = selectedText ?? "Unnamed Bookmark",
                    ParentId = null, // top level by default; extend if you want to assign selected folder
                    Id = selectedId,
                    NotebookName = notebookName,
                    NotebookColor = notebookColor,
                    SectionGroupName = sectionGroupName,
                    SectionName = sectionName,
                    SectionColor = sectionColor,
                    PageName = pageName,
                    ParaContent = paraContent
                };

                // Avoid duplicates: Remove existing bookmark with same Id
                items.RemoveAll(i => i.Type == "Bookmark" && i.Id == newBookmark.Id);
                items.Add(newBookmark);

                SaveToFile();
                RefreshGridDisplay();

                MessageBox.Show("Saved!");
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

            // Remove item and all children if folder
            RemoveItemAndChildren(itemId);

            SaveToFile();
            RefreshGridDisplay();

            MessageBox.Show("Deleted successfully.");
        }

        private void RemoveItemAndChildren(string id)
        {
            var toRemove = items.Where(i => i.Id == id).ToList();

            foreach (var item in toRemove)
            {
                items.Remove(item);
                if (item.Type == "Folder")
                {
                    // Recursively remove children
                    var children = items.Where(c => c.ParentId == item.Id).Select(c => c.Id).ToList();
                    foreach (var childId in children)
                        RemoveItemAndChildren(childId);
                }
            }
        }

        private void LoadTable()
        {
            items.Clear();

            if (!File.Exists(tablePath))
                return;

            try
            {
                var lines = File.ReadAllLines(tablePath);

                foreach (var line in lines)
                {
                    // Format: Type,Id,ParentId,Name,NotebookName,NotebookColor,SectionGroupName,SectionName,SectionColor,PageName,ParaContent
                    var parts = line.Split(new[] { ',' }, 11);
                    if (parts.Length == 11)
                    {
                        items.Add(new BookmarkItem
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
                            ParaContent = parts[10]
                        });
                    }
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
                var lines = items.Select(i =>
                    string.Join(",", new[]{
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
                        EscapeCsv(i.ParaContent)
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
            {
                return "\"" + input.Replace("\"", "\"\"") + "\"";
            }
            return input;
        }

        private void RefreshGridDisplay()
        {
            grid.Columns.Clear();
            grid.Rows.Clear();

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

            // Flatten the hierarchy with indentation
            var flatList = FlattenForDisplay(null, 0);

            foreach (var item in flatList)
            {
                grid.Rows.Add(
                    item.Type,
                    IndentName(item.Name, GetDepth(item)),
                    item.Id,
                    item.NotebookName,
                    item.NotebookColor,
                    item.SectionGroupName,
                    item.SectionName,
                    item.SectionColor,
                    item.PageName,
                    item.ParaContent);
            }
        }

        private List<BookmarkItem> FlattenForDisplay(string parentId, int depth)
        {
            var result = new List<BookmarkItem>();

            // Get folders first
            var folders = items.Where(i => i.ParentId == parentId && i.Type == "Folder")
                               .OrderBy(i => i.Name).ToList();

            foreach (var folder in folders)
            {
                result.Add(folder);
                result.AddRange(FlattenForDisplay(folder.Id, depth + 1));
            }

            // Then add bookmarks under this parent
            var bookmarks = items.Where(i => i.ParentId == parentId && i.Type == "Bookmark")
                                 .OrderBy(i => i.Name).ToList();

            result.AddRange(bookmarks);

            return result;
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

        private string IndentName(string name, int depth)
        {
            return new string(' ', depth * 6) + name;
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    string id = grid.Rows[e.RowIndex].Cells["Id"].Value.ToString();
                    var item = items.FirstOrDefault(i => i.Id == id);
                    if (item == null) return;

                    if (item.Type == "Bookmark")
                    {
                        // Navigate to the actual OneNote location (sample)
                        var app = new Microsoft.Office.Interop.OneNote.Application();
                        app.NavigateTo(id);
                    }
                    else if (item.Type == "Folder")
                    {
                        // Optional: implement collapse/expand logic by filtering items - for simplicity, no expand/collapse here
                        MessageBox.Show("Folder double-clicked: " + item.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening: " + ex.Message);
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

        // Context Menu Handlers

        private void NewFolder_Click(object sender, EventArgs e)
        {
            var currentRow = GetSelectedItem();
            string parentId = currentRow?.Id; // new folder under selected folder or null if selected is bookmark or none

            if (currentRow != null && currentRow.Type == "Bookmark")
            {
                // Parent is the parent of bookmark's folder, or null
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
            RefreshGridDisplay();
        }

        private void Rename_Click(object sender, EventArgs e)
        {
            var currentRow = GetSelectedItem();
            if (currentRow == null) return;

            string oldName = currentRow.Name;
            string prompt = $"Rename {(currentRow.Type == "Folder" ? "Folder" : "Bookmark")}";

            string newName = Prompt.ShowDialog(prompt, "Rename", oldName);
            if (string.IsNullOrEmpty(newName) || newName == oldName)
                return;

            // Update item name
            currentRow.Name = newName;
            SaveToFile();
            RefreshGridDisplay();
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

        // Helper for simple input dialog for rename (used in Rename_Click)
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

    public class CustomMessageFilter : IMessageFilter
    {
        private readonly Form _form;

        public CustomMessageFilter(Form form)
        {
            _form = form;
        }

        public bool PreFilterMessage(ref Message m)
        {
            const int WM_LBUTTONDOWN = 0x0201; // Left mouse button down
            const int WM_RBUTTONDOWN = 0x0204; // Right mouse button down

            if (m.Msg == WM_LBUTTONDOWN || m.Msg == WM_RBUTTONDOWN)
            {
                // Get mouse position in screen coordinates
                Point mousePos = Control.MousePosition;

                // If mouse click is outside the form bounds
                if (!_form.Bounds.Contains(mousePos))
                {
                    // Minimize the form immediately
                    _form.WindowState = FormWindowState.Minimized;
                }
            }

            return false;
        }
    }

    static class GraphicsExtension
    {
        public static void DrawRoundedRectangle(this Graphics g, Pen p, int x, int y, int width, int height, int radius)
        {
            using (var path = new System.Drawing.Drawing2D.GraphicsPath())
            {
                path.AddArc(x, y, radius, radius, 180, 90);
                path.AddArc(x + width - radius, y, radius, radius, 270, 90);
                path.AddArc(x + width - radius, y + height - radius, radius, radius, 0, 90);
                path.AddArc(x, y + height - radius, radius, radius, 90, 90);
                path.CloseAllFigures();
                g.DrawPath(p, path);
            }
        }
    }
}
