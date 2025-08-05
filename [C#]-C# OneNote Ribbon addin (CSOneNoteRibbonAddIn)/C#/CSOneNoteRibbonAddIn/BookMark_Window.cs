using System;
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
        private Button btnDelete;    // Delete button
        private DataGridView grid;
        private string selectedId;
        private string selectedScope;
        private string selectedText;
        private string tablePath;

        public BookMark_Window(string onenoteId, string onenoteScope, string displayText)
        {
            this.Text = "Bookmark Saver";
            this.Width = 600;
            this.Height = 300;
            this.TopMost = true;

            label = new Label() { Location = new System.Drawing.Point(20, 16), AutoSize = true, Text = "Current:" };
            comboScope = new ComboBox() { Location = new System.Drawing.Point(90, 12), Width = 120 };
            comboScope.Items.AddRange(new string[] { "Paragraph", "Page", "Section", "Notebook" });
            comboScope.SelectedItem = onenoteScope;

            btnSave = new Button() { Location = new System.Drawing.Point(220, 11), Text = "Save", Width = 90 };
            btnSave.Click += BtnSave_Click;

            btnDelete = new Button() { Location = new System.Drawing.Point(320, 11), Text = "Delete", Width = 90 };
            btnDelete.Click += BtnDelete_Click;

            grid = new DataGridView()
            {
                Location = new System.Drawing.Point(20, 50),
                Width = 550,
                Height = 180,
                ReadOnly = true,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                MultiSelect = false
            };
            grid.CellDoubleClick += Grid_CellDoubleClick;

            Controls.Add(label);
            Controls.Add(comboScope);
            Controls.Add(btnSave);
            Controls.Add(btnDelete);
            Controls.Add(grid);

            selectedId = onenoteId;
            selectedScope = onenoteScope;
            selectedText = displayText;
            tablePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "bookmarks.txt");

            LoadTable();
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            try
            {
                string row = $"{comboScope.SelectedItem},{selectedId},{selectedText.Replace(',', ' ')}";
                File.AppendAllLines(tablePath, new[] { row });
                LoadTable();
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

            try
            {
                var selectedRow = grid.SelectedRows[0];
                string scope = selectedRow.Cells[0].Value?.ToString();
                string id = selectedRow.Cells[1].Value?.ToString();
                string text = selectedRow.Cells[2].Value?.ToString();

                if (string.IsNullOrEmpty(id))
                {
                    MessageBox.Show("Selected row is invalid.");
                    return;
                }

                if (!File.Exists(tablePath))
                {
                    MessageBox.Show("Bookmark file not found.");
                    return;
                }

                var lines = File.ReadAllLines(tablePath).ToList();

                // Remove matching line(s)
                lines = lines.Where(line =>
                {
                    var parts = line.Split(',');
                    return !(parts.Length == 3
                             && parts[0] == scope
                             && parts[1] == id
                             && parts[2] == text);
                }).ToList();

                File.WriteAllLines(tablePath, lines);

                LoadTable();
                MessageBox.Show("Deleted successfully.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error deleting: " + ex.Message);
            }
        }

        private void LoadTable()
        {
            grid.Columns.Clear();
            grid.Rows.Clear();
            grid.Columns.Add("Scope", "Scope");
            grid.Columns.Add("Id", "OneNote Id");
            grid.Columns.Add("Text", "Text");

            if (File.Exists(tablePath))
            {
                foreach (var line in File.ReadAllLines(tablePath))
                {
                    var parts = line.Split(',');
                    if (parts.Length == 3)
                        grid.Rows.Add(parts);
                }
            }
        }

        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex >= 0)
                {
                    string id = grid.Rows[e.RowIndex].Cells[1].Value.ToString();
                    var app = new Microsoft.Office.Interop.OneNote.Application();
                    app.NavigateTo(id);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error opening: " + ex.Message);
            }
        }
    }
}
