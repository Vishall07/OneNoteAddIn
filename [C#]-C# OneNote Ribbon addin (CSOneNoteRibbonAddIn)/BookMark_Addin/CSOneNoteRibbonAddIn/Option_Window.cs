using System;
using System.Drawing;
using System.Windows.Forms;

namespace CSOneNoteRibbonAddIn
{
    public partial class Option_Window : Form
    {
       
        public Option_Window()
        {
            InitializeComponent();
            // Thin border window style
            this.FormBorderStyle = FormBorderStyle.FixedSingle;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = Color.White;
            this.ForeColor = Color.Black;
            this.Font = new Font("Segoe UI", 10);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.FormBorderStyle = FormBorderStyle.None;
            this.Size = new Size(130, 90);
  
        }
    }

}
