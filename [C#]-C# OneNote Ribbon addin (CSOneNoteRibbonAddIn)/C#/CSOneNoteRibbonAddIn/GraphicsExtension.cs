using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CSOneNoteRibbonAddIn
{
    public static class GraphicsExtension
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

    public class CustomMessageFilter : IMessageFilter
    {
        private readonly Form _form;

        public CustomMessageFilter(Form form)
        {
            _form = form;
        }

        public bool PreFilterMessage(ref Message m)
        {
            const int WM_LBUTTONDOWN = 0x0201;
            const int WM_RBUTTONDOWN = 0x0204;

            if (m.Msg == WM_LBUTTONDOWN || m.Msg == WM_RBUTTONDOWN)
            {
                Point mousePos = Control.MousePosition;
                if (!_form.Bounds.Contains(mousePos))
                {
                    _form.WindowState = FormWindowState.Minimized;
                }
            }
            return false;
        }
    }
}
