using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn2
{
    public partial class ThisAddIn
    {
        private CustomTaskPane myCustomTaskPane;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            UserControl1 myUserControl = new UserControl1();

            // Add the UserControl to the Custom Task Panes collection
            myCustomTaskPane = this.CustomTaskPanes.Add(myUserControl, "Designer");

            // Dock the task pane to the right of the PowerPoint window
            myCustomTaskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;

            // Set the visibility of the task pane
            myCustomTaskPane.Visible = true;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        #endregion
    }
}
