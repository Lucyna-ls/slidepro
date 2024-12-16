using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace PowerPointAddIn2
{
    public partial class CustomRibbon
    {
        private void CustomRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnShowTaskPane_Click_1(object sender, RibbonControlEventArgs e)
        {

            try { 

            foreach (var pane in Globals.ThisAddIn.CustomTaskPanes)
            {
                if (pane.Title == "My Task Pane")
                {
                    pane.Visible = !pane.Visible;
                    return;
                }
            }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }

            // If the Task Pane is not found, create and show it
            var myTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new UserControl1(), "My Task Pane");
            myTaskPane.Visible = true;
        }
    }

}
