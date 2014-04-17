using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace InjectExcel
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            // e.RibbonUI
        }

        private void tbTaskPane_Click(object sender, RibbonControlEventArgs e)
        {
            // Globals.ThisAddIn.TaskPane.Visible = ((RibbonToggleButton)sender).Checked;
            Globals.ThisAddIn.SetTaskPaneVisible(((RibbonToggleButton)sender).Checked);
        }

        private void bShow_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SetTaskPaneVisible(true);
        }

        private void bHide_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.SetTaskPaneVisible(false);
        }


    }
}
