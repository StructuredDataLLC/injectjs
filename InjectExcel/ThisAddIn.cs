using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using System.Reflection;
using System.IO;
using System.Security.Permissions;

namespace InjectExcel
{
    [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public partial class ThisAddIn
    {
        private Microsoft.Office.Tools.CustomTaskPane pane;

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return pane;
            }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            scriptPane ctl = new scriptPane();
            pane = this.CustomTaskPanes.Add(ctl, "Injectjs");
            pane.Visible = true;
            pane.VisibleChanged += new EventHandler(pane_VisibleChanged);
        }

        private void pane_VisibleChanged(object sender, System.EventArgs e)
        {
            // Globals.Ribbons.Ribbon1.toggleButton1.Checked = pane.Visible;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
