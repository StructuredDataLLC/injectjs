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
    /**
     * behavior of 2013 and 2010/2007 is different - 13 uses multiple
     * windows, while 17 use a single window for all documents.  
     * 
     * one question is how to determine the version?
     */

    //[PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
    public partial class ThisAddIn
    {
        private scriptPane _scriptPane = null;
        private Microsoft.Office.Tools.CustomTaskPane pane = null;
        private bool E13 = false;

        private Dictionary<long, Microsoft.Office.Tools.CustomTaskPane> multipane_list;

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get { return pane; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            double version = 0;
            double.TryParse(Application.Version, out version);

            if (version < 12.0) // not supported
            {
                return; 
            }
            else if (version < 15.0) // 2007, 2010
            {
                _scriptPane = new scriptPane();
                pane = this.CustomTaskPanes.Add(_scriptPane, "Inject-js");
                pane.VisibleChanged += new EventHandler(pane_VisibleChanged);
                pane.Visible = InjectExcel.Properties.Settings.Default.visible;
            }
            else // 2013
            {
                E13 = true;
                multipane_list = new Dictionary<long, Microsoft.Office.Tools.CustomTaskPane>();
                Application.WindowActivate += Application_WindowActivate;
                Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;
                scriptPane sp = new scriptPane();
                Microsoft.Office.Tools.CustomTaskPane ctp = this.CustomTaskPanes.Add(sp, "Inject-js", Application.ActiveWindow);
                multipane_list.Add(Application.Hwnd, ctp);
                ctp.Visible = true;

            }
        }

        void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            int hwnd = Application.Hwnd;
            if (multipane_list.ContainsKey(hwnd))
            {
                this.CustomTaskPanes.Remove(multipane_list[hwnd]);
                multipane_list.Remove(hwnd);
            }
        }

        void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            int hwnd = Application.Hwnd;
            if( !multipane_list.ContainsKey( hwnd ))
            {
                scriptPane sp = new scriptPane();
                Microsoft.Office.Tools.CustomTaskPane ctp = this.CustomTaskPanes.Add(sp, "Inject-js", Application.ActiveWindow);
                multipane_list.Add(Application.Hwnd, ctp);
                ctp.Visible = true;
            }
        }

        void Application_WorkbookOpen(Excel.Workbook Wb)
        {
            int hwnd = Application.Hwnd;

            Console.WriteLine("WO" + hwnd);
            //throw new NotImplementedException();
        }

        private void pane_VisibleChanged(object sender, System.EventArgs e)
        {
            Globals.Ribbons.Ribbon1.tbTaskPane.Checked = pane.Visible;
            InjectExcel.Properties.Settings.Default.visible = pane.Visible;

            if (pane.Visible && !_scriptPane.trackWidth)
            {
                System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
                t.Interval = 100;
                t.Tick += new EventHandler((obj, ev) =>
                {
                    TaskPane.Width = InjectExcel.Properties.Settings.Default.width;
                    _scriptPane.trackWidth = true;
                    t.Stop();
                    t.Enabled = false;
                    t.Dispose();
                });
                t.Start();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            InjectExcel.Properties.Settings.Default.Save();
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
