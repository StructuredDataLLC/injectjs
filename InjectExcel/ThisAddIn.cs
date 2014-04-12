using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.IO;

using System.Runtime.InteropServices;

namespace InjectExcel
{
    /**
     * behavior of 2013 and 2010/2007 is different - 13 uses multiple
     * windows, while 17 use a single window for all documents.  
     */
    public partial class ThisAddIn
    {
        private scriptPane _scriptPane = null;
        private Microsoft.Office.Tools.CustomTaskPane pane = null;
        private bool E13 = false;

        private Dictionary<long, Microsoft.Office.Tools.CustomTaskPane> multipane_list;

        string xmlpath = null;

        public String ConfigPath
        {
            get { return xmlpath; }
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get { return pane; }
        }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            double version = 0;
            double.TryParse(Application.Version, out version);

            xmlpath = System.IO.Path.GetTempFileName();
            System.IO.StreamWriter sw = new StreamWriter(xmlpath);
            sw.Write(InjectExcel.Properties.Resources.js);
            sw.Flush();
            sw.Close();

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

                Application.WorkbookBeforeClose += Application_WorkbookBeforeClose17;
                Application.WorkbookActivate += Application_WorkbookActivate;
                Application.WorkbookDeactivate += Application_WorkbookDeactivate;

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
                ctp.VisibleChanged += new EventHandler(pane_VisibleChanged);
                ctp.Visible = InjectExcel.Properties.Settings.Default.visible;

            }
        }

        public void SetTaskPaneVisible ( bool visible, bool all = false )
        {
            if( E13 )
            {
                if (all)
                {
                    foreach (KeyValuePair<long, Microsoft.Office.Tools.CustomTaskPane> entry in multipane_list)
                    {
                        entry.Value.Visible = visible;
                    }
                }
                else
                {
                        int ID = Application.Hwnd;
                        if (multipane_list.ContainsKey(ID)) multipane_list[ID].Visible = visible;
                }
            }
            else
            {
                TaskPane.Visible = visible;
            }
        }

        void Application_WorkbookDeactivate(Excel.Workbook Wb)
        {
            _scriptPane.SaveInstanceState(Wb);

        }

        void Application_WorkbookActivate(Excel.Workbook Wb)
        {
            // for 17, we can (?) use the dispatch pointer to the wb
            // as a unique identifier, that's probably the only 
            // thing that will work...

            // HEADS UP: this is actually a good universal mechanism,
            // should have thought of it before.  survives saves
            // (better than .Name) and has open/close lifecycle.

            // NOTE: this calls addref, see
            // http://msdn.microsoft.com/en-us/library/system.runtime.interopservices.marshal.getidispatchforobject(v=vs.100).aspx

            IntPtr disp = Marshal.GetIDispatchForObject(Wb);
            long ID = (long)disp;
            Marshal.Release(disp);

            // Console.WriteLine("WBA17: " + ((long)disp));

            _scriptPane.SetInstance(ID);

        }

        void Application_WorkbookBeforeClose17(Excel.Workbook Wb, ref bool Cancel)
        {
            // see above

            IntPtr disp = Marshal.GetIDispatchForObject(Wb);
            long ID = (long)disp;
            Marshal.Release(disp);


            Console.WriteLine("WBC17");
        }

        void Application_WorkbookBeforeClose(Excel.Workbook Wb, ref bool Cancel)
        {
            int hwnd = Application.Hwnd;
            if (multipane_list.ContainsKey(hwnd))
            {
                this.CustomTaskPanes.Remove(multipane_list[hwnd]); // dispose?
                multipane_list.Remove(hwnd);

                // TODO: resource cleanup for instance?

            }
        }

        void Application_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
        {
            // this is sufficient for E13 - we'll get any new/opened 
            // documents via window activate (except for the first one,
            // which we've already handled).

            // TODO: for E13, size and open/close prefs

            // FIXME: this works, but could harmonize on disp ptrs, if desired.

            int hwnd = Application.Hwnd;
            if( !multipane_list.ContainsKey( hwnd ))
            {
                scriptPane sp = new scriptPane();
                Microsoft.Office.Tools.CustomTaskPane ctp = this.CustomTaskPanes.Add(sp, "Inject-js", Application.ActiveWindow);
                multipane_list.Add(Application.Hwnd, ctp);
                ctp.VisibleChanged += new EventHandler(pane_VisibleChanged);
                ctp.Visible = InjectExcel.Properties.Settings.Default.visible;
            }
        }

        private void pane_VisibleChanged(object sender, System.EventArgs e)
        {

            Microsoft.Office.Tools.CustomTaskPane ctp = (Microsoft.Office.Tools.CustomTaskPane)sender;
            scriptPane sp = (scriptPane)ctp.Control;
            
            // 13... ?

            // Globals.Ribbons.Ribbon1.tbTaskPane.Checked = ctp.Visible;
            InjectExcel.Properties.Settings.Default.visible = ctp.Visible;

            if (ctp.Visible && !sp.trackWidth)
            {
                System.Windows.Forms.Timer t = new System.Windows.Forms.Timer();
                t.Interval = 100;
                t.Tick += new EventHandler((obj, ev) =>
                {
                    t.Stop();
                    t.Enabled = false;
                    t.Dispose();

                    ctp.Width = InjectExcel.Properties.Settings.Default.width;
                    sp.trackWidth = true;
                });
                t.Start();
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            InjectExcel.Properties.Settings.Default.Save();
            File.Delete(xmlpath);
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
