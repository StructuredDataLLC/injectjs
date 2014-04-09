using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;

using ScintillaNET;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;

using CV82Lib;

using System.Runtime.InteropServices;

namespace InjectExcel
{
    public partial class scriptPane : UserControl
    {
        Scintilla editor = null;
        Scintilla logger = null;

        const string SCRIPT_XML_NAMESPACE = "http://tempuri.org/inject/script";

        Scripto scripto = null;

        public scriptPane()
        {
            InitializeComponent();

            string codeBase = Assembly.GetExecutingAssembly().CodeBase;
            UriBuilder uri = new UriBuilder(codeBase);
            string path = Uri.UnescapeDataString(uri.Path);
            string dir = Path.GetDirectoryName(path);

            string cwd = Directory.GetCurrentDirectory();
            Directory.SetCurrentDirectory(dir);

            editor = new Scintilla();
            SwapControls(editor, tbEditor);
            editor.ConfigurationManager.Language = "js";
            editor.Margins[0].Width = 20;

            logger = new Scintilla();
            SwapControls(logger, tbLog);
            logger.Font = tbLog.Font;

            initContext();

            Globals.ThisAddIn.Application.WorkbookBeforeSave += app_WorkbookBeforeSave;
            Directory.SetCurrentDirectory(cwd);

            editor.TextChanged += editor_TextChanged;

            bClearLog.Click += bClearLog_Click;
            bExecute.Click += bExec_Click;

            Load += taskPane_Load;
        }

        void bClearLog_Click(object sender, EventArgs e)
        {
            throw new NotImplementedException();
        }

        void taskPane_Load(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook == null) return;
            editor.TextChanged -= editor_TextChanged;
            LoadScript();
            editor.TextChanged += editor_TextChanged;
        }

        void app_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            SaveScript();
        }

        void editor_TextChanged(object sender, EventArgs e)
        {
            if (null != Globals.ThisAddIn.Application.ActiveWorkbook) Globals.ThisAddIn.Application.ActiveWorkbook.Saved = false;
        }

        void app_WorkbookOpen(Excel.Workbook Wb)
        {
            editor.TextChanged -= editor_TextChanged;
            LoadScript();
            editor.TextChanged += editor_TextChanged;
        }

        public void SwapControls(Control newCtl, Control oldCtl)
        {
            oldCtl.Parent.Controls.Add(newCtl);
            newCtl.Size = oldCtl.Size;
            newCtl.Top = oldCtl.Top;
            newCtl.Left = oldCtl.Left;
            newCtl.Anchor = oldCtl.Anchor;
            oldCtl.Parent.Controls.Remove(oldCtl);
        }
        
        void initContext()
        {
            scripto = new Scripto();
            scripto.SetDispatch(Globals.ThisAddIn.Application, "Application");
            scripto.OnConsolePrint += scripto_OnConsolePrint;
            scripto.OnAlert += scripto_OnAlert;
            scripto.OnConfirm += scripto_OnConfirm;
        }

        void scripto_OnConfirm(ref string Msg, out bool Rslt)
        {
            object oRslt = Invoke(new Func<String, bool>((msg) =>
            {
                return (DialogResult.OK == MessageBox.Show(msg, "Injectjs Script Alert", MessageBoxButtons.OKCancel, MessageBoxIcon.Question));
            }), new object[] { Msg });

            Rslt = (Boolean)oRslt;
        }

        void scripto_OnAlert(ref string Msg)
        {
            Invoke(new Action<string>((msg) =>
            {
                MessageBox.Show(msg, "Injectjs Script Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }), new object[] { Msg });
        }

        void scripto_OnConsolePrint(ref string Msg, int Flags)
        {
            logMessage(Flags + ": " + Msg);
        }

        void logMessage(string message)
        {
            Invoke(new Action<string>((msg) =>
            {
                logger.AppendText(msg + "\r\n");
                logger.Scrolling.ScrollToLine(logger.Lines.Count);

            }), new object[] { message });
        }

        public void squiggle(int line)
        {
            --line;

            int start = editor.Lines[line].StartPosition;
            int end = editor.Lines[line].EndPosition;

            Range r = editor.GetRange(start, end);
            r.SetIndicator(2);

        }
        public void unsquiggle()
        {
            editor.NativeInterface.IndicatorClearRange(0, editor.Text.Length);
        }

        private void bExec_Click(object sender, EventArgs e)
        {
            if (null == scripto) initContext();

            try
            {
                unsquiggle();
                string script = editor.Text;
                string output;
                bool rslt = scripto.ExecString(script, out output);
                logMessage(output);
                if (!rslt)
                {
                    Regex rex = new Regex(@"^.+?\:(\d+)\:");
                    Match m = rex.Match(output);
                    if (m.Success)
                    {
                        squiggle(Int32.Parse(m.Groups[1].Value));
                    }
                }
            }
            catch (Exception x)
            {
                logMessage("Exception: " + x.Message);
            }

        }

         void SaveScript()
        {
            if (null == Globals.ThisAddIn.Application.ActiveWorkbook) return;
            string xmlString1 =
                      "<script xmlns=\"" + SCRIPT_XML_NAMESPACE + "\"><source><![CDATA[" + editor.Text +
                      "]]></source></script>";

            Office.CustomXMLParts parts = Globals.ThisAddIn.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace(SCRIPT_XML_NAMESPACE);
            if (parts.Count > 0)
            {
                Office.CustomXMLPart part = parts[1];
                part.NamespaceManager.AddNamespace("N", SCRIPT_XML_NAMESPACE);
                Office.CustomXMLNode node = part.SelectSingleNode("//N:source"); 
                if (null != node)
                {
                    node.FirstChild.Delete();
                    node.AppendChildNode("", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeCData, editor.Text);
                }
            }
            else
            {
                Globals.ThisAddIn.Application.ActiveWorkbook.CustomXMLParts.Add(xmlString1);
            }

        }

        void LoadScript()
        {
            if (null == Globals.ThisAddIn.Application.ActiveWorkbook) return;
            Office.CustomXMLParts parts = Globals.ThisAddIn.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace(SCRIPT_XML_NAMESPACE);
            if (parts.Count > 0)
            {
                Office.CustomXMLPart part = parts[1];
                part.NamespaceManager.AddNamespace("N", SCRIPT_XML_NAMESPACE);

                Office.CustomXMLNode node = part.SelectSingleNode("//N:source"); // .FirstChild;
                if (null != node) editor.Text = node.FirstChild.NodeValue.ToString();
            }
        }

        private void bClearLog_Click_1(object sender, EventArgs e)
        {
            logger.Text = "";
        }
    }
}
