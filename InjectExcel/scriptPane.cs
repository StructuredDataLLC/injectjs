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
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;

using System.Web.Script.Serialization;

using CV82Lib;


namespace InjectExcel
{
    public partial class scriptPane : UserControl
    {
        Scintilla editor = null;
        Scintilla logger = null;

        const string SCRIPT_XML_NAMESPACE = "http://tempuri.org/inject/script";

        // Scripto scripto = null;

        public bool trackWidth = false;

        bool cinit = false;

        Dictionary<string, object> masterDict = null;

        Dictionary<long, InstanceData> slist = new Dictionary<long,InstanceData>();

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

            // initContext();

            Globals.ThisAddIn.Application.WorkbookBeforeSave += app_WorkbookBeforeSave;
            Directory.SetCurrentDirectory(cwd);

            editor.TextChanged += editor_TextChanged;
            editor.KeyPress += editor_KeyPress;

            editor.AutoComplete.List.Clear();
            editor.AutoComplete.List.Add("Application");
            editor.AutoComplete.List.Sort();
            //editor.AutoComplete.IsCaseSensitive = true;

            bClearLog.Click += bClearLog_Click;
            bExecute.Click += bExec_Click;

            cbScriptLanguage.Items.Add("Javascript");
            cbScriptLanguage.Items.Add("CoffeeScript");
            cbScriptLanguage.SelectedIndex = 0;

            cbScriptLanguage.SelectedIndexChanged += cbScriptLanguage_SelectedIndexChanged;

            Load += taskPane_Load;
            SizeChanged += scriptPane_SizeChanged;
        }

        /**
         * Excel 2007/2010 use a single window, therefore a single task pane,
         * for multiple documents.  we can support this by using separate instances
         * of the scripting environment per window.  we'll use the wb ID to 
         * identify them.
         * 
         * For 2013 this is not necessary b/c it uses separate windows, and separate
         * task panes, per doc.
         */
        public void SetInstance(long instanceID)
        {
            // we don't actually have to worry about script yet,
            // but update the code window

            IntPtr disp = Marshal.GetIDispatchForObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            long ID = (long)disp;
            Marshal.Release(disp);

            editor.TextChanged -= editor_TextChanged;
            LoadScript();
            editor.TextChanged += editor_TextChanged;

            // FIXME: global/local dict...

            // UPDATE: also the log, which we're preserving per-instance

            if (!slist.ContainsKey(ID)) logger.Text = "";
            else logger.Text = slist[ID].log;
            logger.Scrolling.ScrollToLine(logger.Lines.Count);

        }

        public void SaveInstanceState(Excel.Workbook Wb)
        {
            if( null != Wb )
            {
                // FIXME (MAYBE): use a memory buffer instead?
                // but isn't the xml just in memory, until it's 
                // disk-saved?

                IntPtr disp = Marshal.GetIDispatchForObject(Wb);
                long ID = (long)disp;
                Marshal.Release(disp);

                SaveScript(Wb);

                if (!slist.ContainsKey(ID))
                {
                    slist.Add(ID, new InstanceData());
                }

                slist[ID].log = logger.Text;
            }
        }

        void scriptPane_SizeChanged(object sender, EventArgs e)
        {
            if (trackWidth && null != Globals.ThisAddIn.TaskPane)
                InjectExcel.Properties.Settings.Default.width = Globals.ThisAddIn.TaskPane.Width;
        }

        string findRootType( string token )
        {
            // TODO: types for fields in scope

            if (token.Equals("Application")) return "Application";
            if( masterDict.ContainsKey( token ))
            {
                Dictionary<string, object> d = (Dictionary<string, object>)masterDict[token];
                if (d["type"].Equals("enum")
                    || d["type"].Equals("user")) return token;
            }
            return null;
        }

        string childType( string root, string child )
        {
            if (!masterDict.ContainsKey(root)) return "";
            Dictionary<string, object> d = (Dictionary<string, object>)masterDict[root];

            // note: if root is an enum type, there's no way the child has candidates, 
            // you can just bail

            if (d["type"].Equals("enum")) return "";

            Dictionary<string, object> candidates = Candidates(root, null);
            if (!candidates.ContainsKey(child)) return "";

            Dictionary<string, object> dict = (Dictionary<string, object>)candidates[child];

            if (dict.ContainsKey("type")) return dict["type"].ToString();
            return "";
        }

        Dictionary<string, object> Candidates( string parent, string key )
        {
            // start by looking up the parent type
            if (!masterDict.ContainsKey(parent)) return new Dictionary<string, object>(); 

            Dictionary<string, object> d = (Dictionary<string, object>)masterDict[parent];

            if( d["type"].Equals("coclass"))
            {
                if( null == key ) // return all members of any interfaces
                {
                    Dictionary<string, object> dict = new Dictionary<string, object>();
                    if( d.ContainsKey("default") && masterDict.ContainsKey(d["default"].ToString()))
                    {
                        Dictionary<string, object> d2 = Candidates( d["default"].ToString(), null);
                        foreach (string k2 in d2.Keys) dict.Add(k2, d2[k2]);
                    }
                    if (d.ContainsKey("source") && masterDict.ContainsKey(d["source"].ToString()))
                    {
                        Dictionary<string, object> d2 = Candidates(d["source"].ToString(), null);
                        foreach (string k2 in d2.Keys) if( !dict.ContainsKey( k2 )) dict.Add(k2, d2[k2]);
                    }
                    return dict;
                }
            }
            else if( d["type"].Equals("enum"))
            {
                if( null == key )
                    return (Dictionary<string, object>)d["values"];
            }
            else // interface OR user
            {
                if( null == key ) // return all members of this interface
                    return (Dictionary<string, object>) d["members"];
            }

            return new Dictionary<string, object>();
        }

        void editor_KeyPress(object sender, KeyPressEventArgs e)
        {
            int col;
            string line = editor.GetCurrentLine(out col);
            List< string > list = new List<string>();

            if( e.KeyChar == 'l')
            {
                Regex rex = new Regex(@"(?:^|[\W\s])Xl*", RegexOptions.IgnoreCase);
                line = line.Substring(0, col);
                Match m = rex.Match(line);
                if (m.Success)
                {
                    foreach( string key in masterDict.Keys )
                    {
                        if (key.StartsWith("Xl")) list.Add(key);
                    }
                }
            }
            else if (e.KeyChar == '.')
            {
                line = line.Substring(0, col);
                Regex rex = new Regex( @"[\w\d\._\(\)]+$");
                Regex rexUnfunc = new Regex(@"\(\.*?(?:\)|$)");
                Match m = rex.Match(line);
                if( m.Success )
                {
                    string[] elts = m.Groups[0].ToString().Split(new char[]{'.'});
                    string elt = rexUnfunc.Replace(elts[0], "");

                    string root = findRootType(elts[0]);
                    if (null == root) return;

                    for( int i = 1; i< elts.Length; i++ )
                    {
                        // if elt is a legal child of root, then
                        // find elt's type and set that as root.

                        elt = rexUnfunc.Replace(elts[i], "");
                        root = childType(root, elt);
                    }

                    // if we have a type, find candidates

                    Dictionary<string, object> dict = Candidates(root, null);
                    foreach (string key in (dict.Keys)) list.Add(key);

                }
            }

            if( list.Count > 0 )
            {
                list.Sort();
                editor.AutoComplete.List = list;

                Timer t = new Timer();

                t.Interval = 100;
                t.Tag = editor;
                t.Tick += new EventHandler((obj, ev) =>
                {
                    editor.AutoComplete.Show();
                    t.Stop();
                    t.Enabled = false;
                    t.Dispose();
                });
                t.Start();
            }

        }

        void cbScriptLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (null == Globals.ThisAddIn.Application.ActiveWorkbook)
                Globals.ThisAddIn.Application.ActiveWorkbook.Saved = false;
        }

        void taskPane_Load(object sender, EventArgs e)
        {
            if (Globals.ThisAddIn.Application.ActiveWorkbook == null) return;
            editor.TextChanged -= editor_TextChanged;
            LoadScript();
            editor.TextChanged += editor_TextChanged;

            // initContext();

            ConstructTypeMap();

        }

        void ConstructTypeMap()
        {
            /*
            string map = scripto.MapTypeLib(Globals.ThisAddIn.Application);
            StreamWriter sw = new StreamWriter(@"e:\office-dev\excel.json");
            sw.Write(map);
            sw.Flush();
            sw.Close();
            */

            System.Web.Script.Serialization.JavaScriptSerializer serializer = new JavaScriptSerializer();
            masterDict = (Dictionary<string, object>)(serializer.Deserialize<Object>(InjectExcel.Properties.Resources.excel_interface));

            Dictionary<string, object> tmp = (Dictionary<string, object>)(serializer.Deserialize<Object>(InjectExcel.Properties.Resources.extra_functions));
            foreach( string key in tmp.Keys )
            {
                masterDict.Add(key, tmp[key]);
            }

        }

        void app_WorkbookBeforeSave(Excel.Workbook Wb, bool SaveAsUI, ref bool Cancel)
        {
            SaveScript();
        }

        void editor_TextChanged(object sender, EventArgs e)
        {
            if (null != Globals.ThisAddIn.Application.ActiveWorkbook) Globals.ThisAddIn.Application.ActiveWorkbook.Saved = false;

        }

        /*
        void app_WorkbookOpen(Excel.Workbook Wb)
        {
            editor.TextChanged -= editor_TextChanged;
            LoadScript();
            editor.TextChanged += editor_TextChanged;
        }
        */

        public void SwapControls(Control newCtl, Control oldCtl)
        {
            oldCtl.Parent.Controls.Add(newCtl);
            newCtl.Size = oldCtl.Size;
            newCtl.Top = oldCtl.Top;
            newCtl.Left = oldCtl.Left;
            newCtl.Anchor = oldCtl.Anchor;
            oldCtl.Parent.Controls.Remove(oldCtl);
        }
        
        Scripto ensureContext()
        {
            if (null == Globals.ThisAddIn.Application.ActiveWorkbook) return null;

            IntPtr disp = Marshal.GetIDispatchForObject(Globals.ThisAddIn.Application.ActiveWorkbook);
            long ID = (long)disp;
            Marshal.Release(disp);

            if (!slist.ContainsKey(ID)) slist.Add(ID, new InstanceData());
            if (null != slist[ID].scripto ) return slist[ID].scripto;

            slist[ID].scripto = new Scripto();
            slist[ID].scripto.SetDispatch(Globals.ThisAddIn.Application, "Application", true);
            // ConstructTypeMap();

            slist[ID].scripto.OnConsolePrint += scripto_OnConsolePrint;
            slist[ID].scripto.OnAlert += scripto_OnAlert;
            slist[ID].scripto.OnConfirm += scripto_OnConfirm;

            return slist[ID].scripto;
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

        /**
         * mark a line in the editor window
         */
        public void squiggle(int line)
        {
            --line;

            int start = editor.Lines[line].StartPosition;
            int end = editor.Lines[line].EndPosition;

            Range r = editor.GetRange(start, end);
            r.SetIndicator(2);

        }

        /**
         * clear editor marks 
         */
        public void unsquiggle()
        {
            editor.NativeInterface.IndicatorClearRange(0, editor.Text.Length);
        }

        // FIXME: crummy
        string jsescape( string s )
        {
            StringBuilder sb = new StringBuilder();
            sb.Append("\"");
            foreach (char c in s)
            {
                switch (c)
                {
                    case '\"':
                        sb.Append("\\\"");
                        break;
                    case '\\':
                        sb.Append("\\\\");
                        break;
                    case '\b':
                        sb.Append("\\b");
                        break;
                    case '\f':
                        sb.Append("\\f");
                        break;
                    case '\n':
                        sb.Append("\\n");
                        break;
                    case '\r':
                        sb.Append("\\r");
                        break;
                    case '\t':
                        sb.Append("\\t");
                        break;
                    default:
                        int i = (int)c;
                        if (i < 32 || i > 127)
                        {
                            sb.AppendFormat("\\u{0:X04}", i);
                        }
                        else
                        {
                            sb.Append(c);
                        }
                        break;
                }
            }
            sb.Append("\"");

            return sb.ToString();
        }

        /**
         * execute code in current buffer
         */
        private void bExec_Click(object sender, EventArgs e)
        {
            Scripto scripto = ensureContext();
            if (null == scripto)
            {
                logMessage("No context");
                return;
            }

            try
            {
                unsquiggle();
                string script = editor.Text;
                string output;
                bool rslt;

                // FIXME: this should be done closer to the 
                // script engine, and also, source maps.
                // with v8 we could compile this function once...

                if (cbScriptLanguage.SelectedIndex != 0)
                {
                    // this only has to happen once
                    if (!cinit) {
                        cinit = true;
                        bool xb = scripto.ExecString(InjectExcel.Properties.Resources.coffee_script, out output); 
                        if( !xb ) logMessage( "(" + xb + ") cs exec: " + output );
                    }

                    string composite = "CoffeeScript.eval(";
                    composite += jsescape(script);
                    composite += ");";

                    rslt = scripto.ExecString(composite, out output);
                    logMessage(output);
                    if (!rslt)
                    {
                        logMessage("(without source maps, coffeescript errors may be confusing)");
                    }

                }
                else
                {
                    rslt = scripto.ExecString(script, out output);
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

                if( rslt )
                {
                    string glob = scripto.GetGlobal();
                    Console.WriteLine(glob);
                }
            }
            catch (Exception x)
            {
                logMessage("Exception: " + x.Message);
            }

        }

         void SaveScript ( Excel.Workbook wb = null )
        {
            if (null == wb) wb = Globals.ThisAddIn.Application.ActiveWorkbook ;
            if (null == wb) return;

            string version = "1.1";
            string lang = cbScriptLanguage.SelectedItem.ToString();

            string source = "<source language=\"" + lang + "\"><![CDATA[" + editor.Text + "]]></source>";
            string xmlString1 = "<script version=\"" + version + "\" xmlns=\"" + SCRIPT_XML_NAMESPACE + "\">" + source + "</script>";

            Office.CustomXMLParts parts = wb.CustomXMLParts.SelectByNamespace(SCRIPT_XML_NAMESPACE);
            if (parts.Count > 0)
            {
                Office.CustomXMLPart part = parts[1];
                part.NamespaceManager.AddNamespace("N", SCRIPT_XML_NAMESPACE);

                Office.CustomXMLNode node = part.SelectSingleNode("//N:script/@version");
                if (null != node) node.Delete();
                node = part.SelectSingleNode("//N:script");
                node.AppendChildNode("version", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, version);

                node = part.SelectSingleNode("//N:source/@language");
                if (null != node) node.Delete();
                node = part.SelectSingleNode("//N:source");
                node.FirstChild.Delete();
                node.AppendChildNode("", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeCData, editor.Text);
                node.AppendChildNode("language", "", Office.MsoCustomXMLNodeType.msoCustomXMLNodeAttribute, lang);
            }
            else
            {
                wb.CustomXMLParts.Add(xmlString1);
            }

        }

        void LoadScript()
        {
            string text = "";
            string lang = "Javascript";

            if (null != Globals.ThisAddIn.Application.ActiveWorkbook)
            {
                Office.CustomXMLParts parts = Globals.ThisAddIn.Application.ActiveWorkbook.CustomXMLParts.SelectByNamespace(SCRIPT_XML_NAMESPACE);
                if (parts.Count > 0)
                {
                    Office.CustomXMLPart part = parts[1];
                    part.NamespaceManager.AddNamespace("N", SCRIPT_XML_NAMESPACE);

                    Office.CustomXMLNode node = part.SelectSingleNode("//N:source"); // .FirstChild;
                    if (null != node) text = node.FirstChild.NodeValue.ToString();

                    node = part.SelectSingleNode("//N:source/@language");
                    if (null != node) 
                    {
                        lang = node.NodeValue.ToString();
                    }
                }
            }

            cbScriptLanguage.SelectedIndex = 0;
            for (int i = 0; i < cbScriptLanguage.Items.Count; i++)
            {
                if (cbScriptLanguage.Items[i].ToString().Equals(lang))
                {
                    cbScriptLanguage.SelectedIndex = i;
                    break;
                }
            }
            editor.Text = text;
        }

        private void bClearLog_Click(object sender, EventArgs e)
        {
            logger.Text = "";
        }
    }

    class InstanceData
    {
        public Scripto scripto = null;
        public string buffer = "";
        public string log = "";
    }


}
