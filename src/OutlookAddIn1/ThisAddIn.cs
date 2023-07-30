using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.Net;
using Microsoft.Office.Interop.Word;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        private Word.Selection _selection;
        public Word.Selection SelectionBeforeClick => _selection;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // This is needed to setup the correct TLS
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls
                | SecurityProtocolType.Tls11
                | SecurityProtocolType.Tls12
                | SecurityProtocolType.Ssl3;

            Outlook.Application application = this.Application;
            // Add a new Inspector
            application.Inspectors.NewInspector +=
                new Outlook.InspectorsEvents_NewInspectorEventHandler(
                    Inspectors_AddTextAtPosition);
        }

        public Word.Selection selection() { return _selection; }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new ContextMenuRibbon();
        }

        private void Inspectors_AddTextAtPosition(Outlook.Inspector inspector)
        {
            Outlook.MailItem mailItem = inspector.CurrentItem as Outlook.MailItem;
            if (mailItem != null)
            {
                // FIXME: this seems to work only sometimes..
                if (inspector.EditorType == OlEditorType.olEditorWord && inspector.IsWordMail())
                {
                    // Get the Word document
                    Word.Document document = inspector.WordEditor;
                    if (document != null)
                    {
                        // Subscribe to the BeforeDoubleClick event of the Word document
                        document.Application.WindowBeforeRightClick +=
                            new Word.ApplicationEvents4_WindowBeforeRightClickEventHandler(
                                ApplicationOnWindowBeforeRightClick);
                    }
                }
            }
        }
        private void ApplicationOnWindowBeforeRightClick(Word.Selection selection, ref bool cancel)
        {
            // Get the selected word
            _selection = selection;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
