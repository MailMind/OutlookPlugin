using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace MailMind.OutlookAddIn
{
    [ComVisible(true)]
    public class ContextMenuRibbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI Ribbon { get; set; }

        public ContextMenuRibbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("OutlookAddIn1.ContextMenuRibbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        public void GetButtonID(Office.IRibbonControl control)
        {
            var sel = Globals.ThisAddIn.selection();
            sel.InsertAfter("{ResponseFromChatGPT}");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            Ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
