using MailMind.OutlookAddIn.Dialogs;
using MailMind.OutlookAddIn.Resources;
using System;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace MailMind.OutlookAddIn.UI
{
    [ComVisible(true)]
    public class RibbonButton : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI RibbonUI { get; set; }

        private const string ComposeWindowRibbonId = "Microsoft.Outlook.Mail.Compose";

        public RibbonButton()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {

            if (ribbonID == ComposeWindowRibbonId)
            {
                return GetResourceText("MailMind.OutlookAddIn.UI.RibbonButton.xml");
            }
            return null;
        }


        public stdole.IPictureDisp GetGenerateButtonIcon(Office.IRibbonControl _)
        {
            var buttonIcon = (Image)ImageResource.generate_email_button_icon;
            return Util.IconConverter.GetIPictureDispFromImage(buttonIcon);
        }

        #endregion

        #region Ribbon Callbacks

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            RibbonUI = ribbonUI;

        }

        public void OpenMailGenerationDialog(Office.IRibbonControl control)
        {
            var emailGenerationDialog = new EmailGenerationDialog();
            emailGenerationDialog.ShowDialog();
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
