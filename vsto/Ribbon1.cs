using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookArchiveOnReply
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        public Ribbon1()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Mail.Read")
            {
                return GetResourceText("OutlookArchiveOnReply.Ribbon1.xml");
            }
            return null;
        }

        public void OnArchiveReply(Office.IRibbonControl control)
        {
            if (control.Context is Inspector inspector
                && inspector.CurrentItem is MailItem mail)
            {
                Globals.ThisAddIn.StartArchiveReply(mail);
            }
        }

        #region Hilfsmethoden

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith(resourceName, StringComparison.OrdinalIgnoreCase))
                {
                    using (Stream stream = asm.GetManifestResourceStream(name))
                    using (var reader = new StreamReader(stream))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
