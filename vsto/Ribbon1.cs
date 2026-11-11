using System;
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
                return
                    "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                    "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
                    "<ribbon><tabs>" +
                    "<tab idMso=\"TabReadMessage\">" +
                    "<group id=\"ArchiveGroup\" label=\"Archivierung\">" +
                    "<button id=\"archiveAndReplyButton\" label=\"Antworten + Archivieren\" " +
                    "size=\"large\" imageMso=\"MailReplyAll\" onAction=\"OnArchiveReply\" visible=\"true\" />" +
                    "</group>" +
                    "</tab>" +
                    "</tabs></ribbon>" +
                    "</customUI>";
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
    }
}
