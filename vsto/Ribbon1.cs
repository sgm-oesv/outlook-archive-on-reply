using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookArchiveOnReply
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private const string EXPLORER_XML =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
            "<contextMenus>" +
            "<contextMenu idMso=\"contextMenuMailItem\">" +
            "<button id=\"archiveAndReplyContextMenu\" label=\"Antworten + Archivieren\" " +
            "imageMso=\"MailReplyAll\" onAction=\"OnArchiveReply\" />" +
            "</contextMenu>" +
            "</contextMenus>" +
            "</customUI>";

        private const string READ_XML =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
            "<ribbon><tabs>" +
            "<tab idMso=\"TabReadMessage\">" +
            "<group idMso=\"GroupRespond\">" +
            "<button id=\"archiveAndReplyButton\" label=\"Antworten + Archivieren\" " +
            "size=\"large\" imageMso=\"MailReplyAll\" onAction=\"OnArchiveReply\" />" +
            "</group>" +
            "</tab>" +
            "</tabs></ribbon>" +
            "</customUI>";

        private const string COMPOSE_XML =
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
            "<customUI xmlns=\"http://schemas.microsoft.com/office/2009/07/customui\">" +
            "<ribbon><tabs>" +
            "<tab idMso=\"TabMessage\">" +
            "<group idMso=\"GroupSend\">" +
            "<checkBox id=\"archiveOnSendCheck\" label=\"Original nach Senden archivieren\" " +
            "getPressed=\"OnGetArchivePressed\" onAction=\"OnArchiveCheckToggle\" />" +
            "</group>" +
            "</tab>" +
            "</tabs></ribbon>" +
            "</customUI>";

        public Ribbon1()
        {
        }

        public string GetCustomUI(string ribbonID)
        {
            switch (ribbonID)
            {
                case "Microsoft.Outlook.Explorer":
                    return EXPLORER_XML;
                case "Microsoft.Outlook.Mail.Read":
                    return READ_XML;
                case "Microsoft.Outlook.Mail.Compose":
                    return COMPOSE_XML;
                default:
                    return null;
            }
        }

        public void OnArchiveReply(Office.IRibbonControl control)
        {
            if (control.Context is Inspector inspector)
            {
                if (inspector.CurrentItem is MailItem mail)
                {
                    Globals.ThisAddIn.StartArchiveReply(mail);
                }
            }
            else if (control.Context is Explorer explorer)
            {
                Selection selection = explorer.Selection;
                if (selection.Count > 0 && selection[1] is MailItem mail)
                {
                    Globals.ThisAddIn.StartArchiveReply(mail);
                }
            }
        }

        public bool OnGetArchivePressed(Office.IRibbonControl control)
        {
            if (control.Context is Inspector inspector
                && inspector.CurrentItem is MailItem reply)
            {
                return Globals.ThisAddIn.IsArchivePending(reply);
            }
            return false;
        }

        public void OnArchiveCheckToggle(Office.IRibbonControl control, bool pressed)
        {
            if (control.Context is Inspector inspector
                && inspector.CurrentItem is MailItem reply)
            {
                Globals.ThisAddIn.SetArchivePending(reply, pressed);
            }
        }
    }
}
