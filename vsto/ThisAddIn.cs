using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookArchiveOnReply
{
    public partial class ThisAddIn
    {
        private const string ARCHIVE_FOLDER_NAME = "Archiv";
        private const string ARCHIVE_FOLDER_NAME_EN = "Archive";
        private const int MAX_SCAN_ITEMS = 300;

        private readonly Dictionary<MailItem, MailItem> _pendingReplies =
            new Dictionary<MailItem, MailItem>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.ItemSend += Application_ItemSend;
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
        }

        internal void StartArchiveReply(MailItem original)
        {
            MailItem reply = original.Reply();
            _pendingReplies[reply] = original;

            ((ItemEvents_10_Event)reply).Close += delegate (ref bool cancel)
            {
                _pendingReplies.Remove(reply);
            };

            reply.Display();
        }

        internal bool IsArchivePending(MailItem reply)
        {
            return _pendingReplies.ContainsKey(reply);
        }

        internal void SetArchivePending(MailItem reply, bool pressed)
        {
            if (pressed)
            {
                if (_pendingReplies.ContainsKey(reply))
                {
                    return;
                }

                MailItem original = FindOriginalForReply(reply);
                if (original == null)
                {
                    System.Windows.Forms.MessageBox.Show(
                        "Die Original-E-Mail konnte nicht im Posteingang gefunden werden.",
                        "Antworten + Archivieren",
                        System.Windows.Forms.MessageBoxButtons.OK,
                        System.Windows.Forms.MessageBoxIcon.Information);
                    return;
                }

                _pendingReplies[reply] = original;

                ((ItemEvents_10_Event)reply).Close += delegate (ref bool cancel)
                {
                    _pendingReplies.Remove(reply);
                };
            }
            else
            {
                _pendingReplies.Remove(reply);
            }
        }

        private MailItem FindOriginalForReply(MailItem reply)
        {
            string replyIndex = reply.ConversationIndex;
            if (string.IsNullOrEmpty(replyIndex) || replyIndex.Length < 32)
            {
                return null;
            }

            string originalIndex = replyIndex.Substring(0, replyIndex.Length - 10);

            MAPIFolder inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);

            int scanned = 0;
            foreach (object obj in inbox.Items)
            {
                if (++scanned > MAX_SCAN_ITEMS)
                {
                    break;
                }

                if (obj is MailItem candidate)
                {
                    string candidateIndex = candidate.ConversationIndex;
                    if (string.Equals(candidateIndex, originalIndex, StringComparison.OrdinalIgnoreCase))
                    {
                        return candidate;
                    }
                }
            }

            return null;
        }

        private void Application_ItemSend(object item, ref bool cancel)
        {
            if (!(item is MailItem sentReply))
            {
                return;
            }

            MailItem original = null;
            foreach (KeyValuePair<MailItem, MailItem> entry in _pendingReplies)
            {
                if (ReferenceEquals(entry.Key, sentReply))
                {
                    original = entry.Value;
                    break;
                }
            }

            if (original == null)
            {
                return;
            }

            _pendingReplies.Remove(sentReply);

            try
            {
                MoveToArchive(original);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("Archivierung fehlgeschlagen: " + ex.Message);
            }
        }

        private void MoveToArchive(MailItem item)
        {
            MAPIFolder archiveFolder = FindArchiveFolder();
            if (archiveFolder != null)
            {
                item.Move(archiveFolder);
            }
        }

        private MAPIFolder FindArchiveFolder()
        {
            MAPIFolder inbox = Application.Session.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
            MAPIFolder root = inbox.Parent;

            foreach (MAPIFolder folder in root.Folders)
            {
                if (IsArchiveFolder(folder))
                {
                    return folder;
                }
            }

            foreach (MAPIFolder folder in root.Folders)
            {
                MAPIFolder match = FindArchiveFolderRecursive(folder);
                if (match != null)
                {
                    return match;
                }
            }

            return null;
        }

        private MAPIFolder FindArchiveFolderRecursive(MAPIFolder folder)
        {
            foreach (MAPIFolder sub in folder.Folders)
            {
                if (IsArchiveFolder(sub))
                {
                    return sub;
                }

                MAPIFolder match = FindArchiveFolderRecursive(sub);
                if (match != null)
                {
                    return match;
                }
            }
            return null;
        }

        private bool IsArchiveFolder(MAPIFolder folder)
        {
            string name = folder.Name;
            return string.Equals(name, ARCHIVE_FOLDER_NAME, StringComparison.OrdinalIgnoreCase)
                || string.Equals(name, ARCHIVE_FOLDER_NAME_EN, StringComparison.OrdinalIgnoreCase);
        }

        #region VSTO generated code

        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
