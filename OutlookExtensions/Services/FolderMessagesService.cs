using Microsoft.Office.Interop.Outlook;
using NLog;

namespace OutlookExtensions.Services
{
    public class FolderMessagesService
    {
        private ILogger _logger;

        public FolderMessagesService()
        {
            _logger = LogManager.GetLogger(nameof(FolderMessagesService));
        }

        internal void ProcessFolder(Folder folder, bool unread, bool processChildFolders)
        {
            _logger.Info($"Processing {folder.Name} catalog");
            int itemCount = unread ? folder.UnReadItemCount : folder.Items.Count - folder.UnReadItemCount;
            while (itemCount > 0)
            {
                _logger.Info($"{folder.Name} contains {folder.UnReadItemCount} ${(unread ? "unread" : "read")} messages");
                Items unreadMessages = folder.Items.Restrict($"[Unread] = {unread.ToString()}");

                foreach (MailItem mail in unreadMessages)
                {
                    if (mail.UnRead == unread)
                    {
                        itemCount--;
                        mail.UnRead = !unread;
                        mail.Close(OlInspectorClose.olSave);
                    }
                }

            }

            if (processChildFolders)
            {
                ProcessChildFolders(folder.Folders, unread, processChildFolders);
            }
        }

        internal void ProcessChildFolders(Folders folders, bool unread, bool processChildFolders)
        {
            if (folders != null)
            {
                foreach (Folder folder in folders)
                {
                    ProcessFolder(folder, unread, processChildFolders);
                }
            }
        }
    }
}
