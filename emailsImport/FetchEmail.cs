using MailKit;
using MailKit.Net.Imap;
using MailKit.Search;
using MimeKit;
using System;
using System.IO;
using System.Net;
using static emailsImport.emImport;

namespace emailsImport
{
    public static class FetchEmail
    {
        public static void FetchAndSaveEmails(string imapServer, int imapPort, string username, string password, string attachmentFolder)
        {
            try
            {
                using (var client = new ImapClient())
                {
                    ServicePointManager.CheckCertificateRevocationList = false;
                    ServicePointManager.ServerCertificateValidationCallback = (sender, certificate, chain, sslPolicyErrors) => true;
                    Logger.Log(GlobalConfig.ClientID, "Connecting to IMAP server..");
                    client.Connect(imapServer, imapPort, true);
                    Logger.Log(GlobalConfig.ClientID, "Authenticating to IMAP server..");
                    client.Authenticate(username, password);
                    Logger.Log(GlobalConfig.ClientID, "Opening Inbox folder..");
                    client.Inbox.Open(FolderAccess.ReadWrite);
                    var query = SearchQuery.Not(SearchQuery.Seen);
                    var uids = client.Inbox.Search(query);
                    foreach (var uid in uids)
                    {
                        
                        Logger.Log(GlobalConfig.ClientID, "Getting GUID..");
                        var message = client.Inbox.GetMessage(uid);
                        var guid = Guid.NewGuid().ToString();
                        Logger.Log(GlobalConfig.ClientID, $"Generated GUID: {guid}");
                        Logger.Log(GlobalConfig.ClientID, "Saving email metadata into the database..");
                        DBUtil.SaveEmailDataToDatabase(message, attachmentFolder, guid);

                        //client.Inbox.SetFlags(uid, MessageFlags.Seen | MessageFlags.Deleted, true);
                    }
                    // client.Inbox.Expunge();
                    client.Disconnect(true);
                    Logger.Log(GlobalConfig.ClientID, "Emails fetched, saved, and marked as read and deleted successfully.");
                }
            }
            catch (Exception ex)
            {
                Logger.Log(GlobalConfig.ClientID, $"Error occurred while fetching, saving, and marking emails as read and deleted: {ex.Message}");
            }
        }
    }
}
