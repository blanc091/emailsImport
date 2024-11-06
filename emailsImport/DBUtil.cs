using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;
using MimeKit;
using Spire.Doc;
using Spire.Doc.Documents;
using System.Drawing;
using static System.Net.Mime.MediaTypeNames;
using System.Dynamic;
using System.Collections;
using static emailsImport.emImport;

namespace emailsImport
{
    public static class DBUtil
    {
        public static void FetchFileFromDatabase(string outputDirectory, string searchQuery)
        {
            string query = @"
        SELECT AttachmentName, Attachments, Sender
        FROM EmailData 
        WHERE Subject LIKE @SearchQuery 
           OR Sender LIKE @SearchQuery 
           OR AttachmentName LIKE @SearchQuery";

            using (SqlConnection connection = new SqlConnection(GlobalConfig.connectionString))
            {
                connection.Open();
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@SearchQuery", "%" + searchQuery + "%");
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        DataTable resultTable = new DataTable();
                        adapter.Fill(resultTable);
                        if (resultTable.Rows.Count > 0)
                        {
                            foreach (DataRow row in resultTable.Rows)
                            {
                                string encryptedSender = row["Sender"].ToString();
                                string decryptedSender = EncryptionUtil.Decrypt(encryptedSender);
                                Logger.Log(GlobalConfig.ClientID, "Decrypted Sender: " + decryptedSender);
                                string fileName = row["AttachmentName"].ToString();
                                byte[] attachmentData = (byte[])row["Attachments"];
                                string filePath = Path.Combine(outputDirectory, fileName);
                                File.WriteAllBytes(filePath, attachmentData);
                                Logger.Log(GlobalConfig.ClientID, "Attachment saved as: " + filePath);
                            }
                        }
                        else
                        {
                            Logger.Log(GlobalConfig.ClientID, "No results found for the given search query.");
                        }
                    }
                }
            }
        }
        public static void SaveEmailDataToDatabase(MimeMessage message, string attachmentFolder, string guid)
        {
            try
            {
                using (SqlConnection connection = new SqlConnection(GlobalConfig.connectionString))
                {
                    Logger.Log(GlobalConfig.ClientID, "Connecting to database..");
                    connection.Open();
                    string query = "INSERT INTO EmailData (Subject, Sender, Receiver, Date, AttachmentName, Attachments, URN, GUID) VALUES (@Subject, @Sender, @Receiver, current_timestamp, @AttachmentName, @Attachments, @URN, @GUID)";
                    Logger.Log(GlobalConfig.ClientID, "Query: " + query);
                    using (SqlCommand command = new SqlCommand(query, connection))
                    {
                        int startIndex = message.From.ToString().IndexOf('<');
                        int endIndex = message.From.ToString().IndexOf('>');
                        string emailAddress = message.From.ToString().Substring(startIndex + 1, endIndex - startIndex - 1);
                        Logger.Log(GlobalConfig.ClientID, "Email Address (String Manipulation): " + emailAddress);
                        string pattern = @"\<(.*?)\>";
                        Match match = Regex.Match(message.From.ToString(), pattern);
                        if (match.Success)
                        {
                            emailAddress = match.Groups[1].Value;
                            Logger.Log(GlobalConfig.ClientID, "Email Address (Regular Expressions): " + emailAddress);
                        }
                        command.Parameters.AddWithValue("@Subject", EncryptionUtil.Encrypt(message.Subject));
                        command.Parameters.AddWithValue("@Sender", EncryptionUtil.Encrypt(emailAddress));
                        command.Parameters.AddWithValue("@Receiver", EncryptionUtil.Encrypt(message.To.ToString()));
                        command.Parameters.Add("@AttachmentName", System.Data.SqlDbType.NVarChar, 1000);
                        command.Parameters.Add("@Attachments", SqlDbType.VarBinary);
                        command.Parameters.AddWithValue("@URN", generateURN.GenerateURN());
                        command.Parameters.AddWithValue("@GUID", guid);
                        var fileName = "";
                        var filePath = "";
                        var filePathTIFF = "";
                       // var filePathGUID = "";
                        if (message.Attachments.Any())
                        {
                            int errFlag = 0;
                            bool hasInvalidAttachments = false;
                            foreach (var attachment in message.Attachments.OfType<MimePart>())
                            {

                                fileName = attachment.FileName;
                                filePath = Path.Combine(attachmentFolder, guid, fileName);
                                filePathTIFF = Path.Combine(attachmentFolder, guid);
                                if (!Directory.Exists(filePathTIFF))
                                {
                                    Directory.CreateDirectory(filePathTIFF);
                                }
                                Logger.Log(GlobalConfig.ClientID, "Filename: " + fileName);
                                Logger.Log(GlobalConfig.ClientID, "File path: " + filePath);

                                var attachmentData = ConvertMimePartToSqlBinary(attachment);
                                if (!Directory.Exists(Path.Combine(attachmentFolder, guid)))
                                {
                                    Directory.CreateDirectory(Path.Combine(attachmentFolder, guid));
                                }
                                using (var stream = File.Create(filePath))
                                {
                                    Logger.Log(GlobalConfig.ClientID, "Saving file: " + filePath);
                                    attachment.Content.DecodeTo(stream);
                                    Logger.Log(GlobalConfig.ClientID, "Saved file: " + filePath);
                                }
                                string extension = Path.GetExtension(fileName).ToLower();
                                if (extension.ToLower() != ".pdf" && extension.ToLower() != ".tif" && extension.ToLower() != ".tiff")
                                {
                                    if (File.Exists(filePath))
                                        {
                                        File.Delete(filePath);
                                        }
                                    errFlag = 4;
                                    hasInvalidAttachments = true;
                                }
                                else
                                {                                    
                                    Logger.Log(GlobalConfig.ClientID, "Attachments are valid.");
                                    errFlag = 0;
                                    ////// call GDPicture PDF converter here if they are pdfs, if multipage tiffs, split
                                    //fileConversions converter = new fileConversions();
                                    //converter.createInternalTIFFS(filePath, filePathTIFF, attachmentFolder);
                                }

                                command.Parameters["@AttachmentName"].Value = EncryptionUtil.Encrypt(fileName);
                                command.Parameters["@Attachments"].Value = attachmentData;
                                Logger.Log(GlobalConfig.ClientID, "Adding attachment name: " + fileName);

                                command.ExecuteNonQuery();
                            }
                            if (hasInvalidAttachments)
                            {
                                errFlag = 4; // Both valid and invalid attachments are present
                            }
                            else
                            {
                                errFlag = 0; // Only valid attachments are present
                            }
                            Logger.Log(GlobalConfig.ClientID, "Saving email body...");
                            if (message.HtmlBody != null && (errFlag == 0 || errFlag == 4))
                            {
                                string imagePath = Path.Combine(attachmentFolder, guid, Path.GetFileNameWithoutExtension(fileName) + "_emailbody.tif");
                                Logger.Log(GlobalConfig.ClientID, "Filename for body: " + Path.GetFileNameWithoutExtension(fileName));
                                try
                                {
                                    using (Document document = new Document())
                                    {
                                        Section section = document.AddSection();
                                        Paragraph paragraph = section.AddParagraph();
                                        paragraph.AppendHTML(message.HtmlBody);

                                        System.Drawing.Image[] images = document.SaveToImages(Spire.Doc.Documents.ImageType.Metafile);
                                        if (images.Length > 0)
                                        {
                                            System.Drawing.Image image = images[0];
                                            image.Save(imagePath, System.Drawing.Imaging.ImageFormat.Tiff);
                                        }
                                        else
                                        {
                                            Logger.Log(GlobalConfig.ClientID, "Error: Unable to convert email body to image.");
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.Log(GlobalConfig.ClientID, $"Error occurred while saving email body: {ex.Message}");
                                }
                            }
                            else
                            {
                                Logger.Log(GlobalConfig.ClientID, "There is no email body");
                            }
                            autoReplies repliesHandler = new autoReplies();
                            repliesHandler.SendAlertEmail(emailAddress, message.Subject, errFlag); // call autoreply handling for attachments
                        }
                        else
                        {
                            autoReplies repliesHandler = new autoReplies();
                            repliesHandler.SendAlertEmail(emailAddress, message.Subject, 2); // call autoreply handling for no attachments
                        }
                    }
                    Logger.Log(GlobalConfig.ClientID, "Email data saved to database successfully.");
                }
            }
            catch (Exception ex)
            {
                Logger.Log(GlobalConfig.ClientID, $"Error occurred while handling emails: {ex.Message}");
                throw;
            }
        }
        private static SqlBinary ConvertMimePartToSqlBinary(MimePart attachment)
        {
            using (var memoryStream = new MemoryStream())
            {
                attachment.Content.DecodeTo(memoryStream);
                byte[] byteArray = memoryStream.ToArray();
                return new SqlBinary(byteArray);
            }
        }
    }
}
