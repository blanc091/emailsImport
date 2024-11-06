using System;
using System.Data.SqlClient;
using System.Data;
using System.IO;

namespace emailsImport
{    
    public class emImport
    {
        public static class GlobalConfig
        {
            public static string ClientID { get; } = "00001";
            public static string connectionString { get; } = "*******";
            public static string connectionStringLogging { get; } = "********";
        }
        public static void Main(string[] args)
        {
            // hardcoded these for test
            string imapServer = "********";  
            int imapPort = 993;                           
            string username = "*********";     
            string password = "********";       
            string attachmentFolder = Path.Combine(Path.GetTempPath(), "Attachments");
            if (!Directory.Exists(Path.Combine(Path.GetTempPath(), "Attachments")))
                {
                Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "Attachments"));
                }
           // DataTable clientData = GetClientConnectionData(GlobalConfig.ClientID);
                /* if (clientData.Rows.Count > 0)
                 {
                     DataRow row = clientData.Rows[0];

                     string imapServer = row["imapServer"].ToString();
                     int imapPort = Convert.ToInt32(row["imapPort"]);
                     string username = row["username"].ToString();
                     string password = row["password"].ToString();
                     string attachmentFolder = row["attachmentFolder"].ToString();
                 }*/
                Logger.Log(GlobalConfig.ClientID, "Starting process..");
                FetchEmail.FetchAndSaveEmails(imapServer, imapPort, username, password, attachmentFolder);
                Logger.Log(GlobalConfig.ClientID, "Ending process..");
        }
        private static DataTable GetClientConnectionData(string client)
        {
            DataTable clientData = new DataTable();
            string query = "SELECT imapServer, imapPort, username, password, attachmentFolder FROM [ClientConnections] WHERE ClientID = @ClientID";
            using (SqlConnection connection = new SqlConnection(GlobalConfig.connectionString))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ClientID", client);
                    using (SqlDataAdapter adapter = new SqlDataAdapter(command))
                    {
                        try
                        {
                            connection.Open();
                            DataTable resultTable = new DataTable();
                            adapter.Fill(resultTable);
                            foreach (DataRow row in resultTable.Rows)
                            {
                                row["imapServer"] = EncryptionUtil.Decrypt(row["imapServer"].ToString());
                                row["username"] = EncryptionUtil.Decrypt(row["username"].ToString());
                                row["password"] = EncryptionUtil.Decrypt(row["password"].ToString());
                                row["attachmentFolder"] = EncryptionUtil.Decrypt(row["attachmentFolder"].ToString());
                            }
                            clientData = resultTable;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error retrieving client connection data: {ex.Message}");
                        }
                    }
                }
            }
            return clientData;
        }
    }
}
