using System;
using System.Data.SqlClient;
using System.IO;
using static emailsImport.emImport;

namespace emailsImport
{
    public static class Logger
    {
        private static readonly string logDirectory = Path.Combine(Path.GetTempPath(), "Logging");
        private static readonly string logFilePath = Path.Combine(logDirectory, DateTime.Now.ToString("yyyyMMdd") + "_log.txt");

        public static void Log(string project, string message)
        {
            try
            {
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }

                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    writer.WriteLine($"{DateTime.Now}: {message}");
                }

                LogToDatabase(project, message);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error occurred while logging: {ex.Message}");
            }
        }

        private static void LogToDatabase(string project, string message)
        {
            string query = "INSERT INTO serverSideLogging (ClientID, LogMessage, Timestamp) VALUES (@ClientID, @LogMessage, current_timestamp)";

            using (SqlConnection connection = new SqlConnection(GlobalConfig.connectionStringLogging))
            {
                using (SqlCommand command = new SqlCommand(query, connection))
                {
                    command.Parameters.AddWithValue("@ClientID", GlobalConfig.ClientID);
                    command.Parameters.AddWithValue("@LogMessage", message);

                    connection.Open();
                    command.ExecuteNonQuery();
                }
            }
        }
    }
}
