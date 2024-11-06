using System;
using System.Data.SqlClient;
using System.Security.Policy;
using System.Threading.Tasks;
using static emailsImport.emImport;

namespace emailsImport
{
    internal class generateURN
    {
        public static string GenerateURN()
        {
            int sequencer = GetAndSetSeq();
            string URN = GlobalConfig.ClientID + sequencer.ToString().PadLeft(7, '0');
            return URN;
        }

        public static int GetAndSetSeq(int maxRetries = 3, int delayBetweenRetries = 1000)
        {
            int sequence = 0;
            bool success = false;
            int retryCount = 0;

            while (!success && retryCount < maxRetries)
            {
                try
                {
                    using (SqlConnection myConnection = new SqlConnection(GlobalConfig.connectionString))
                    {
                        myConnection.Open();
                        using (SqlTransaction transaction = myConnection.BeginTransaction())
                        {
                            try
                            {
                                string getCommand = "SELECT CurrentSequence FROM [ClientData].[dbo].[URNSequencer] WITH (NOLOCK) WHERE ClientID = @ClientID";
                                using (SqlCommand getCmd = new SqlCommand(getCommand, myConnection, transaction))
                                {
                                    getCmd.Parameters.AddWithValue("@ClientID", GlobalConfig.ClientID);
                                    using (SqlDataReader reader = getCmd.ExecuteReader())
                                    {
                                        if (reader.Read())
                                        {
                                            sequence = Convert.ToInt32(reader["CurrentSequence"]);
                                        }
                                        else
                                        {
                                            throw new Exception("ClientID not found in URNSequencer table.");
                                        }
                                        reader.Close();
                                    }
                                }

                                if (sequence == 9999999)
                                {
                                    string resetCommand = "UPDATE [ClientData].[dbo].[URNSequencer] SET CurrentSequence = 1 WHERE ClientID = @ClientID";
                                    using (SqlCommand resetCmd = new SqlCommand(resetCommand, myConnection, transaction))
                                    {
                                        resetCmd.Parameters.AddWithValue("@ClientID", GlobalConfig.ClientID);
                                        resetCmd.ExecuteNonQuery();
                                        sequence = 1;
                                    }
                                }
                                else
                                {
                                    string updateCommand = "UPDATE [ClientData].[dbo].[URNSequencer] SET CurrentSequence = CurrentSequence + 1 WHERE ClientID = @ClientID";
                                    using (SqlCommand updateCmd = new SqlCommand(updateCommand, myConnection, transaction))
                                    {
                                        updateCmd.Parameters.AddWithValue("@ClientID", GlobalConfig.ClientID);
                                        updateCmd.ExecuteNonQuery();
                                        sequence += 1;
                                    }
                                }

                                transaction.Commit();
                                success = true; // Mark as successful, no retries needed
                            }
                            catch (Exception ex)
                            {
                                Logger.Log(GlobalConfig.ClientID, "Error processing URN sequencer: " + ex.Message);
                                Console.WriteLine("Error processing URN sequencer: " + ex.Message);
                                transaction.Rollback();
                                retryCount++;
                                if (retryCount < maxRetries)
                                {
                                    Task.Delay(delayBetweenRetries * retryCount).Wait();
                                }
                                else
                                {
                                    throw;
                                }
                            }
                        }
                    }
                }
                catch (SqlException sqlEx)
                {
                    Logger.Log(GlobalConfig.ClientID, "SQL error: " + sqlEx.Message);
                    Console.WriteLine("SQL error: " + sqlEx.Message);
                    retryCount++;
                    if (retryCount < maxRetries)
                    {
                        Task.Delay(delayBetweenRetries * retryCount).Wait();
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (Exception ex)
                {
                    Logger.Log(GlobalConfig.ClientID, "General error: " + ex.Message);
                    Console.WriteLine("General error: " + ex.Message);
                    retryCount++;
                    if (retryCount >= maxRetries)
                    {
                        throw;
                    }
                    Task.Delay(delayBetweenRetries * retryCount).Wait();
                }
            }

            return sequence;
        }
    }

}
