using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace webapp45
{
    class DBManager
    {
        private string connection = ConfigurationManager.ConnectionStrings["GroupDBConnection"].ConnectionString;
       
        public void AddGroup(string groupName, string webhookUrl)
        {
            // make sure the group name doesn't exist first
            if (string.IsNullOrEmpty(GetWebhookUrl(groupName)))
            {
                string insertGroupCommand = "INSERT INTO [group] (groupname,webhookurl,createddate) VALUES (@val1, @val2, @val3)";
                using (SqlConnection conn = new SqlConnection(connection))
                {
                    using (SqlCommand cmd = new SqlCommand())
                    {
                        cmd.Connection = conn;
                        cmd.CommandText = insertGroupCommand;
                        cmd.Parameters.AddWithValue("@val1", groupName);
                        cmd.Parameters.AddWithValue("@val2", webhookUrl);
                        cmd.Parameters.AddWithValue("@val3", DateTime.UtcNow);

                        try
                        {
                            if (conn.State != System.Data.ConnectionState.Open)
                                conn.Open();

                            cmd.ExecuteNonQuery();
                        }
                        catch (SqlException sqlex)
                        {
                            throw sqlex;
                        }
                        finally
                        {
                            conn.Close();
                        }
                    }
                }
            }
        }
        
        public string GetWebhookUrl(string groupName)
        {
            string selectCommand = "SELECT webhookurl from [group] where groupname=@val1";
            using (SqlConnection conn = new SqlConnection(connection))
            {
                using (SqlCommand cmd = new SqlCommand())
                {
                    cmd.Connection = conn;
                    cmd.CommandText = selectCommand;
                    cmd.Parameters.AddWithValue("@val1", groupName);

                    try
                    {
                        if (conn.State != System.Data.ConnectionState.Open)
                            conn.Open();

                        SqlDataReader reader= cmd.ExecuteReader();

                        if (reader.HasRows)
                        {
                            while (reader.Read())
                            {
                                return reader.GetString(0);
                            }
                        }
                        else
                            return string.Empty;
                    }
                    catch (SqlException sqlex)
                    {
                        throw sqlex;
                        
                    }
                }
            }
            return string.Empty;
        }

    }
}
