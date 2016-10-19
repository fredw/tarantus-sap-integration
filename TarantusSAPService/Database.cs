using System;
using System.Configuration;
using System.Data.SqlClient;

namespace TarantusSAPService
{
    public static class Database
    {
        /// <summary>
        /// Returns database connection
        /// </summary>
        /// <returns>SqlConnection</returns>
        public static SqlConnection GetConnection()
        {
            String host = ConfigurationManager.AppSettings["Database.Host"].ToString();
            String user = ConfigurationManager.AppSettings["Database.User"].ToString();
            String password = ConfigurationManager.AppSettings["Database.Password"].ToString();
            String database = ConfigurationManager.AppSettings["Database.Database"].ToString();

            SqlConnection connection = null;

            try
            {
                connection = new SqlConnection(
                    "Data Source=" + host + ";" +
                    "Initial Catalog=" + database + ";" +
                    "User ID=" + user + ";" +
                    "Password =" + password + ";" +
                    "MultipleActiveResultSets=True;"
                );
                connection.Open();
                return connection;
            }
            catch (System.Exception ex)
            {
                throw new Exception.CriticalException("Database connection error: " + ex.ToString());
            }
        }

        /// <summary>
        /// Returns database command
        /// </summary>
        /// <returns>SqlCommand</returns>
        public static SqlCommand ExecuteCommand(
            String sql, 
            SqlConnection connection,
            SqlTransaction transaction = null
        ) {
            try
            {
                SqlCommand command = new SqlCommand(sql, connection);
                if (transaction != null)
                {
                    command.Transaction = transaction;
                }
                return command;
            }
            catch (System.Exception ex)
            {
                throw new System.Exception("Database command error: " + ex.ToString());
            }
        }
    }
}
