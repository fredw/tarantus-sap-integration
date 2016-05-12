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
            String DatabaseHost = ConfigurationManager.AppSettings["Database.Host"].ToString();
            String DatabaseUser = ConfigurationManager.AppSettings["Database.User"].ToString();
            String DatabasePassword = ConfigurationManager.AppSettings["Database.Password"].ToString();
            String DatabaseDatabase = ConfigurationManager.AppSettings["Database.Database"].ToString();

            SqlConnection Connection = null;

            try
            {
                Connection = new SqlConnection(
                    "Data Source=" + DatabaseHost + ";" +
                    "Initial Catalog=" + DatabaseDatabase + ";" +
                    "User ID=" + DatabaseUser + ";" +
                    "Password =" + DatabasePassword + ";" +
                    "MultipleActiveResultSets=True;"
                );
                Connection.Open();
                return Connection;
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
