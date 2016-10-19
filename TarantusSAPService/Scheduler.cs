using System;
using System.Configuration;
using System.ServiceProcess;
using System.Timers;
using System.Data.SqlClient;

namespace TarantusSAPService
{
    public partial class Scheduler : ServiceBase
    {
        private System.Timers.Timer timer = null;

        public Scheduler()
        {
            InitializeComponent();
        }

        /// <summary>
        /// On service start
        /// </summary>
        /// <param name="args"></param>
        protected override void OnStart(string[] args)
        {
            // Default interval: 30 seconds
            int interval = 30;

            if (!int.TryParse(ConfigurationManager.AppSettings["Scheduler.Interval"].ToString(), out interval))
            {
                LogWriter.write("Error: invalid Scheduler.Interval. It was assumed the default value of 30 seconds");
            }

            // Create a timer to excute periodically
            timer = new System.Timers.Timer();
            timer.Interval = interval * 1000;
            timer.Elapsed += new ElapsedEventHandler(this.TimerTick);
            timer.Enabled = true;
            LogWriter.write("Service started");
        }
        
        /// <summary>
        /// On service stop
        /// </summary>
        protected override void OnStop()
        {
            // Stop timer and store log
            timer.Enabled = false;
            LogWriter.write("Service stopped");
        }

        /// <summary>
        /// Recursive function executed every time elapsed
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void TimerTick(object sender, ElapsedEventArgs e)
        {
            SAPbobsCOM.Company company = null;
            SqlConnection connection = null;

            try
            {               
                // Get SAP company and database connection
                company = Company.GetCompany();
                connection = Database.GetConnection();
                
                if (company.Connected)
                {
                    LogWriter.write("SAP Company connection successfuly!");
                    // Execute tasks
                    Task.Order.Execute(company, connection);
                }

            // When critical error occur, write log with error and stop service
            } catch (Exception.CriticalException ex)
            {                
                LogWriter.write("Critical error: " + ex.ToString());
                (new ServiceController("TarantusSAPIntegrationService")).Stop();
                Environment.Exit(0);
                LogWriter.write("Service stopped");
            // When generic error occur, only write a log error
            } catch (System.Exception ex)
            {                
                LogWriter.write("Error: " + ex.ToString());
            } finally
            {
                // Disconnect database
                if (connection != null)
                {
                    connection.Close();
                }
                // Disconnect companmy
                if (company != null)
                {
                    company.Disconnect();
                }
            }            
        }
    }
}
