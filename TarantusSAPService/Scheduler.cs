using System;
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
            // Create a timer to excute every 30 seconds
            timer = new System.Timers.Timer();
            timer.Interval = 30000;
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
            SAPbobsCOM.Company oCompany = null;
            SqlConnection Connection = null;

            try
            {               
                // Get SAP company and database connection
                oCompany = Company.GetCompany();
                Connection = Database.GetConnection();
                
                if (oCompany.Connected)
                {
                    LogWriter.write("SAP Company connection successfuly!");
                    // Execute tasks
                    Task.Order.Execute(oCompany, Connection);
                }

            } catch (Exception.CriticalException ex)
            {
                // When critical error occur, write log with error and stop service
                LogWriter.write(ex.ToString());
                (new ServiceController("TarantusSAPIntegrationService")).Stop();
                Environment.Exit(0);
                LogWriter.write("Service stopped");
            } catch (System.Exception ex)
            {
                // When generic error occur, only write a log error
                LogWriter.write("Error: " + ex.ToString());
            } finally
            {
                // Disconnect database
                if (Connection != null)
                {
                    Connection.Close();
                }
                // Disconnect companmy
                if (oCompany != null)
                {
                    oCompany.Disconnect();
                }
            }            
        }
    }
}
