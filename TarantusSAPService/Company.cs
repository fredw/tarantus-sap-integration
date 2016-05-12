using System;
using System.Configuration;

namespace TarantusSAPService
{
    public static class Company
    {
        /// <summary>
        /// Returns company
        /// </summary>
        /// <returns>SAPbobsCOM.Company</returns>
        public static SAPbobsCOM.Company GetCompany()
        {
            // Application settings
            String SAPCompany = ConfigurationManager.AppSettings["SAP.Company"].ToString();
            String SAPUser = ConfigurationManager.AppSettings["SAP.User"].ToString();
            String SAPPassword = ConfigurationManager.AppSettings["SAP.Password"].ToString();
            String DatabaseType = ConfigurationManager.AppSettings["Database.Type"].ToString();
            String DatabaseHost = ConfigurationManager.AppSettings["Database.Host"].ToString();
            String DatabaseUser = ConfigurationManager.AppSettings["Database.User"].ToString();
            String DatabasePassword = ConfigurationManager.AppSettings["Database.Password"].ToString();

            SAPbobsCOM.Company Company = null;

            try
            {
                Company = new SAPbobsCOM.Company();
                Company.UseTrusted = true; // true = login with SAP user/password, false = login with database user/password
                Company.language = SAPbobsCOM.BoSuppLangs.ln_English;
                Company.Server = DatabaseHost;
                //Company.LicenseServer = DatabaseHost;
                Company.CompanyDB = SAPCompany;
                Company.DbUserName = DatabaseUser;            
                Company.DbPassword = DatabasePassword;
                
                switch (DatabaseType)
                {
                    case "MSSQL":
                        Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL;
                        break;
                    case "MSSQL2005":
                        Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
                        break;
                    case "MSSQL2008":
                        Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                        break;
                    case "MSSQL2012":
                        Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                        break;
                    default:
                        throw new Exception.CriticalException("Database not supported: " + DatabaseType);
                }

                // SAP (only needed if UseTrust = true)
                Company.UserName = SAPUser;
                Company.Password = SAPPassword;

                // If error occurs when try to connect
                if (Company.Connect() != 0)
                {
                    throw new Exception.CriticalException("SAP Company connection error: " + Company.GetLastErrorDescription());
                }

                return Company;
            }
            catch (System.Exception ex)
            {
                throw new Exception.CriticalException("SAP Company error: " + ex.ToString());
            }
        }
    }
}
