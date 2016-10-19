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

            SAPbobsCOM.Company company = null;

            try
            {
                company = new SAPbobsCOM.Company();
                company.UseTrusted = true; // true = login with SAP user/password, false = login with database user/password
                company.language = SAPbobsCOM.BoSuppLangs.ln_English;
                company.Server = DatabaseHost;
                //Company.LicenseServer = DatabaseHost;
                company.CompanyDB = SAPCompany;
                company.DbUserName = DatabaseUser;            
                company.DbPassword = DatabasePassword;
                
                switch (DatabaseType)
                {
                    case "MSSQL":
                        company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL;
                        break;
                    case "MSSQL2005":
                        company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005;
                        break;
                    case "MSSQL2008":
                        company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008;
                        break;
                    case "MSSQL2012":
                        company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012;
                        break;
                    default:
                        throw new Exception.CriticalException("Database not supported: " + DatabaseType);
                }

                // SAP (only needed if UseTrust = true)
                company.UserName = SAPUser;
                company.Password = SAPPassword;

                // If error occurs when try to connect
                if (company.Connect() != 0)
                {
                    throw new Exception.CriticalException("SAP Company connection error: " + company.GetLastErrorDescription());
                }

                return company;
            }
            catch (System.Exception ex)
            {
                throw new Exception.CriticalException("SAP Company error: " + ex.ToString());
            }
        }
    }
}
