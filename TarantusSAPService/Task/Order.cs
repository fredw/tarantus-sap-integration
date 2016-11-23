using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using TarantusSAPService.Exception;

namespace TarantusSAPService.Task
{
    public static class Order
    {
        public static void Execute(SAPbobsCOM.Company company, SqlConnection connection)
        {
            // Default records limit: 30
            int limit = 30;

            if (!int.TryParse(ConfigurationManager.AppSettings["Scheduler.Records.Process.Limit"].ToString(), out limit))
            {
                LogWriter.write("Error: invalid Scheduler.Records.Process.Limit. It was assumed the default value of 30 records");
            }

            // Select OPEN orders from temporary table
            SqlCommand commandOrder = Database.ExecuteCommand(
                "SELECT " +
                    "TOP " + limit + " " +
                    "DocEntry, " +
                    "CardCode, " +
                    "AddressCharge, " +
                    "AddressDelivery, " +
                    "DocDate, " +
                    "DeliveryDate, " +
                    "DiscountPrice, " +
                    "SlpCode, " +
                    "ClientOrderNum, " +
                    "Observation, " +
                    "PaymentCondition, " +
                    "PriceTable, " +
                    "Carrier, " +
                    "Shipping " +
                "FROM [@Order] " +
                "WHERE Status = @Status",
                connection
            );
            commandOrder.Parameters.AddWithValue("@Status", "O");
            SqlDataReader dataOrders = commandOrder.ExecuteReader();

            if (!dataOrders.HasRows)
            {
                LogWriter.write("There is no open orders to import.");
                return;
            }

            // Convert order result to a list to be able to loop more than once
            var dt = new DataTable();
            dt.Load(dataOrders);
            List<DataRow> orders = dt.AsEnumerable().ToList();

            // Join all DocEntrys on a string
            string DocEntry = string.Join(",", orders.Select(row => row["DocEntry"].ToString()).ToList());

            // Change all OPEN orders status to "W" (Waiting) to prevent these orders to be processed twice
            SqlCommand commandUpdateWaiting = Database.ExecuteCommand("UPDATE [@Order] SET Status = @Status WHERE DocEntry in (@DocEntry)", connection);
            commandUpdateWaiting.Parameters.AddWithValue("@Status", "W");
            commandUpdateWaiting.Parameters.AddWithValue("@DocEntry", DocEntry);
            commandUpdateWaiting.ExecuteNonQuery();
            LogWriter.write("Changed Status of temporary orders to *Waiting*: " + DocEntry);
            
            // Process orders
            foreach (DataRow order in orders)
            {
                try
                {
                    // Customer
                    SAPbobsCOM.BusinessPartners customerSAP = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                    // Try to find customer
                    if (!customerSAP.GetByKey((string) order["CardCode"]))
                    {
                        throw new SAPObjectException("Customer not found " + order["CardCode"] + " (OCRD.CardCode)", (int) order["DocEntry"]);
                    }

                    // Update customer price table and payment condition
                    customerSAP.PriceListNum = (int) order["PriceTable"];
                    customerSAP.PayTermsGrpCode = (int) order["PaymentCondition"];
                    customerSAP.Update();

                    if (customerSAP.Update() != 0)
                    {
                        throw new SAPObjectException("Error to update customer " + order["CardCode"] + " - " + company.GetLastErrorDescription(), (int) order["DocEntry"]);
                    }

                    // Get last DocNum from ORDR table
                    SqlCommand commandLastOrdr = Database.ExecuteCommand("SELECT TOP 1 DocNum FROM ORDR ORDER BY DocNum desc", connection);
                    int docNumNext = Int32.Parse(commandLastOrdr.ExecuteScalar().ToString()) + 1;

                    Double discountPrice = order["DiscountPrice"] == DBNull.Value ? 0 : Double.Parse(((decimal) order["DiscountPrice"]).ToString());

                    // Create the order object and set his properties
                    SAPbobsCOM.Documents orderSAP = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                    orderSAP.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                    orderSAP.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
                    orderSAP.BPL_IDAssignedToInvoice = 1; // First business place (for multibranch companies)
                    orderSAP.DocNum = docNumNext;
                    orderSAP.CardCode = (string) order["CardCode"];
                    orderSAP.PayToCode = (string) order["AddressCharge"];
                    orderSAP.ShipToCode = (string) order["AddressDelivery"];
                    orderSAP.DocDate = (DateTime) order["DocDate"];
                    orderSAP.TaxDate = (DateTime) order["DocDate"];
                    orderSAP.DocDueDate = (DateTime) order["DeliveryDate"];
                    orderSAP.NumAtCard = order["ClientOrderNum"] == DBNull.Value ? null : (string) order["ClientOrderNum"];
                    orderSAP.Comments = (string) order["Observation"];
                    orderSAP.SalesPersonCode = (int) order["SlpCode"];
                    orderSAP.GroupNumber = (int) order["PaymentCondition"];
                    orderSAP.TaxExtension.Carrier = order["Carrier"] == DBNull.Value ? null : (string) order["Carrier"];
                    orderSAP.TaxExtension.Incoterms = order["Shipping"].ToString();
                    // Custom fields
                    orderSAP.UserFields.Fields.Item("U_LstNum").Value = (int) order["PriceTable"];
                    orderSAP.UserFields.Fields.Item("U_Paciente").Value = ((int) order["PriceTable"]).ToString();
                    orderSAP.UserFields.Fields.Item("U_Desc_Global").Value = discountPrice;

                    // Select temporary itens from this temporary order
                    SqlCommand commandOrderItens = Database.ExecuteCommand(
                        "SELECT DocEntry, LineNum, ItemCode, Quantity, PriceUnitFinal " +
                        "FROM [@OrderItem] " +
                        "WHERE DocEntry = @DocEntry",
                        connection
                    );
                    commandOrderItens.Parameters.AddWithValue("@DocEntry", (int) order["DocEntry"]);
                    SqlDataReader orderItens = commandOrderItens.ExecuteReader();

                    int line = 0;
                    while (orderItens.Read())
                    {
                        if (line != 0)
                        {
                            // Create new item line
                            orderSAP.Lines.Add();
                        }

                        // Select item default wharehouse
                        SqlCommand commandItem = Database.ExecuteCommand("SELECT DfltWH FROM OITM WHERE ItemCode = @ItemCode", connection);
                        commandItem.Parameters.AddWithValue("@ItemCode", (string) orderItens["ItemCode"]);
                        SqlDataReader item = commandItem.ExecuteReader(System.Data.CommandBehavior.SingleRow);

                        // If not found item on database
                        if (!item.Read())
                        {
                            throw new SAPObjectException("Item not found " + orderItens["ItemCode"] + " (OITM.ItemCode)", (int) order["DocEntry"]);
                        }
                        // If item doesn't have a default wharehouse defined
                        else if (item.IsDBNull(item.GetOrdinal("DfltWH")))
                        {
                            throw new SAPObjectException("Item " + orderItens["ItemCode"] + " hasn't default wharehouse defined (OITM.DfltWH)", (int) order["DocEntry"]);
                        }

                        // Set item properties 
                        // Observation 1: Price doesn`t need to be setted because its loaded by customer price table (OCRD.ListNum)
                        // Observation 2: LineTotal doesn`t need to be setted because its automatically calculated (Quantity * Price)
                        orderSAP.Lines.BaseLine = line;
                        orderSAP.Lines.ItemCode = (string) orderItens["ItemCode"];
                        orderSAP.Lines.WarehouseCode = (string) item["DfltWH"];
                        orderSAP.Lines.Quantity = (int) orderItens["Quantity"];
                        orderSAP.Lines.TaxCode = "";
                        orderSAP.Lines.DiscountPercent = discountPrice;

                        line++;
                    }

                    // Order error
                    if (orderSAP.Add() != 0)
                    {
                        // If the currency exchange is not setted
                        if (company.GetLastErrorCode() == -4006)
                        {
                            throw new SAPObjectException("Currency Exchange rate hasn't been set", (int) order["DocEntry"]);
                        }
                        throw new SAPObjectException("Error to save order - " + company.GetLastErrorDescription(), (int) order["DocEntry"]);
                    }

                    // Order success
                    // Change status of this temporary order to Closed
                    SqlCommand commandUpdate = Database.ExecuteCommand("UPDATE [@Order] SET Status = @Status WHERE DocEntry = @DocEntry", connection);
                    commandUpdate.Parameters.AddWithValue("@Status", "C");
                    commandUpdate.Parameters.AddWithValue("@DocEntry", (int) order["DocEntry"]);
                    commandUpdate.ExecuteNonQuery();
                    LogWriter.write("Temporary order #" + order["DocEntry"] + " imported sucessfully! Generated order: #" + docNumNext + " (ORDR.DocNum)");

                } catch (System.Exception ex)
                {
                    string message = ex.Message;
                    if (ex is SAPObjectException)
                    {
                        SAPObjectException exSap = (SAPObjectException) ex;
                        message = exSap.getFormattedMessage();
                    }

                    // Change status of this temporary order to Error
                    SqlCommand commandUpdate = Database.ExecuteCommand("UPDATE [@Order] SET Status = @Status WHERE DocEntry = @DocEntry", connection);
                    commandUpdate.Parameters.AddWithValue("@Status", "E");
                    commandUpdate.Parameters.AddWithValue("@DocEntry", (int) order["DocEntry"]);
                    commandUpdate.ExecuteNonQuery();

                    // Add new error record
                    SqlCommand commandError = Database.ExecuteCommand("INSERT INTO [@OrderError] (DocEntry, Description, Date) VALUES (@DocEntry, @Description, @Date)", connection);
                    commandError.Parameters.AddWithValue("@DocEntry", (int) order["DocEntry"]);
                    commandError.Parameters.AddWithValue("@Description", message);
                    commandError.Parameters.AddWithValue("@Date", DateTime.Now);
                    commandError.ExecuteNonQuery();

                    // Send e-mail with error detail
                    Email.Send("Order error", message);

                    throw new System.Exception(message, ex);
                }
            }
        }
    }
}
