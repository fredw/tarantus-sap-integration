using System;
using System.Data.SqlClient;
using TarantusSAPService.Exception;

namespace TarantusSAPService.Task
{
    public static class Order
    {
        public static void Execute(SAPbobsCOM.Company company, SqlConnection connection)
        {
            // Select OPEN orders from temporary table
            SqlCommand commandOrder = Database.ExecuteCommand(
                "SELECT " +
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
            SqlDataReader orders = commandOrder.ExecuteReader();

            if (!orders.HasRows)
            {
                LogWriter.write("There is no open orders to import.");
                return;
            }

            while (orders.Read())
            {
                try
                {
                    // Customer
                    SAPbobsCOM.BusinessPartners customer = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                    // Try to find customer
                    if (!customer.GetByKey((string) orders["CardCode"]))
                    {
                        throw new SAPObjectException("Customer not found " + orders["CardCode"] + " (OCRD.CardCode)", (int) orders["DocEntry"]);
                    }

                    // Update customer price table and payment condition
                    customer.PriceListNum = (int) orders["PriceTable"];
                    customer.PayTermsGrpCode = (int) orders["PaymentCondition"];
                    customer.Update();

                    if (customer.Update() != 0)
                    {
                        throw new SAPObjectException("Error to update customer " + orders["CardCode"] + " - " + company.GetLastErrorDescription(), (int) orders["DocEntry"]);
                    }

                    // Get last DocNum from ORDR table
                    SqlCommand commandLastOrdr = Database.ExecuteCommand("SELECT TOP 1 DocNum FROM ORDR ORDER BY DocNum desc", connection);
                    int docNumNext = Int32.Parse(commandLastOrdr.ExecuteScalar().ToString()) + 1;

                    Double discountPrice = orders.IsDBNull(orders.GetOrdinal("DiscountPrice")) ? 0 : Double.Parse(((decimal) orders["DiscountPrice"]).ToString());

                    // Create the order object and set his properties
                    SAPbobsCOM.Documents order = company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                    order.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                    order.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
                    order.BPL_IDAssignedToInvoice = 1; // First business place (for multibranch companies)
                    order.DocNum = docNumNext;
                    order.CardCode = (string) orders["CardCode"];
                    order.PayToCode = (string) orders["AddressCharge"];
                    order.ShipToCode = (string) orders["AddressDelivery"];
                    order.DocDate = (DateTime) orders["DocDate"];
                    order.TaxDate = (DateTime) orders["DocDate"];
                    order.DocDueDate = (DateTime) orders["DeliveryDate"];
                    order.NumAtCard = orders.IsDBNull(orders.GetOrdinal("ClientOrderNum")) ? null : (string) orders["ClientOrderNum"];
                    order.Comments = (string) orders["Observation"];
                    order.SalesPersonCode = (int) orders["SlpCode"];
                    order.GroupNumber = (int) orders["PaymentCondition"];
                    order.TaxExtension.Carrier = orders.IsDBNull(orders.GetOrdinal("Carrier")) ? null : (string) orders["Carrier"];
                    order.TaxExtension.Incoterms = orders["Shipping"].ToString();
                    // Custom fields
                    order.UserFields.Fields.Item("U_LstNum").Value = (int) orders["PriceTable"];
                    order.UserFields.Fields.Item("U_Paciente").Value = ((int) orders["PriceTable"]).ToString();
                    order.UserFields.Fields.Item("U_Desc_Global").Value = discountPrice;

                    // Select temporary itens from this temporary order
                    SqlCommand commandOrderItens = Database.ExecuteCommand(
                        "SELECT DocEntry, LineNum, ItemCode, Quantity, PriceUnitFinal " +
                        "FROM [@OrderItem] " +
                        "WHERE DocEntry = @DocEntry",
                        connection
                    );
                    commandOrderItens.Parameters.AddWithValue("@DocEntry", (int) orders["DocEntry"]);
                    SqlDataReader orderItens = commandOrderItens.ExecuteReader();

                    int line = 0;
                    while (orderItens.Read())
                    {
                        if (line != 0)
                        {
                            // Create new item line
                            order.Lines.Add();
                        }

                        // Select item default wharehouse
                        SqlCommand commandItem = Database.ExecuteCommand("SELECT DfltWH FROM OITM WHERE ItemCode = @ItemCode", connection);
                        commandItem.Parameters.AddWithValue("@ItemCode", (string) orderItens["ItemCode"]);
                        SqlDataReader item = commandItem.ExecuteReader(System.Data.CommandBehavior.SingleRow);

                        // If not found item on database
                        if (!item.Read())
                        {
                            throw new SAPObjectException("Item not found " + orderItens["ItemCode"] + " (OITM.ItemCode)", (int) orders["DocEntry"]);
                        }
                        // If item doesn't have a default wharehouse defined
                        else if (item.IsDBNull(item.GetOrdinal("DfltWH")))
                        {
                            throw new SAPObjectException("Item " + orderItens["ItemCode"] + " hasn't default wharehouse defined (OITM.DfltWH)", (int) orders["DocEntry"]);
                        }

                        // Set item properties 
                        // Observation 1: Price doesn`t need to be setted because its loaded by customer price table (OCRD.ListNum)
                        // Observation 2: LineTotal doesn`t need to be setted because its automatically calculated (Quantity * Price)
                        order.Lines.BaseLine = line;
                        order.Lines.ItemCode = (string) orderItens["ItemCode"];
                        order.Lines.WarehouseCode = (string) item["DfltWH"];
                        order.Lines.Quantity = (int) orderItens["Quantity"];
                        order.Lines.TaxCode = "";
                        order.Lines.DiscountPercent = discountPrice;

                        line++;
                    }

                    // Order error
                    if (order.Add() != 0)
                    {
                        // If the currency exchange is not setted
                        if (company.GetLastErrorCode() == -4006)
                        {
                            throw new SAPObjectException("Currency Exchange rate hasn't been set", (int) orders["DocEntry"]);
                        }
                        throw new SAPObjectException("Error to save order - " + company.GetLastErrorDescription(), (int) orders["DocEntry"]);
                    }

                    // Order success
                    // Change status of this temporary order to Closed
                    SqlCommand commandUpdate = Database.ExecuteCommand("UPDATE [@Order] SET Status = @Status WHERE DocEntry = @DocEntry", connection);
                    commandUpdate.Parameters.AddWithValue("@Status", "C");
                    commandUpdate.Parameters.AddWithValue("@DocEntry", (int) orders["DocEntry"]);
                    commandUpdate.ExecuteNonQuery();
                    LogWriter.write("Temporary order #" + orders["DocEntry"] + " imported sucessfully! Generated order: #" + docNumNext + " (ORDR.DocNum)");

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
                    commandUpdate.Parameters.AddWithValue("@DocEntry", (int) orders["DocEntry"]);
                    commandUpdate.ExecuteNonQuery();

                    // Add new error record
                    SqlCommand commandError = Database.ExecuteCommand("INSERT INTO [@OrderError] (DocEntry, Description, Date) VALUES (@DocEntry, @Description, @Date)", connection);
                    commandError.Parameters.AddWithValue("@DocEntry", (int) orders["DocEntry"]);
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
