using System;
using System.Data.SqlClient;

namespace TarantusSAPService.Task
{
    public static class Order
    {
        public static void Execute(SAPbobsCOM.Company oCompany, SqlConnection Connection)
        {
            // Select OPEN orders from temporary table
            SqlCommand CommandOrder = Database.ExecuteCommand(
                "SELECT " +
                    "DocEntry, " +
                    "CardCode, " +
                    "CardCodeCharge, " +
                    "CardCodeDelivery, " +
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
                Connection
            );
            CommandOrder.Parameters.AddWithValue("@Status", "O");
            SqlDataReader orders = CommandOrder.ExecuteReader();

            if (!orders.HasRows)
            {
                LogWriter.write("There is no open orders to import.");
            } else
            {
                while (orders.Read())
                {
                    // Customer
                    SAPbobsCOM.BusinessPartners Customer = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oBusinessPartners);

                    // Try to find customer
                    if (!Customer.GetByKey(orders.GetString(orders.GetOrdinal("CardCode"))))
                    {
                        throw new System.Exception("Order #" + orders.GetInt32(orders.GetOrdinal("DocEntry")) + " customer not found: " + orders.GetString(orders.GetOrdinal("CardCode")));
                    } else
                    {
                        // Update customer price table
                        Customer.PriceListNum = orders.GetInt32(orders.GetOrdinal("PriceTable"));
                        Customer.Update();

                        if (Customer.Update() != 0)
                        {
                            throw new System.Exception("Order #" + orders.GetInt32(orders.GetOrdinal("DocEntry")) + " error to update customer price table: " + oCompany.GetLastErrorDescription());
                        }
                    }

                    // Get last DocNum from ORDR table
                    SqlCommand CommandLastOrdr = Database.ExecuteCommand("SELECT TOP 1 DocNum FROM ORDR ORDER BY DocNum desc", Connection);
                    int DocNumNext = Int32.Parse(CommandLastOrdr.ExecuteScalar().ToString()) + 1;

                    // Create the order object and set his properties
                    SAPbobsCOM.Documents Order = oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders);
                    Order.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Items;
                    Order.HandWritten = SAPbobsCOM.BoYesNoEnum.tNO;
                    Order.BPL_IDAssignedToInvoice = 1; // First business place (to multibranch companies)
                    Order.DocNum = DocNumNext;
                    Order.CardCode = orders.GetString(orders.GetOrdinal("CardCode"));
                    Order.PayToCode = orders.GetString(orders.GetOrdinal("CardCodeCharge"));
                    Order.ShipToCode = orders.GetString(orders.GetOrdinal("CardCodeDelivery"));
                    Order.DocDate = orders.GetDateTime(orders.GetOrdinal("DocDate"));
                    Order.TaxDate = orders.GetDateTime(orders.GetOrdinal("DocDate"));
                    Order.DocDueDate = orders.GetDateTime(orders.GetOrdinal("DeliveryDate"));
                    if (!orders.IsDBNull(orders.GetOrdinal("ClientOrderNum"))) {
                        Order.NumAtCard = orders.GetString(orders.GetOrdinal("ClientOrderNum"));
                    }                    
                    Order.Comments = orders.GetString(orders.GetOrdinal("Observation"));
                    Order.SalesPersonCode = orders.GetInt32(orders.GetOrdinal("SlpCode"));
                    Order.GroupNumber = orders.GetInt32(orders.GetOrdinal("PaymentCondition"));
                    Order.UserFields.Fields.Item("U_LstNum").Value = orders.GetInt32(orders.GetOrdinal("PriceTable"));
                    Order.UserFields.Fields.Item("U_Paciente").Value = orders.GetInt32(orders.GetOrdinal("PriceTable")).ToString();
                    if (!orders.IsDBNull(orders.GetOrdinal("DiscountPrice"))) {
                        Order.UserFields.Fields.Item("U_Desc_Global").Value = Double.Parse(orders.GetDecimal(orders.GetOrdinal("DiscountPrice")).ToString());
                    }
                    if (!orders.IsDBNull(orders.GetOrdinal("Carrier"))) {
                        Order.TaxExtension.Carrier = orders.GetString(orders.GetOrdinal("Carrier"));
                    }                    
                    Order.TaxExtension.Incoterms = orders.GetInt32(orders.GetOrdinal("Shipping")).ToString();
                    
                    // Select temporary itens from this temporary order
                    SqlCommand CommandOrderItens = Database.ExecuteCommand(
                        "SELECT DocEntry, LineNum, ItemCode, Quantity, PriceUnitFinal " +
                        "FROM [@OrderItem] " +
                        "WHERE DocEntry = @DocEntry",
                        Connection
                    );
                    CommandOrderItens.Parameters.AddWithValue("@DocEntry", orders.GetInt32(orders.GetOrdinal("DocEntry")));
                    SqlDataReader orderItens = CommandOrderItens.ExecuteReader();
                    
                    int line = 0;
                    while (orderItens.Read())
                    {
                        if (line != 0)
                        {
                            // Create new item line
                            Order.Lines.Add();
                        }

                        // Select item default wharehouse
                        SqlCommand CommandItem = Database.ExecuteCommand("SELECT DfltWH FROM OITM WHERE ItemCode = @ItemCode", Connection);
                        CommandItem.Parameters.AddWithValue("@ItemCode", orderItens.GetString(orderItens.GetOrdinal("ItemCode")));
                        SqlDataReader item = CommandItem.ExecuteReader(System.Data.CommandBehavior.SingleRow);

                        if (!item.Read())
                        {
                            throw new System.Exception("Order #" + orders.GetInt32(orders.GetOrdinal("DocEntry")) + " error: item not found on OITM: " + orderItens.GetString(orderItens.GetOrdinal("ItemCode")));
                        }

                        // Set item properties 
                        // Observation 1: Price doesn`t need to be setted because its loaded by customer price table (OCRD.ListNum)
                        // Observation 2: LineTotal doesn`t need to be setted because its automatically calculated (Quantity * Price)
                        Order.Lines.BaseLine = line;
                        Order.Lines.ItemCode = orderItens.GetString(orderItens.GetOrdinal("ItemCode"));
                        Order.Lines.WarehouseCode = item.GetString(item.GetOrdinal("DfltWH"));
                        Order.Lines.Quantity = orderItens.GetInt32(orderItens.GetOrdinal("Quantity"));
                        Order.Lines.TaxCode = "";
                        if (!orders.IsDBNull(orders.GetOrdinal("DiscountPrice")))
                        {
                            Order.Lines.DiscountPercent = Double.Parse(orders.GetDecimal(orders.GetOrdinal("DiscountPrice")).ToString());
                        }
                        
                        line++;
                    }
                    
                    // Order error
                    if (Order.Add() != 0)
                    {
                        // If the currency exchange is not setted
                        if (oCompany.GetLastErrorCode() == -4006)
                        {
                            throw new System.Exception("Order #" + orders.GetInt32(orders.GetOrdinal("DocEntry")) + " error: Currency Exchange - exchange rate has not been set for today. set the exchange rate");
                        }
                        throw new System.Exception("Order #" + orders.GetInt32(orders.GetOrdinal("DocEntry")) + " error: " + oCompany.GetLastErrorDescription());
                    // Order success
                    } else
                    {
                        // Change status of this temporary order to Closed
                        SqlCommand CommandUpdate = Database.ExecuteCommand("UPDATE [@Order] SET Status = @Status WHERE DocEntry = @DocEntry", Connection);
                        CommandUpdate.Parameters.AddWithValue("@Status", "C");
                        CommandUpdate.Parameters.AddWithValue("@DocEntry", orders.GetInt32(orders.GetOrdinal("DocEntry")));
                        CommandUpdate.ExecuteNonQuery();
                        // Log success
                        LogWriter.write("Order #" + DocNumNext.ToString() + " (ORDR.DocNum) imported sucessfully!");
                    }
                }
            }
        }
    }
}
