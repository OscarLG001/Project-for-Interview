//OscarLaytonGonzalez
using System;
using System.Data;
using System.Data.OleDb;
using System.Xml.Linq;
using System.Globalization;

class Program
{
    static void Main()
    {
        string inputFile = "BillFile.xml";
        string outputFile = $"BillFile-{DateTime.Now:MMddyyyy}.rpt";



        //Creation of rpt File header
        string outputFormat = "1~FR|2~[{0}]|3~Sample UT file|4~[{1}]|5~[{2}]|6~[{3}]";

        XDocument xml = XDocument.Load(inputFile);

        string clientGuid = "8203ACC7-2094-43CC-8F7A-B8F19AA9BDA2";
        string currentDate = DateTime.Now.ToString("MM/dd/yyyy");
        string invoiceRecordCount = xml.Descendants("BILL_HEADER").Count().ToString();
        string invoiceRecordTotalAmount = GetInvoiceRecordTotalAmount(xml);

        string fileHeader = string.Format(outputFormat, clientGuid, currentDate, invoiceRecordCount, invoiceRecordTotalAmount);
        WriteToFile(outputFile, fileHeader);




        //Connection for Access database for Billing.mdb
        string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Billing.mdb"; // Replace with your Access database file path.

        using (OleDbConnection connection = new OleDbConnection(connectionString))
        {
            connection.Open();









//BillFile.xml parse and .rpt write

var invoices = xml.Descendants("BILL_HEADER");
foreach (var invoice in invoices)
{
    // Extract data fields from the invoice element
    string Account_No = GetFieldValue(invoice, "Account_No");
    string Customer_Name = GetFieldValue(invoice, "Customer_Name");
    string Mailing_Address_1 = GetFieldValue(invoice, "Mailing_Address_1");
    string Mailing_Address_2 = GetFieldValue(invoice, "Mailing_Address_2");
    string City = GetFieldValue(invoice, "City");
    string State = GetFieldValue(invoice, "State");
    string Zip = GetFieldValue(invoice, "Zip");

    string invoiceFormat = "8E2FEA69-5D77-4D0F-898E-DFA25677D19E";
    string Invoice_No = GetFieldValue(invoice, "Invoice_No");



    //Parse exactly into format MM/dd/yyyy
    
    string BillDt = GetFieldValue(invoice, "Bill_Dt");
    DateTime BillDate = DateTime.ParseExact(BillDt, "MMM-dd-yyyy", CultureInfo.InvariantCulture);
    string Bill_Dt = BillDate.ToString("MM/dd/yyyy");

    string DueDt = GetFieldValue(invoice, "Due_Dt");
    DateTime DueDate = DateTime.ParseExact(DueDt, "MMM-dd-yyyy", CultureInfo.InvariantCulture);
    string Due_Dt = DueDate.ToString("MM/dd/yyyy");

    string Bill_Amount = GetFieldValue(invoice, "Bill_Amount");

    // Parse the Bill_Dt string to a DateTime object
    DateTime BillfivedaysDate = DateTime.ParseExact(BillDt, "MMM-dd-yyyy", CultureInfo.InvariantCulture);
    // Add 5 days from the BillDate
    DateTime first_NotificationDate = BillfivedaysDate.AddDays(5);
    // Format the newDate back to the "MM/dd/yyyy" format
    string firstNotificationDate = first_NotificationDate.ToString("MM/dd/yyyy");

    // Parse the Due_Dt string to a DateTime object
    DateTime dueDate = DateTime.ParseExact(DueDt, "MMM-dd-yyyy", CultureInfo.InvariantCulture);
    // Subtract 3 days from the dueDate
    DateTime second_NotificationDate = dueDate.AddDays(-3);
    // Format the newDate back to the "MM/dd/yyyy" format
    string secondNotificationDate = second_NotificationDate.ToString("MM/dd/yyyy");

    string Balance_Due = GetFieldValue(invoice, "Balance_Due");
    string serviceAddress = "30 Braintree Hl Office Park Ste 303, Braintree, Massachusetts, 02184, United States"; // Service Address which would be invoice cloud


    // Remove commas from number fields
    string Bill_AmountWithoutCommas = Bill_Amount.Replace(",", string.Empty);
    string Balance_DueWithoutCommas = Balance_Due.Replace(",", string.Empty);



    // Replace placeholders in the template with actual data
    string invoiceRecord = $"AA~CT|BB~[{Account_No}]|VV~[{Customer_Name}]|CC~[{Mailing_Address_1}]|DD~[{Mailing_Address_2}]|EE~[{City}]|FF~[{State}]|GG~[{Zip}]" +
        $"|HH~IH|II~R|JJ~[{invoiceFormat}]|KK~[{Invoice_No}]|LL~[{Bill_Dt}]|MM~[{Due_Dt}]|NN~[{Bill_AmountWithoutCommas}]|OO~[{firstNotificationDate}]" +
        $"|PP~[{secondNotificationDate}]|QQ~[{Balance_DueWithoutCommas}]|RR~[{currentDate}]|SS~[{serviceAddress}]";

    // Write the invoice record to the output file
    WriteToFile(outputFile, invoiceRecord);

    






                // Get the ID of the last inserted customer
                int customerId;
                using (OleDbCommand cmd = new OleDbCommand("SELECT @@IDENTITY", connection))
                {
                    customerId = Convert.ToInt32(cmd.ExecuteScalar());
                }






                // Access Database Billing.mdb
                // Insert data into the "Bills" table
                string insertBillQuery = "INSERT INTO Bills (BillDate, BillNumber, BillAmount, FormatGUID, AccountBalance, DueDate, ServiceAddress, FirstEmailDate, SecondEmailDate, DateAdded, CustomerID) " +
                                          "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)";
                using (OleDbCommand cmd = new OleDbCommand(insertBillQuery, connection))
                {
                    cmd.Parameters.Add("BillDate", OleDbType.VarChar).Value = Bill_Dt;
                    cmd.Parameters.Add("BillNumber", OleDbType.VarChar).Value = Account_No;
                    cmd.Parameters.Add("BillAmount", OleDbType.VarChar).Value = Bill_AmountWithoutCommas;
                    cmd.Parameters.Add("FormatGUID", OleDbType.VarChar).Value = "8E2FEA69-5D77-4D0F-898E-DFA25677D19E";
                    cmd.Parameters.Add("AccountBalance", OleDbType.VarChar).Value = Balance_DueWithoutCommas;
                    cmd.Parameters.Add("DueDate", OleDbType.VarChar).Value = Due_Dt;
                    cmd.Parameters.Add("ServiceAddress", OleDbType.VarChar).Value = "30 Braintree Hl Office Park Ste 303, Braintree, Massachusetts, 02184, United States";
                    cmd.Parameters.Add("FirstEmailDate", OleDbType.VarChar).Value = firstNotificationDate;
                    cmd.Parameters.Add("SecondEmailDate", OleDbType.VarChar).Value = secondNotificationDate;
                    cmd.Parameters.Add("DateAdded", OleDbType.VarChar).Value = DateTime.Now;
                    cmd.Parameters.Add("CustomerID", OleDbType.VarChar).Value = customerId;
                    
                    try
                    {
                        cmd.ExecuteNonQuery();
                    }
                    catch (OleDbException ex)
                    {
                        Console.WriteLine("Customer insert error: " + ex.Message);
                    }
                }

    

    // Insert data into the "Customer" table
string insertCustomerQuery = "INSERT INTO Customer (CustomerName, AccountNumber, CustomerAddress, CustomerCity, CustomerState, CustomerZip, DateAdded) " +
                              "VALUES (?, ?, ?, ?, ?, ?, ?)";
using (OleDbCommand cmd = new OleDbCommand(insertCustomerQuery, connection))
{
    // Add parameters with explicit data types
    cmd.Parameters.Add("CustomerName", OleDbType.VarChar).Value = Customer_Name;
    cmd.Parameters.Add("AccountNumber", OleDbType.VarChar).Value = Account_No;
    cmd.Parameters.Add("CustomerAddress", OleDbType.VarChar).Value = $"{Mailing_Address_1} {Mailing_Address_2}";
    cmd.Parameters.Add("CustomerCity", OleDbType.VarChar).Value = City;
    cmd.Parameters.Add("CustomerState", OleDbType.VarChar).Value = State;
    cmd.Parameters.Add("CustomerZip", OleDbType.VarChar).Value = Zip;
    cmd.Parameters.Add("DateAdded", OleDbType.Date).Value = DateTime.Now;

    try
    {
        cmd.ExecuteNonQuery();
    }
    catch (OleDbException ex)
    {
        Console.WriteLine("Customer insert error: " + ex.Message);
    }
}



} //End of loop for each invoice



        // Showing in terminal that file was imported and created
        Console.WriteLine($"{outputFile} created successfully!");
           
           
            connection.Close();
        } //End of access database connection


    

        // Showing in terminal that data was imported successfully into Billing.mdb
        Console.WriteLine($"Data from {outputFile} to Billing.mdb imported successfully!");



        

        




        //CSV File BillingReport.csv 


        string outputFilecsv = "BillingReport.csv";

        // Create the CSV header
        string csvHeader = "Customer.ID,Customer.CustomerName,Customer.AccountNumber,Customer.CustomerAddress,Customer.CustomerCity,Customer.CustomerState,Customer.CustomerZip,Bills.ID,Bills.BillDate,Bills.BillNumber,Bills.AccountBalance,Bills.DueDate,Bills.BillAmount,Bills.FormatGUID,Customer.DateAdded";

        // Write the CSV header to the output file
        WriteToFile(outputFilecsv, csvHeader);

        string connectionString2 = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=Billing.mdb"; //Access database file path.

        using (OleDbConnection connection2 = new OleDbConnection(connectionString2))
        {
            connection2.Open();

            // Retrieve data from both tables
            string query = "SELECT Customer.*, Bills.* FROM Customer LEFT JOIN Bills ON Customer.ID = Bills.ID";

            using (OleDbCommand cmd = new OleDbCommand(query, connection2))
            using (OleDbDataReader reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    // Extract data from the reader
                    string customerID = reader["customer.ID"].ToString();
                    string customerName = reader["CustomerName"].ToString().Replace(",", ""); // Remove commas from the name
                    string accountNumber = reader["AccountNumber"].ToString();
                    string customerAddress = reader["CustomerAddress"].ToString();
                    string customerCity = reader["CustomerCity"].ToString();
                    string customerState = reader["CustomerState"].ToString();
                    string customerZip = reader["CustomerZip"].ToString();
                    string billID = reader["Bills.ID"].ToString();
                    DateTime billDate = reader.GetDateTime(reader.GetOrdinal("BillDate"));
                    string billNumber = reader["BillNumber"].ToString();
                    decimal accountBalance = reader.GetDecimal(reader.GetOrdinal("AccountBalance"));
                    DateTime dueDate = reader.GetDateTime(reader.GetOrdinal("DueDate"));
                    decimal billAmount = reader.GetDecimal(reader.GetOrdinal("BillAmount"));
                    string formatGUID = reader["FormatGUID"].ToString();
                    DateTime dateAdded = reader.GetDateTime(reader.GetOrdinal("Bills.DateAdded"));

                    // Format the data as a CSV line
                    string csvLine = $"{customerID},{customerName},{accountNumber},{customerAddress},{customerCity},{customerState},{customerZip}," +
                        $"{billID},{billDate:MM/dd/yyyy},{billNumber},{accountBalance},{dueDate:MM/dd/yyyy},{billAmount},{formatGUID},{dateAdded:MM/dd/yyyy}";

                    // Write the CSV line to the output file
                    WriteToFile(outputFilecsv, csvLine);
                }
            }

            connection2.Close();
        }

        // Showing in the terminal that the file was created
        Console.WriteLine($"Data exported to {outputFilecsv} successfully!");












    } // End of Main Function






    static string GetFieldValue(XContainer element, string fieldName)
    {
        var field = element.Descendants()
            .FirstOrDefault(e => e.Name.LocalName.Equals(fieldName, StringComparison.OrdinalIgnoreCase));

        return field?.Value ?? string.Empty;
    }

    static string GetInvoiceRecordTotalAmount(XDocument xml)
    {
        decimal totalAmount = xml.Descendants("Bill_Amount")
            .Sum(billAmount => decimal.Parse(billAmount.Value));

        return totalAmount.ToString("F2");
    }

    static void WriteToFile(string fileName, string content)
    {
        using (StreamWriter writer = new StreamWriter(fileName, true))
        {
            writer.WriteLine(content);
        }
    }


    




  
}