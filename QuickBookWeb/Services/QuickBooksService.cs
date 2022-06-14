using QuickBookWeb.Models;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Web;

namespace QuickBookWeb.Services
{
    public class QuickBooksService
    {
        #region Consts
        private const string InsertCustomerSql = "INSERT INTO customer (name, firstname, lastname, companyName, contact, accountNumber, BillAddressAddr1, BillAddressAddr2, BillAddressCity, BillAddressState, BillAddressPostalCode, BillAddressCountry, Phone, Fax, Email, TermsRefListID) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}',  '{15}')";

        //private const string InsertCustomerSql = "INSERT INTO customer (name, firstname, lastname, companyName, contact, accountNumber, BillAddressAddr1, BillAddressAddr2, BillAddressAddr3, BillAddressAddr4, BillAddressCity, BillAddressState, BillAddressPostalCode, BillAddressCountry, Phone, Fax, Email, TermsRefListID) VALUES ('{0}', '{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}',  '{15}', '{16}',  '{17}')";
        private const string InsertInvoiceLineItemSql = "INSERT INTO InvoiceLine (InvoiceLineItemRefListID, InvoiceLineDesc, InvoiceLineRate, InvoiceLineAmount,  FQSaveToCache) VALUES ('{0}', '{1}', 1.00000,  {2}, 1)";
        private const string InsertInvoiceSql = "INSERT INTO Invoice (CustomerRefListID, ARAccountRefListID, TxnDate,  RefNumber, IsPending, TermsRefListID, ShipDate, Memo) VALUES ('{0}', '{1}', {2}, '{3}',  0, '{4}', {5}, '{6}') ";
        private const string GetCustomerSql = "SELECT ListId, AccountNumber FROM Customer";
        private const string GetNewCustomerIdSql = "SP_LASTINSERTID customer";
        #endregion

        #region Fields
        private OdbcConnection _connection = null;
        #endregion

        #region construction / destruction
        public QuickBooksService()
        {
            _connection = new OdbcConnection("DSN=Calhono QuickBooks Data");
            //var str="DSN=Calhono QuickBooks Data;DFQ = C:\\Users\\Public\\Documents\\Intuit\\QuickBooks\\Company Files\\" + "ourQuickBooksFile.qbw;OLE DB Services=-2;OpenMode=S";
            //Server=App01;DSN=Calhono QuickBooks Data;
        }

        public void Dispose()
        {
            if (null != _connection)
            {
                if (_connection.State == ConnectionState.Open)
                {
                    _connection.Close();
                }
                _connection = null;
            }
        }
        #endregion

        #region IQuickBooksResourceProvider implementation
        public List<QBCustomerIdentifiers> GetCustomerIds()
        {
            string error = string.Empty;
            List<QBCustomerIdentifiers> customers = new List<QBCustomerIdentifiers>();
            OdbcCommand cmd = new OdbcCommand(GetCustomerSql, _connection);
            try
            {
                if (_connection.State == ConnectionState.Closed)
                {
                    _connection.Open();
                }
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }

            OdbcDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                bool success = false;
                QBCustomerIdentifiers ids = new QBCustomerIdentifiers();
                try
                {
                    ids.QBId = reader[0] as string;
                    ids.CHId = Convert.ToInt32(reader[1]);
                    success = true;
                }
                catch (Exception ex)
                {
                    error = string.Format("Failed to read Customer Ids record {0} - ignoring this record.", ids.QBId);
                    throw new Exception(error);
                }
                if (success)
                {
                    customers.Add(ids);
                }
            }

            return customers;
        }

        public QBCustomer AddCustomer(QBCustomer customer)
        {
            string error = string.Empty;
            
            try
            {
                string query = string.Format(InsertCustomerSql, SqlClean(customer.CompanyName), SqlClean(customer.FirstName), SqlClean(customer.LastName), SqlClean(customer.CompanyName), SqlClean(customer.Name), customer.AccountNumber, SqlClean(customer.Address1), SqlClean(customer.Address2), SqlClean(customer.City), customer.State, customer.ZipCode, customer.Country, customer.Phone, customer.Fax, SqlClean(customer.Email), customer.Terms);
                //string query = string.Format(InsertCustomerSql, SqlClean(customer.CompanyName), SqlClean(customer.FirstName), SqlClean(customer.LastName), SqlClean(customer.CompanyName), SqlClean(customer.Name), customer.AccountNumber, SqlClean(customer.Name), SqlClean(customer.CompanyName), SqlClean(customer.Address1), SqlClean(customer.Address2), SqlClean(customer.City), customer.State, customer.ZipCode, customer.Country, customer.Phone, customer.Fax, SqlClean(customer.Email), customer.Terms);
                OdbcCommand cmd = new OdbcCommand(query, _connection);

                if (_connection.State == ConnectionState.Closed)
                {
                    _connection.Open();
                }

                int result = cmd.ExecuteNonQuery();
                if (result > 0)
                {
                    cmd.CommandText = GetNewCustomerIdSql;
                    customer.QBId = cmd.ExecuteScalar() as string;
                }
            }
            catch (Exception ex)
            {
                error = string.Format("Could not add the Customer {0}. An Exception occurred Adding the Customer Record to QuickBooks: {1}", customer.AccountNumber, ex.Message);
                    
                throw new ApplicationException(error);
            }
            

            return customer;
        }

        public QBInvoice AddInvoice(QBInvoice invoice, QBCustomer customer)
        {
            string error = string.Empty;
            int result = 0;
            if (null != customer)
            {
                try
                {
                    string query = string.Format(InsertInvoiceLineItemSql, invoice.LineItemReferenceId, invoice.Description, invoice.Amount);
                    OdbcCommand cmd = new OdbcCommand(query, _connection);
                    if (_connection.State == ConnectionState.Closed)
                    {
                        _connection.Open();
                    }
                    result = cmd.ExecuteNonQuery();
                    if (result > 0)
                    {
                        query = string.Format(InsertInvoiceSql, invoice.CustomerReferenceId, invoice.AccountReferenceId, GetOdbcDate(invoice.Date), invoice.InvoiceNumber, invoice.Terms, GetOdbcDate(invoice.Date), invoice.Notes);
                        cmd.CommandText = query;
                        result = cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    error = string.Format("Could not add the Invoice {0} for Customer {1}. An Exception occurred Adding the Invoice Record: {2}", invoice.InvoiceNumber, customer.AccountNumber, ex.Message);
                    throw new Exception(error);

                }
            }
            else
            {
                error = string.Format("Could not add the Invoice {0} for Customer {1}. The Customer record supplied was null.", invoice.InvoiceNumber, customer.AccountNumber);
                throw new Exception(error);
            }            
            return invoice;
        }

        public QBStatement GetStatement(int companyId, DateTime dateFrom, DateTime dateTo)
        {
            return new QBStatement();
        }
        #endregion

        #region private methods
        private string GetOdbcDate(DateTime date)
        {
            return "{d'" + date.ToString("yyyy-MM-dd") + "'}";
        }

        private string SqlClean(string input)
        {
            string result = string.Empty;
            if (!string.IsNullOrEmpty(input))
            {
                result = input.Replace("'", "''");
            }

            return result;
        }

        public bool ConnectToQB()
        {
            List<string> items = new List<string>();
            OdbcConnection con = new OdbcConnection("DSN=Calhono QuickBooks Data");
            con.Open();
            OdbcDataAdapter dAdapter = new OdbcDataAdapter("SELECT CustomerRefFullName, RefNumber, TxnDate, BalanceRemaining, AppliedAmount, Memo FROM Invoice WHERE TxnDate > {d'2022-04-01'}", con);
            DataTable result = new DataTable();
            dAdapter.Fill(result);
            DataTableReader reader = new DataTableReader(result);
            while (reader.Read())
            {
                items.Add("Invoice #: " + reader.GetString(1));
            }
            con.Close();

            return true;
        }
        #endregion
    }
}