using Interop.QBXMLRP2;
using QuickBookWeb.Models;
using System;
using System.Collections.Generic;
using System.Xml;

namespace QuickBookWeb.Services
{
    public class QuickBooksSDKService
    {        

        #region IQuickBooksResourceProvider implementation
        public List<QBCustomerIdentifiers> GetCustomerIds()
        {
            string error = string.Empty;
            List<QBCustomerIdentifiers> customers = new List<QBCustomerIdentifiers>();
            
            try
            {
                
            }
            catch (Exception ex)
            {
                throw new ApplicationException(ex.Message);
            }
            

            return customers;
        }

        public QBCustomer AddCustomer(QBCustomer customer)
        {
            string error = string.Empty;
            
            try
            {
                customer.QBId = SendAddCustomerQbXMLReq(customer);
            }
            catch (Exception ex)
            {
                error = string.Format("Could not add the Customer {0}. An Exception occurred Adding the Customer Record to QuickBooks: {1}", customer.AccountNumber, ex.Message);
                    
                throw new ApplicationException(error);
            }
            
            return customer;
        }

        private string SendAddCustomerQbXMLReq(QBCustomer customer)
        {
            //step2: create the qbXML request
            XmlDocument inputXMLDoc = new XmlDocument();
            try
            {                
                inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
                inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"2.0\""));
                XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
                inputXMLDoc.AppendChild(qbXML);
                XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
                qbXML.AppendChild(qbXMLMsgsRq);
                qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
                XmlElement custAddRq = inputXMLDoc.CreateElement("CustomerAddRq");
                qbXMLMsgsRq.AppendChild(custAddRq);
                custAddRq.SetAttribute("requestID", "1");
                XmlElement custAdd = inputXMLDoc.CreateElement("CustomerAdd");
                custAddRq.AppendChild(custAdd);
                custAdd.AppendChild(inputXMLDoc.CreateElement("Name")).InnerText = customer.CompanyName;
                custAdd.AppendChild(inputXMLDoc.CreateElement("CompanyName")).InnerText = customer.CompanyName;
                custAdd.AppendChild(inputXMLDoc.CreateElement("FirstName")).InnerText = customer.FirstName;
                custAdd.AppendChild(inputXMLDoc.CreateElement("LastName")).InnerText = customer.LastName;
                custAdd.AppendChild(inputXMLDoc.CreateElement("IsActive")).InnerText = "1";

                custAdd.AppendChild(inputXMLDoc.CreateElement("Contact")).InnerText = customer.Name;
                custAdd.AppendChild(inputXMLDoc.CreateElement("AccountNumber")).InnerText = customer.AccountNumber;
                if (!string.IsNullOrEmpty(customer.Phone))
                    custAdd.AppendChild(inputXMLDoc.CreateElement("Phone")).InnerText = customer.Phone;
                if (!string.IsNullOrEmpty(customer.Fax))
                    custAdd.AppendChild(inputXMLDoc.CreateElement("AltPhone")).InnerText = customer.Fax;
                if (!string.IsNullOrEmpty(customer.Email))
                    custAdd.AppendChild(inputXMLDoc.CreateElement("Email")).InnerText = customer.Email;

                XmlElement billAddr = inputXMLDoc.CreateElement("BillAddress");
                custAdd.AppendChild(billAddr);
                if (!string.IsNullOrEmpty(customer.Address1))
                    billAddr.AppendChild(inputXMLDoc.CreateElement("Addr1")).InnerText = customer.Address1;
                if (!string.IsNullOrEmpty(customer.Address2))
                    billAddr.AppendChild(inputXMLDoc.CreateElement("Addr2")).InnerText = customer.Address2;
                if (!string.IsNullOrEmpty(customer.City))
                    billAddr.AppendChild(inputXMLDoc.CreateElement("City")).InnerText = customer.City;
                if (!string.IsNullOrEmpty(customer.State))
                    billAddr.AppendChild(inputXMLDoc.CreateElement("State")).InnerText = customer.State;
                if (!string.IsNullOrEmpty(customer.ZipCode))
                    billAddr.AppendChild(inputXMLDoc.CreateElement("PostalCode")).InnerText = customer.ZipCode;
                if (!string.IsNullOrEmpty(customer.Country))
                    billAddr.AppendChild(inputXMLDoc.CreateElement("Country")).InnerText = customer.Country;

                if (!string.IsNullOrEmpty(customer.Terms))
                {
                    XmlElement termsRef = inputXMLDoc.CreateElement("TermsRef");
                    custAdd.AppendChild(termsRef);
                    termsRef.AppendChild(inputXMLDoc.CreateElement("ListID")).InnerText = customer.Terms;
                }
            }
            catch (Exception exe)
            {
                throw new Exception("Failed to build QBXML request message, Error Message: " + exe.Message);
            }

            string input = inputXMLDoc.OuterXml;
            //step3: do the qbXMLRP request
            RequestProcessor2 rp = null;
            string ticket = null;
            string response = null;
            try
            {
                rp = new RequestProcessor2();
                rp.OpenConnection("", "QuickBook Web Service for FMS2");
                ticket = rp.BeginSession("", QBFileMode.qbFileOpenDoNotCare);
                response = rp.ProcessRequest(ticket, input);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                throw new Exception("Could not connected to QuickBook, Error Message: " + ex.Message);                
            }
            finally
            {
                if (ticket != null)
                {
                    rp.EndSession(ticket);
                }
                if (rp != null)
                {
                    rp.CloseConnection();
                }
            };

            string listID = string.Empty;
            //step4: parse the XML response and show a message
            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);
            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("CustomerAddRs");

            if (qbXMLMsgsRsNodeList.Count == 1) //it's always true, since we added a single Customer
            {                
                XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
                //get the status Code, info and Severity
                string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;
                if ("0".Equals(retStatusCode))
                {
                    //get the CustomerRet node for detailed info
                    //a CustomerAddRs contains max one childNode for "CustomerRet"
                    XmlNodeList custAddRsNodeList = qbXMLMsgsRsNodeList.Item(0).ChildNodes;
                    if (custAddRsNodeList.Count == 1 && custAddRsNodeList.Item(0).Name.Equals("CustomerRet"))
                    {
                        XmlNodeList custRetNodeList = custAddRsNodeList.Item(0).ChildNodes;

                        foreach (XmlNode custRetNode in custRetNodeList)
                        {
                            if (custRetNode.Name.Equals("ListID"))
                            {
                                listID = custRetNode.InnerText;
                                break;
                            }
                        }
                    } // End of customerRet
                }
                else
                {
                    throw new Exception("The QuickBooks server processed your request unsuccessfully, Error Message: "+ retStatusMessage);
                }
                

                
            } //End of customerAddRs
            return listID;
        }


        public QBInvoice AddInvoice(QBInvoice invoice, QBCustomer customer)
        {
            string error = string.Empty;
            int result = 0;
            if (null != customer)
            {
                try
                {
                    SendAddInvoiceQbXMLReq(invoice,customer);
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

        private void SendAddInvoiceQbXMLReq(QBInvoice invoice, QBCustomer customer)
        {
            //step2: create the qbXML request
            XmlDocument inputXMLDoc = new XmlDocument();
            try
            {
                inputXMLDoc.AppendChild(inputXMLDoc.CreateXmlDeclaration("1.0", null, null));
                inputXMLDoc.AppendChild(inputXMLDoc.CreateProcessingInstruction("qbxml", "version=\"2.0\""));
                XmlElement qbXML = inputXMLDoc.CreateElement("QBXML");
                inputXMLDoc.AppendChild(qbXML);
                XmlElement qbXMLMsgsRq = inputXMLDoc.CreateElement("QBXMLMsgsRq");
                qbXML.AppendChild(qbXMLMsgsRq);
                qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
                XmlElement invoiceAddRq = inputXMLDoc.CreateElement("InvoiceAddRq");
                qbXMLMsgsRq.AppendChild(invoiceAddRq);
                invoiceAddRq.SetAttribute("requestID", "1");

                XmlElement invoiceAdd = inputXMLDoc.CreateElement("InvoiceAdd");                
                invoiceAddRq.AppendChild(invoiceAdd);

                invoiceAdd.SetAttribute("defMacro", "TxnID:NewInvoice");
                invoiceAdd.AppendChild(inputXMLDoc.CreateElement("TxnDate")).InnerText = invoice.Date.ToString("yyyy-MM-0dd");
                invoiceAdd.AppendChild(inputXMLDoc.CreateElement("RefNumber")).InnerText = invoice.InvoiceNumber;
                invoiceAdd.AppendChild(inputXMLDoc.CreateElement("IsPending")).InnerText = "0";
                invoiceAdd.AppendChild(inputXMLDoc.CreateElement("ShipDate")).InnerText = invoice.ShipDate.ToString("yyyy-MM-0dd");
                invoiceAdd.AppendChild(inputXMLDoc.CreateElement("Memo")).InnerText = invoice.Notes;

                XmlElement customerRef = inputXMLDoc.CreateElement("CustomerRef");
                invoiceAdd.AppendChild(customerRef);
                customerRef.AppendChild(inputXMLDoc.CreateElement("ListID")).InnerText = invoice.CustomerReferenceId;

                XmlElement arAccountRef = inputXMLDoc.CreateElement("ARAccountRef");
                invoiceAdd.AppendChild(arAccountRef);
                arAccountRef.AppendChild(inputXMLDoc.CreateElement("ListID")).InnerText = invoice.AccountReferenceId;

                XmlElement termsRef = inputXMLDoc.CreateElement("TermsRef");
                invoiceAdd.AppendChild(termsRef);
                termsRef.AppendChild(inputXMLDoc.CreateElement("ListID")).InnerText = invoice.Terms;


                XmlElement invoiceLineAdd = inputXMLDoc.CreateElement("InvoiceLineAdd");
                invoiceAdd.AppendChild(invoiceLineAdd);
                invoiceLineAdd.AppendChild(inputXMLDoc.CreateElement("Desc")).InnerText = invoice.Description;
                invoiceLineAdd.AppendChild(inputXMLDoc.CreateElement("Amount")).InnerText = invoice.Amount.ToString();
                invoiceLineAdd.AppendChild(inputXMLDoc.CreateElement("Rate")).InnerText = "1.00";

                XmlElement itemRef = inputXMLDoc.CreateElement("ItemRef");
                invoiceLineAdd.AppendChild(itemRef);
                itemRef.AppendChild(inputXMLDoc.CreateElement("ListID")).InnerText = invoice.LineItemReferenceId;
                
            }
            catch (Exception exe)
            {
                throw new Exception("Failed to build QBXML request message, Error Message: " + exe.Message);
            }

            string input = inputXMLDoc.OuterXml;
            //step3: do the qbXMLRP request
            RequestProcessor2 rp = null;
            string ticket = null;
            string response = null;
            try
            {
                rp = new RequestProcessor2();
                rp.OpenConnection("", "QuickBook Web Service for FMS2");
                ticket = rp.BeginSession("Cal Hono Freight Fowarders.qbw", QBFileMode.qbFileOpenDoNotCare);
                response = rp.ProcessRequest(ticket, input);

            }
            catch (System.Runtime.InteropServices.COMException ex)
            {
                throw new Exception("Could not connected to QuickBook, Error Message: " + ex.Message);
            }
            finally
            {
                if (ticket != null)
                {
                    rp.EndSession(ticket);
                }
                if (rp != null)
                {
                    rp.CloseConnection();
                }
            };

            string listID = string.Empty;
            //step4: parse the XML response and show a message
            XmlDocument outputXMLDoc = new XmlDocument();
            outputXMLDoc.LoadXml(response);
            XmlNodeList qbXMLMsgsRsNodeList = outputXMLDoc.GetElementsByTagName("InvoiceAddRs");

            if (qbXMLMsgsRsNodeList.Count == 1) //it's always true, since we added a single Customer
            {
                XmlAttributeCollection rsAttributes = qbXMLMsgsRsNodeList.Item(0).Attributes;
                //get the status Code, info and Severity
                string retStatusCode = rsAttributes.GetNamedItem("statusCode").Value;
                string retStatusSeverity = rsAttributes.GetNamedItem("statusSeverity").Value;
                string retStatusMessage = rsAttributes.GetNamedItem("statusMessage").Value;
                if (!"0".Equals(retStatusCode))
                {                  
                    throw new Exception("The QuickBooks server processed your request unsuccessfully, Error Message: " + retStatusMessage);
                }

               
            } 
            
        }


        private string GetOdbcDate(DateTime date)
        {
            return "{d'" + date.ToString("yyyy-MM-dd") + "'}";
        }
        #endregion


    }
}