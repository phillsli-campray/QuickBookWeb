using Newtonsoft.Json;
using QuickBookWeb.Models;
using QuickBookWeb.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web.Http;
using System.Web.Http.Results;

namespace QuickBookWeb.Controllers
{
    public class QuickBookController : ApiController
    {
        [HttpGet]
        public IHttpActionResult GetCustomerIds()
        {
            try
            {
                var qbService = new QuickBooksService();
                List<QBCustomerIdentifiers>  ids = qbService.GetCustomerIds();
                return Json(ids);
            }
            catch (Exception exe)
            {
                return BadRequest(exe.Message);
            }
        }
        
        [HttpPost]
        public IHttpActionResult AddCustomer()
        {
            try
            {
                string value = Request.Content.ReadAsStringAsync().Result;
                QBCustomer customer = JsonConvert.DeserializeObject<QBCustomer>(value);               
                var qbService = new QuickBooksService();
                customer=qbService.AddCustomer(customer);
                return Json(customer);
            }
            catch(Exception exe)
            {
                return BadRequest(exe.Message);                
            }
        }

        [HttpPost]
        public IHttpActionResult AddInvoice(string id)
        {
            try
            {
                string value = Request.Content.ReadAsStringAsync().Result;
                QBInvoice invoice = JsonConvert.DeserializeObject<QBInvoice>(value);
                QBCustomer customer = new QBCustomer();
                customer.AccountNumber = id;
                var qbService = new QuickBooksService();
                invoice = qbService.AddInvoice(invoice, customer);
                return Json(invoice);
            }
            catch (Exception exe)
            {
                return BadRequest(exe.Message);
            }
        }
    }
}
