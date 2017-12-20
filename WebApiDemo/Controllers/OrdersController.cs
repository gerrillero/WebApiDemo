using Syncfusion.EJ.Export;
using Syncfusion.JavaScript.Models;
using Syncfusion.XlsIO;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.Entity.Infrastructure;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Reflection;
using System.Web;
using System.Web.Http;
using System.Web.Http.Description;
using System.Web.Http.OData;
using System.Web.Http.OData.Query;
using System.Web.Script.Serialization;
using WebApiDemo.Models;

namespace WebApiDemo.Controllers
{
    public class OrdersController : ApiController
    {
        private northwindEntities db = new northwindEntities();

        //public IQueryable<Order> GetOrders()
        //{
        //    List<Order> orders = new List<Order>();

        //    foreach (var order in db.Orders)
        //    {
        //        Order neworder = new Order();
        //        neworder.OrderID = order.OrderID;
        //        neworder.CustomerID = order.CustomerID;
        //        neworder.Freight = order.Freight;
        //        neworder.OrderDate = order.OrderDate;
        //        neworder.ShipCity = order.ShipCity;
        //        orders.Add(neworder);
        //    }
        //    return orders.AsQueryable();
        //}

        public IHttpActionResult Get(ODataQueryOptions<Order> opts)
        {
            List<Order> emp = db.Orders.ToList();

            var result = new PageResult<Order>(opts.ApplyTo(emp.AsQueryable()) as IEnumerable<Order>, null, emp.Count);

            //var result = emp;

            return Ok(result);
        }


        [System.Web.Http.ActionName("ExcelExport")]
        [AcceptVerbs("POST")]
        public void ExcelExport()
        {
            string gridModel = HttpContext.Current.Request.Params["GridModel"];
            GridProperties gridProperty = ConvertGridObject(gridModel);
            ExcelExport exp = new ExcelExport();
            IEnumerable<Order> result = db.Orders.ToList();
            exp.Export(gridProperty, result, "Export.xlsx", ExcelVersion.Excel2010, false, false, "flat-saffron");
        }

        private GridProperties ConvertGridObject(string gridProperty)
        {
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            IEnumerable div = (IEnumerable)serializer.Deserialize(gridProperty, typeof(IEnumerable));
            GridProperties gridProp = new GridProperties();
            foreach (KeyValuePair<string, object> datasource in div)
            {
                var property = gridProp.GetType()
                    .GetProperty(datasource.Key, BindingFlags.Instance | BindingFlags.Public | BindingFlags.IgnoreCase);
                if (property != null)
                {
                    Type type = property.PropertyType;
                    string serialize = serializer.Serialize(datasource.Value);
                    object value = serializer.Deserialize(serialize, type);
                    property.SetValue(gridProp, value, null);
                }
            }
            return gridProp;
        }


    }
}