using System;
using System.Collections.Generic;
using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Newtonsoft.Json;
using CsvHelper;
using System.IO;
using System.Globalization;
using Microsoft.Extensions.Configuration;
using System.Reflection;

namespace AccountingServices
{
    public static class Oberlo
    {
        public static List<OberloOrder> ScrapeOberloOrders(IMyConfiguration configuration, DateTime from, DateTime to)
        {
            var oberloOrders = new List<OberloOrder>();

            string userDataDir = configuration.GetValue("UserDataDir");
            string oberloUsername = configuration.GetValue("OberloUsername");
            string oberloPassword = configuration.GetValue("OberloPassword");

            ChromeOptions options = new ChromeOptions();
            string userDataArgument = string.Format("user-data-dir={0}", userDataDir);
            options.AddArguments(userDataArgument);
            options.AddArguments("--start-maximized");
            options.AddArgument("--log-level=3");
            //options.AddArgument("--headless");

            //string chromeDriverExePath = Utils.GetFilePathRelativeToAssembly(@"..\..\..\..\AccountingServices\bin\debug\netcoreapp2.0");
            string chromeDriverExePath = configuration.GetValue("ChromeDriverExePath");
            IWebDriver driver = new ChromeDriver(chromeDriverExePath, options);

            // https://app.oberlo.com/orders?from=2017-01-01&to=2017-12-31&page=1
            int page = 1;
            string url = string.Format("https://app.oberlo.com/orders?from={0:yyyy-MM-dd}&to={1:yyyy-MM-dd}&page={2}", from, to, page);
            driver.Navigate().GoToUrl(url);

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            var ready = wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            // login if login form is present
            if (SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@name='email']"))
                && SeleniumUtils.IsElementPresent(driver, By.XPath("//input[@name='password']")))
            {
                IWebElement username = driver.FindElement(By.XPath("//input[@name='email']"));
                IWebElement password = driver.FindElement(By.XPath("//input[@name='password']"));

                username.Clear();
                username.SendKeys(oberloUsername);

                // if the above crash, it might very well be Chrome and ChromeDriver incompability
                // please update the ChromeDriver to fit the Chrome version used.

                password.Clear();
                password.SendKeys(oberloPassword);

                // use password field to submit form
                password.Submit();

                var wait2 = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
                var ready2 = wait2.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));
            }

            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            var json = js.ExecuteScript("return window.App.payload.orders;");

            // convert to json dynamic object
            string jsonString = JsonConvert.SerializeObject(json);
            dynamic jsonD = JsonConvert.DeserializeObject(jsonString);

            // https://app.oberlo.com/orders?from=2017-01-01&to=2017-12-31&page=1
            // identify how many pages on order page
            // current_page
            // last_page
            // per_page
            // data (System.Collections.ObjectModel)

            int current_page = jsonD.current_page;
            int last_page = jsonD.last_page;
            int per_page = jsonD.per_page;

            // process orders on page
            var orders = jsonD.data;
            foreach (var order in orders)
            {
                // order_name
                // processed_at
                // total_price
                // shipping_name
                // shipping_zip
                // shipping_city
                // shipping_address1
                // orderitems
                var orderName = order.order_name;
                var processedAt = order.processed_at;
                var totalPrice = order.total_price;
                var financialStatus = order.financial_status;
                var fulfillmentStatus = order.fulfillment_status;
                var shippingName = order.shipping_name;
                var shippingZip = order.shipping_zip;
                var shippingCity = order.shipping_city;
                var shippingAddress1 = order.shipping_address1;
                var shippingAddress2 = order.shipping_address2;
                var orderNote = order.local_note;

                var orderitems = order.orderitems;
                foreach (var orderitem in orderitems)
                {
                    var aliOrderNumber = orderitem.ali_order_no;
                    var SKU = orderitem.sku;
                    var supplier = orderitem.supplier_name;
                    var productName = orderitem.title;
                    var variant = orderitem.variant_title;
                    var cost = orderitem.cost;
                    var quantity = orderitem.quantity;
                    var price = orderitem.price;

                    string trackingNumber = "";
                    foreach (var fulfillment in orderitem.fulfillments)
                    {
                        if (trackingNumber.Equals(""))
                        {
                            trackingNumber = fulfillment.tracking_number;
                        }
                        else
                        {
                            trackingNumber += ", " + fulfillment.tracking_number;
                        }
                    }

                    var oberloOrder = new OberloOrder();
                    oberloOrder.OrderNumber = orderName;
                    oberloOrder.CreatedDate = processedAt;
                    oberloOrder.FinancialStatus = financialStatus;
                    oberloOrder.FulfillmentStatus = fulfillmentStatus;
                    oberloOrder.Supplier = supplier;
                    oberloOrder.SKU = SKU;
                    oberloOrder.ProductName = productName;
                    oberloOrder.Variant = variant;
                    oberloOrder.Quantity = quantity;
                    oberloOrder.ProductPrice = price;
                    oberloOrder.TrackingNumber = trackingNumber;
                    oberloOrder.AliOrderNumber = aliOrderNumber;
                    oberloOrder.CustomerName = shippingName;
                    oberloOrder.CustomerAddress = shippingAddress1;
                    oberloOrder.CustomerAddress2 = shippingAddress2;
                    oberloOrder.CustomerCity = shippingCity;
                    oberloOrder.CustomerZip = shippingZip;
                    oberloOrder.OrderNote = orderNote;
                    //oberloOrder.OrderState = orderState;
                    oberloOrder.TotalPrice = totalPrice;
                    oberloOrder.Cost = cost;

                    oberloOrders.Add(oberloOrder);
                }
            }

            // and process the rest of the pages
            for (int i = current_page + 1; i <= last_page; i++)
            {
                ScrapeOberloOrders(driver, oberloOrders, from, to, i);
            }

            driver.Close();
            return oberloOrders;
        }

        static void ScrapeOberloOrders(IWebDriver driver, List<OberloOrder> oberloOrders, DateTime from, DateTime to, int page)
        {
            // https://app.oberlo.com/orders?from=2017-01-01&to=2017-12-31&page=1
            string url = string.Format("https://app.oberlo.com/orders?from={0:yyyy-MM-dd}&to={1:yyyy-MM-dd}&page={2}", from, to, page);
            driver.Navigate().GoToUrl(url);

            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            var ready = wait.Until(d => ((IJavaScriptExecutor)d).ExecuteScript("return document.readyState").Equals("complete"));

            IJavaScriptExecutor js = driver as IJavaScriptExecutor;
            var json = js.ExecuteScript("return window.App.payload.orders;");

            // convert to json dynamic object
            string jsonString = JsonConvert.SerializeObject(json);
            dynamic jsonD = JsonConvert.DeserializeObject(jsonString);

            // https://app.oberlo.com/orders?from=2017-01-01&to=2017-12-31&page=1
            // identify how many pages on order page
            // current_page
            // last_page
            // per_page
            // data (System.Collections.ObjectModel)

            int current_page = jsonD.current_page;
            int last_page = jsonD.last_page;
            int per_page = jsonD.per_page;

            // process orders on page
            var orders = jsonD.data;
            foreach (var order in orders)
            {
                // order_name
                // processed_at
                // total_price
                // shipping_name
                // shipping_zip
                // shipping_city
                // shipping_address1
                // orderitems
                var orderName = order.order_name;
                var processedAt = order.processed_at;
                var totalPrice = order.total_price;
                var financialStatus = order.financial_status;
                var fulfillmentStatus = order.fullfullment_status;
                var shippingName = order.shipping_name;
                var shippingZip = order.shipping_zip;
                var shippingCity = order.shipping_city;
                var shippingAddress1 = order.shipping_address1;
                var shippingAddress2 = order.shipping_address2;
                var orderNote = order.local_note;

                var orderitems = order.orderitems;
                foreach (var orderitem in orderitems)
                {
                    var aliOrderNumber = orderitem.ali_order_no;
                    var SKU = orderitem.sku;
                    var supplier = orderitem.supplier_name;
                    var productName = orderitem.title;
                    var variant = orderitem.variant_title;
                    var cost = orderitem.cost;
                    var quantity = orderitem.quantity;

                    string trackingNumber = "";
                    foreach (var fulfillment in orderitem.fulfillments)
                    {
                        if (trackingNumber.Equals(""))
                        {
                            trackingNumber = fulfillment.tracking_number;
                        }
                        else
                        {
                            trackingNumber += ", " + fulfillment.tracking_number;
                        }
                    }

                    var oberloOrder = new OberloOrder();
                    oberloOrder.OrderNumber = orderName;
                    oberloOrder.CreatedDate = processedAt;
                    oberloOrder.FinancialStatus = financialStatus;
                    oberloOrder.FulfillmentStatus = fulfillmentStatus;
                    oberloOrder.Supplier = supplier;
                    oberloOrder.SKU = SKU;
                    oberloOrder.ProductName = productName;
                    oberloOrder.Variant = variant;
                    oberloOrder.Quantity = quantity;
                    oberloOrder.TrackingNumber = trackingNumber;
                    oberloOrder.AliOrderNumber = aliOrderNumber;
                    oberloOrder.CustomerName = shippingName;
                    oberloOrder.CustomerAddress = shippingAddress1;
                    oberloOrder.CustomerAddress2 = shippingAddress2;
                    oberloOrder.CustomerCity = shippingCity;
                    oberloOrder.CustomerZip = shippingZip;
                    oberloOrder.OrderNote = orderNote;
                    //oberloOrder.OrderState = orderState;
                    oberloOrder.Cost = cost;

                    oberloOrders.Add(oberloOrder);
                }
            }
        }
    }
}
