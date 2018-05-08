using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace AccountingServices
{
    public class Shopify
    {
        public static int CountShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, string querySuffix)
        {
            // GET /admin/orders/count.json
            // Retrieve a count of all the orders

            string url = String.Format("https://{0}/admin/orders/count.json?{1}", shopifyDomain, querySuffix);

            using (var client = new WebClient())
            {
                // make sure we read in utf8
                client.Encoding = System.Text.Encoding.UTF8;

                // have to use the header field since normal GET url doesn't work, i.e.
                // string url = String.Format("https://{0}:{1}@{2}/admin/orders.json", shopifyAPIKey, shopifyAPIPassword, shopifyDomain);
                // https://stackoverflow.com/questions/28177871/shopify-and-private-applications
                client.Headers.Add("X-Shopify-Access-Token", shopifyAPIPassword);
                string json = client.DownloadString(url);

                // parse json
                dynamic jsonDe = JsonConvert.DeserializeObject(json);

                return jsonDe.count;
            }
        }

        public static void ReadShopifyOrdersPage(List<ShopifyOrder> shopifyOrders, string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, int limit, int page, string querySuffix)
        {
            // GET /admin/orders.json?limit=250&page=1
            // Retrieve a list of Orders(OPEN Orders by default, use status=any for ALL orders)

            // GET /admin/orders/#{id}.json
            // Receive a single Order

            // parameters:
            // financial_status=paid
            // status=any

            // By default that Orders API endpoint can give you a maximum of 50 orders. 
            // You can increase the limit to 250 orders by adding &limit=250 to the URL. 
            // If your query has more than 250 results then you can page through them 
            // by using the page URL parameter: https://help.shopify.com/api/reference/order
            // limit: Amount of results (default: 50)(maximum: 250)
            // page: Page to show, (default: 1)

            string url = String.Format("https://{0}/admin/orders.json?limit={1}&page={2}&{3}", shopifyDomain, limit, page, querySuffix);

            using (var client = new WebClient())
            {
                // make sure we read in utf8
                client.Encoding = System.Text.Encoding.UTF8;

                // have to use the header field since normal GET url doesn't work, i.e.
                // string url = String.Format("https://{0}:{1}@{2}/admin/orders.json", shopifyAPIKey, shopifyAPIPassword, shopifyDomain);
                // https://stackoverflow.com/questions/28177871/shopify-and-private-applications
                client.Headers.Add("X-Shopify-Access-Token", shopifyAPIPassword);
                string json = client.DownloadString(url);

                // parse json
                dynamic jsonDe = JsonConvert.DeserializeObject(json);

                foreach (var order in jsonDe.orders)
                {
                    var shopifyOrder = new ShopifyOrder();

                    shopifyOrder.Id = order.id;
                    shopifyOrder.CreatedAt = order.created_at;
                    shopifyOrder.ProcessedAt = order.processed_at;
                    shopifyOrder.UpdatedAt = order.updated_at;
                    shopifyOrder.Name = order.name;
                    shopifyOrder.FinancialStatus = order.financial_status;
                    string fulfillmentStatusTmp = order.fulfillment_status;
                    fulfillmentStatusTmp = (fulfillmentStatusTmp == null ? "unfulfilled" : fulfillmentStatusTmp);
                    shopifyOrder.FulfillmentStatus = fulfillmentStatusTmp;

                    shopifyOrder.Gateway = order.gateway;
                    if (null != order.payment_details)
                    {
                        shopifyOrder.PaymentId = string.Format("{0} {1}", order.payment_details.credit_card_company, order.payment_details.credit_card_bin);
                    }

                    shopifyOrder.TotalPrice = order.total_price;
                    shopifyOrder.TotalTax = order.total_tax;
                    shopifyOrder.CustomerEmail = order.contact_email;

                    shopifyOrder.CustomerName = string.Format("{0} {1}", order.customer.first_name, order.customer.last_name);
                    shopifyOrder.CustomerAddress = order.customer.default_address.address1;
                    shopifyOrder.CustomerAddress2 = order.customer.default_address.address2;
                    shopifyOrder.CustomerCity = order.customer.default_address.city;
                    shopifyOrder.CustomerZipCode = order.customer.default_address.zip;

                    // check if cancelled_at exists (meaning the order has been cancelled)
                    var cancelledAt = order.cancelled_at;
                    if (cancelledAt != null && cancelledAt.Type != JTokenType.Null)
                    {
                        shopifyOrder.CancelledAt = order.cancelled_at;
                    }

                    // also add note
                    shopifyOrder.Note = order.note;

                    if (shopifyOrder.Name.Equals("#1103"))
                    {
                        // breakpoint here
                    }
                    if (shopifyOrder.CustomerEmail.Equals("janne.braseth@gmail.com"))
                    {
                        // breakpoint here
                    }

                    if (order.refunds != null)
                    {
                        decimal refundSubTotal = 0;
                        decimal refundTotalTax = 0;

                        // calculate refund
                        foreach (var refund in order.refunds)
                        {
                            var refundItems = refund.refund_line_items;
                            foreach (var refundItem in refundItems)
                            {
                                refundSubTotal += (decimal) refundItem.subtotal;
                                refundTotalTax += (decimal) refundItem.total_tax;
                            }

                            var orderAdjustments = refund.order_adjustments;
                            foreach (var orderAdjustment in orderAdjustments)
                            {
                                refundSubTotal += -((decimal) orderAdjustment.amount);
                                refundTotalTax += -((decimal) orderAdjustment.tax_amount);
                            }
                        }

                        // perform refund
                        shopifyOrder.TotalPrice -= refundSubTotal;
                        shopifyOrder.TotalTax -= refundTotalTax;
                    }

                    shopifyOrders.Add(shopifyOrder);
                }
            }
        }

        public static List<ShopifyOrder> ReadShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, int totalShopifyOrders, string querySuffix)
        {
            // https://ecommerce.shopify.com/c/shopify-apis-and-technology/t/paginate-api-results-113066
            // Use the /admin/products/count.json to get the count of all products. 
            // Then divide that number by 250 to get the total amount of pages.

            // the web api requires a pagination to read in all orders
            // max orders per page is 250

            var shopifyOrders = new List<ShopifyOrder>();

            int limit = 250;
            if (totalShopifyOrders > 0)
            {
                int numPages = (int)Math.Ceiling((double)totalShopifyOrders / limit);
                for (int i = 1; i <= numPages; i++)
                {
                    ReadShopifyOrdersPage(shopifyOrders, shopifyDomain, shopifyAPIKey, shopifyAPIPassword, limit, i, querySuffix);
                }
            }

            return shopifyOrders;
        }

        public static List<ShopifyOrder> ReadShopifyOrders(string shopifyDomain, string shopifyAPIKey, string shopifyAPIPassword, string querySuffix = "status=any")
        {
            int totalShopifyOrders = CountShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword, querySuffix);
            return ReadShopifyOrders(shopifyDomain, shopifyAPIKey, shopifyAPIPassword, totalShopifyOrders, querySuffix);
        }
    }
}
