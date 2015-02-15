using JdSdk;
using JdSdk.Domain.Order;
using JdSdk.Request.Order;
using JdSdk.Response.Order;
using Quartz;
using Quartz.Impl;
using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Transactions;
using System.Web.Configuration;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace Finance.Models
{
    /// <summary>
    /// 调用京东API的Job
    /// </summary>
    public class JDJob : IJob
    {
        //连接字符串
        public static string connectionString = WebConfigurationManager.ConnectionStrings["JDConnectionString"].ToString();

        //京东

        public static string accessToken = ConfigurationManager.AppSettings["accessToken"].ToString();

        public static string appSecret = ConfigurationManager.AppSettings["appSecret"].ToString();

        public static string appKey = ConfigurationManager.AppSettings["appKey"].ToString();

        public static string serverUrl = ConfigurationManager.AppSettings["serverUrl"].ToString();

        public static int pageSize = int.Parse(ConfigurationManager.AppSettings["pageSize"].ToString());

        public void Execute(IJobExecutionContext context)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = "logServer.txt";
            string savePath = Path.Combine(path, fileName);
            if (!System.IO.File.Exists(savePath))
            {
                System.IO.FileStream fsnew = System.IO.File.Create(savePath);
                fsnew.Close();
            }
            using (System.IO.FileStream fs = System.IO.File.Open(savePath, System.IO.FileMode.Append, System.IO.FileAccess.Write))
            {
                using (System.IO.StreamWriter w = new System.IO.StreamWriter(fs))
                {
                    w.WriteLine("---------------------------------------------------------------------------------------------");
                    w.WriteLine("执行时间:{0}", DateTime.Now.ToString());
                    w.Flush();
                    w.Close();

                }
            }
            //调用京东api
            OrderSearchRequest orderSearchRequest = new OrderSearchRequest();
            orderSearchRequest.StartDate = DateTime.Now.AddHours(-48).ToString("yyyy-MM-dd hh:mm:ss");
            orderSearchRequest.EndDate = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
            orderSearchRequest.OrderState = "WAIT_SELLER_STOCK_OUT";
            orderSearchRequest.Page = "1";
            orderSearchRequest.PageSize = pageSize.ToString();
            orderSearchRequest.OptionalFields = "delivery_type,logistics_id,order_state_remark,order_state,order_payment,order_remark,order_id,consignee_info,pay_type,item_info_list,order_source,balance_used,order_total_price,payment_confirm_time,coupon_detail_list,invoice_info,waybill,parent_order_id,freight_price,store_order,modified,order_start_time,pin,return_order,seller_discount,order_seller_price,vender_id,vender_remark,order_type,vat_invoice_info";
            //orderSearchRequest.SortType = "";
            //orderSearchRequest.DateType = "";
            DefaultJdClient client = new DefaultJdClient(serverUrl, appKey, appSecret);
            OrderSearchResponse response = client.Execute(orderSearchRequest, accessToken);
            if (!response.IsError)
            {
                //引用事务机制，出错时，事物回滚
                using (TransactionScope transaction = new TransactionScope())
                {
                    using (SqlConnection connection = new SqlConnection(connectionString))
                    {
                        connection.Open();
                        //订单信息
                        DataTable orderInfoDT = new DataTable();
                        orderInfoDT.Columns.Add("order_id", typeof(string));
                        orderInfoDT.Columns.Add("order_source", typeof(string));
                        orderInfoDT.Columns.Add("vender_id", typeof(string));
                        orderInfoDT.Columns.Add("pay_type", typeof(string));
                        orderInfoDT.Columns.Add("order_total_price", typeof(decimal));
                        orderInfoDT.Columns.Add("order_seller_price", typeof(decimal));
                        orderInfoDT.Columns.Add("order_payment", typeof(decimal));
                        orderInfoDT.Columns.Add("freight_price", typeof(decimal));
                        orderInfoDT.Columns.Add("seller_discount", typeof(decimal));
                        orderInfoDT.Columns.Add("order_state", typeof(string));
                        orderInfoDT.Columns.Add("order_state_remark", typeof(string));
                        orderInfoDT.Columns.Add("delivery_type", typeof(string));
                        orderInfoDT.Columns.Add("invoice_info", typeof(string));
                        orderInfoDT.Columns.Add("order_remark", typeof(string));
                        orderInfoDT.Columns.Add("order_start_time", typeof(DateTime));
                        orderInfoDT.Columns.Add("order_end_time", typeof(DateTime));
                        orderInfoDT.Columns.Add("modified", typeof(DateTime));
                        orderInfoDT.Columns.Add("consignee_info", typeof(string));
                        orderInfoDT.Columns.Add("vender_remark", typeof(string));
                        orderInfoDT.Columns.Add("balance_used", typeof(string));
                        orderInfoDT.Columns.Add("payment_confirm_time", typeof(DateTime));
                        orderInfoDT.Columns.Add("waybill", typeof(string));
                        orderInfoDT.Columns.Add("logistics_id", typeof(string));
                        orderInfoDT.Columns.Add("taxpayer_ident", typeof(string));
                        orderInfoDT.Columns.Add("registered_address", typeof(string));
                        orderInfoDT.Columns.Add("registered_phone", typeof(string));
                        orderInfoDT.Columns.Add("deposit_bank", typeof(string));
                        orderInfoDT.Columns.Add("bank_account", typeof(string));
                        orderInfoDT.Columns.Add("parent_order_id", typeof(string));
                        orderInfoDT.Columns.Add("pin", typeof(string));
                        orderInfoDT.Columns.Add("return_order", typeof(string));
                        orderInfoDT.Columns.Add("order_type", typeof(string));
                        orderInfoDT.Columns.Add("store_order", typeof(string));
                        orderInfoDT.Columns.Add("bInvoiceMade", typeof(bool));
                        //订单行项目信息
                        DataTable itemInfoDT = new DataTable();
                        itemInfoDT.Columns.Add("order_id", typeof(string));
                        itemInfoDT.Columns.Add("sku_id", typeof(string));
                        itemInfoDT.Columns.Add("outer_sku_id", typeof(string));
                        itemInfoDT.Columns.Add("sku_name", typeof(string));
                        itemInfoDT.Columns.Add("jd_price", typeof(decimal));
                        itemInfoDT.Columns.Add("gift_point", typeof(string));
                        itemInfoDT.Columns.Add("ware_id", typeof(string));
                        itemInfoDT.Columns.Add("item_total", typeof(int));
                        itemInfoDT.Columns.Add("product_no", typeof(string));
                        //优惠详细信息
                        DataTable couponDetailDT = new DataTable();
                        couponDetailDT.Columns.Add("order_id", typeof(string));
                        couponDetailDT.Columns.Add("sku_id", typeof(string));
                        couponDetailDT.Columns.Add("coupon_type", typeof(string));
                        couponDetailDT.Columns.Add("coupon_price", typeof(decimal));


                        foreach (var orderInfo in response.OrderResult.OrderInfoList)
                        {

                            string sql = string.Format("select * from order_info where order_id={0}", orderInfo.OrderId);
                            SqlCommand cmd = new SqlCommand(sql, connection);
                            var data = cmd.ExecuteScalar();
                            if (data == null)
                            {
                                AddDataToDataTable(orderInfo, orderInfoDT, itemInfoDT, couponDetailDT);
                            }
                        }
                        if (pageSize >= response.OrderResult.OrderTotal)
                        {
                        }
                        else
                        {
                            //循环调用京东api
                            for (int page = 2; (page - 1) * pageSize < response.OrderResult.OrderTotal; page++)
                            {
                                orderSearchRequest.StartDate = DateTime.Now.AddHours(-48).ToString("yyyy-MM-dd hh:mm:ss");
                                orderSearchRequest.EndDate = DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss");
                                orderSearchRequest.OrderState = "WAIT_SELLER_STOCK_OUT";
                                orderSearchRequest.Page = page.ToString();
                                orderSearchRequest.PageSize = pageSize.ToString();
                                orderSearchRequest.OptionalFields = "delivery_type,logistics_id,order_state_remark,order_state,order_payment,order_remark,order_id,consignee_info,pay_type,item_info_list,order_source,balance_used,order_total_price,payment_confirm_time,coupon_detail_list,invoice_info,waybill,parent_order_id,freight_price,store_order,modified,order_start_time,pin,return_order,seller_discount,order_seller_price,vender_id,vender_remark,order_type,vat_invoice_info";
                                //orderSearchRequest.SortType = "";
                                //orderSearchRequest.DateType = "";
                                OrderSearchResponse responseNew = client.Execute(orderSearchRequest, accessToken);
                                if (!response.IsError)
                                {
                                    foreach (var orderInfo in responseNew.OrderResult.OrderInfoList)
                                    {
                                        string sql = string.Format("select * from order_info where order_id={0}", orderInfo.OrderId);
                                        SqlCommand cmd = new SqlCommand(sql, connection);
                                        var data = cmd.ExecuteScalar();
                                        if (data == null)
                                        {
                                            AddDataToDataTable(orderInfo, orderInfoDT, itemInfoDT, couponDetailDT);
                                        }
                                    }
                                }
                                else
                                {
                                    //记录日志
                                    WriteLog(response);
                                }
                            }
                        }
                        using (SqlBulkCopy orderInfoSqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction))
                        {
                            orderInfoSqlbulkcopy.DestinationTableName = "order_info";//数据库中的表名
                            foreach (System.Data.DataColumn k in orderInfoDT.Columns)
                            {
                                orderInfoSqlbulkcopy.ColumnMappings.Add(k.ColumnName.ToString(), k.ColumnName.ToString());
                            }
                            orderInfoSqlbulkcopy.WriteToServer(orderInfoDT);
                        }
                        using (SqlBulkCopy itemInfoSqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction))
                        {
                            itemInfoSqlbulkcopy.DestinationTableName = "item_info";//数据库中的表名
                            foreach (System.Data.DataColumn k in itemInfoDT.Columns)
                            {
                                itemInfoSqlbulkcopy.ColumnMappings.Add(k.ColumnName.ToString(), k.ColumnName.ToString());
                            }
                            itemInfoSqlbulkcopy.WriteToServer(itemInfoDT);
                        }
                        using (SqlBulkCopy couponDetailSqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction))
                        {
                            couponDetailSqlbulkcopy.DestinationTableName = "coupon_detail";//数据库中的表名
                            foreach (System.Data.DataColumn k in couponDetailDT.Columns)
                            {
                                couponDetailSqlbulkcopy.ColumnMappings.Add(k.ColumnName.ToString(), k.ColumnName.ToString());
                            }
                            couponDetailSqlbulkcopy.WriteToServer(couponDetailDT);
                        }
                        transaction.Complete();
                    }
                }
                //记录日志
                WriteLog(response);
            }
            else
            {
                //记录日志
                WriteLog(response);
            }


        }

        /// <summary>
        /// 订单数据填充到DataTable
        /// </summary>
        /// <param name="orderInfo">订单信息</param>
        /// <param name="orderInfoDT">订单DataTable</param>
        /// <param name="itemInfoDT">订单行项目DataTable</param>
        /// <param name="couponDetailDT">优惠详细信息DataTable</param>
        public void AddDataToDataTable(OrderSearchInfo orderInfo, DataTable orderInfoDT, DataTable itemInfoDT, DataTable couponDetailDT)
        {
            //订单
            DataRow newOrderInfoRow = orderInfoDT.NewRow();
            newOrderInfoRow["order_id"] = orderInfo.OrderId;
            newOrderInfoRow["order_source"] = "京东";
            newOrderInfoRow["vender_id"] = orderInfo.VenderId;
            newOrderInfoRow["pay_type"] = orderInfo.PayType;
            newOrderInfoRow["order_total_price"] = string.IsNullOrEmpty(orderInfo.OrderTotalPrice) ? 0 : decimal.Parse(orderInfo.OrderTotalPrice);
            newOrderInfoRow["order_seller_price"] = string.IsNullOrEmpty(orderInfo.OrderSellerPrice) ? 0 : decimal.Parse(orderInfo.OrderSellerPrice);
            newOrderInfoRow["order_payment"] = string.IsNullOrEmpty(orderInfo.OrderPayment) ? 0 : decimal.Parse(orderInfo.OrderPayment);
            newOrderInfoRow["freight_price"] = string.IsNullOrEmpty(orderInfo.FreightPrice) ? 0 : decimal.Parse(orderInfo.FreightPrice);
            newOrderInfoRow["seller_discount"] = string.IsNullOrEmpty(orderInfo.SellerDiscount) ? 0 : decimal.Parse(orderInfo.SellerDiscount);
            newOrderInfoRow["order_state"] = orderInfo.OrderState;
            newOrderInfoRow["order_state_remark"] = orderInfo.OrderStateRemark;
            newOrderInfoRow["delivery_type"] = orderInfo.DeliveryType;
            newOrderInfoRow["invoice_info"] = orderInfo.InvoiceInfo;
            newOrderInfoRow["order_remark"] = orderInfo.OrderRemark;
            newOrderInfoRow["order_start_time"] = string.IsNullOrEmpty(orderInfo.OrderStartTime) ? DateTime.MaxValue.AddMilliseconds(-10) : DateTime.Parse(orderInfo.OrderStartTime);
            newOrderInfoRow["order_end_time"] = string.IsNullOrEmpty(orderInfo.OrderEndTime) ? DateTime.MaxValue.AddMilliseconds(-10) : DateTime.Parse(orderInfo.OrderEndTime);
            newOrderInfoRow["modified"] = string.IsNullOrEmpty(orderInfo.Modified) ? DateTime.MaxValue.AddMilliseconds(-10) : DateTime.Parse(orderInfo.Modified);
            newOrderInfoRow["consignee_info"] = orderInfo.ConsigneeInfo;
            newOrderInfoRow["vender_remark"] = orderInfo.VenderRemark;
            newOrderInfoRow["balance_used"] = orderInfo.BalanceUsed;
            newOrderInfoRow["payment_confirm_time"] = string.IsNullOrEmpty(orderInfo.PaymentConfirmTime) ? DateTime.MaxValue.AddMilliseconds(-10) : DateTime.Parse(orderInfo.PaymentConfirmTime);
            newOrderInfoRow["waybill"] = orderInfo.Waybill;
            newOrderInfoRow["logistics_id"] = orderInfo.LogisticsId;
            newOrderInfoRow["taxpayer_ident"] = orderInfo.VatInvoiceInfo == null ? string.Empty : orderInfo.VatInvoiceInfo.TaxpayerIdent;
            newOrderInfoRow["registered_address"] = orderInfo.VatInvoiceInfo == null ? string.Empty : orderInfo.VatInvoiceInfo.RegisteredAddress;
            newOrderInfoRow["registered_phone"] = orderInfo.VatInvoiceInfo == null ? string.Empty : orderInfo.VatInvoiceInfo.RegisteredPhone;
            newOrderInfoRow["deposit_bank"] = orderInfo.VatInvoiceInfo == null ? string.Empty : orderInfo.VatInvoiceInfo.DepositBank;
            newOrderInfoRow["bank_account"] = orderInfo.VatInvoiceInfo == null ? string.Empty : orderInfo.VatInvoiceInfo.BankAccount;
            newOrderInfoRow["parent_order_id"] = orderInfo.ParentOrderId;
            newOrderInfoRow["pin"] = orderInfo.Pin;
            newOrderInfoRow["return_order"] = orderInfo.ReturnOrder;
            newOrderInfoRow["order_type"] = string.Empty;
            newOrderInfoRow["store_order"] = string.Empty;
            newOrderInfoRow["bInvoiceMade"] = 0;
            orderInfoDT.Rows.Add(newOrderInfoRow);
            //订单行项目
            foreach (var itemInfo in orderInfo.ItemInfoList)
            {
                DataRow newItemInfoRow = itemInfoDT.NewRow();
                newItemInfoRow["order_id"] = orderInfo.OrderId;
                newItemInfoRow["sku_id"] = itemInfo.SkuId;
                newItemInfoRow["outer_sku_id"] = itemInfo.OuterSkuId;
                newItemInfoRow["sku_name"] = itemInfo.SkuName;
                newItemInfoRow["jd_price"] = string.IsNullOrEmpty(itemInfo.JdPrice) ? 0 : decimal.Parse(itemInfo.JdPrice);
                newItemInfoRow["gift_point"] = itemInfo.GiftPoint;
                newItemInfoRow["ware_id"] = itemInfo.WareId;
                newItemInfoRow["item_total"] = string.IsNullOrEmpty(itemInfo.ItemTotal) ? 0 : int.Parse(itemInfo.ItemTotal);
                newItemInfoRow["product_no"] = itemInfo.ProductNo;
                itemInfoDT.Rows.Add(newItemInfoRow);
            }
            //优惠详细信息
            foreach (var couponDetail in orderInfo.CouponDetailList)
            {
                DataRow newCouponDetailRow = couponDetailDT.NewRow();
                newCouponDetailRow["order_id"] = couponDetail.OrderId;
                newCouponDetailRow["sku_id"] = couponDetail.SkuId;
                newCouponDetailRow["coupon_type"] = couponDetail.CouponType;
                newCouponDetailRow["coupon_price"] = string.IsNullOrEmpty(couponDetail.CouponPrice) ? 0 : decimal.Parse(couponDetail.CouponPrice);
                couponDetailDT.Rows.Add(newCouponDetailRow);
            }
        }

        /// <summary>
        /// 记录日志
        /// </summary>
        /// <param name="response"></param>
        public void WriteLog(OrderSearchResponse response)
        {
            //记录日志
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string fileName = "logServer.txt";
            string savePath = Path.Combine(path, fileName);
            if (!System.IO.File.Exists(savePath))
            {
                System.IO.FileStream fsnew = System.IO.File.Create(savePath);
                fsnew.Close();
            }
            using (System.IO.FileStream fs = System.IO.File.Open(savePath, System.IO.FileMode.Append, System.IO.FileAccess.Write))
            {
                using (System.IO.StreamWriter w = new System.IO.StreamWriter(fs))
                {
                    if (response.IsError)
                    {
                        w.WriteLine("错误代码:{0}", response.ErrCode);
                        w.WriteLine("错误信息:{0}", response.ErrMsg);
                        w.WriteLine("参考信息:{0}", response.ZhErrMsg);
                        w.WriteLine("---------------------------------------------------------------------------------------------");
                        w.Flush();
                        w.Close();
                    }
                    else
                    {
                        w.WriteLine("完成更新:共更新{0}条数据", response.OrderResult.OrderTotal);
                        w.WriteLine("---------------------------------------------------------------------------------------------");
                        w.Flush();
                        w.Close();
                    }
                }
            }
        }
    }

    public class JobScheduler
    {
        public static void Start()
        {
            //调度器
            IScheduler scheduler = StdSchedulerFactory.GetDefaultScheduler();

            //job详情
            IJobDetail job = JobBuilder.Create<JDJob>()
            .WithIdentity("job1", "group1")//定义name/group
            .Build();

            /*
            startTimeOfDay 每天开始时间
            endTimeOfDay 每天结束时间
            daysOfWeek 需要执行的星期
            interval 执行间隔
            intervalUnit 执行间隔的单位（秒，分钟，小时，天，月，年，星期）
            repeatCount 重复次数
            */

            //触发器
            ITrigger trigger = TriggerBuilder.Create()
                .WithIdentity("trigger1", "group1")//定义name/group
                .WithDailyTimeIntervalSchedule
                  (s =>
                     s.WithIntervalInHours(12) //每5小时执行一次
                    .OnEveryDay()
                         //.StartingDailyAt(TimeOfDay.HourAndMinuteOfDay(DateTime.Now.Hour,DateTime.Now.Minute))
                    .StartingDailyAt(TimeOfDay.HourAndMinuteOfDay(0, 0))  //每天0:00开始
                  )
                .Build();

            //关联job和触发器  
            scheduler.ScheduleJob(job, trigger);

            //执行  
            scheduler.Start();
        }
    }
}