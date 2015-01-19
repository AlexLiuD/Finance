using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Transactions;
using System.Web;
using System.Web.Configuration;
using System.Web.Mvc;
using System.Web.Script.Serialization;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace Finance.Controllers
{
    public class HomeController : Controller
    {
        string connectionString = WebConfigurationManager.ConnectionStrings["JDConnectionString"].ToString();

        #region 导入导出
        public ActionResult Index()
        {
            //ViewBag.Success = "导入成功.";
            //ViewBag.Error = "导入失败.";
            return View();
        }

        //[HttpPost]
        //public ActionResult Index(HttpPostedFileBase filebase)
        //{
        //    HttpPostedFileBase file = Request.Files["files"];
        //    string FileName;
        //    string savePath;
        //    if (file == null || file.ContentLength <= 0)
        //    {
        //        ViewBag.error = "文件不能为空";
        //        return View();
        //    }
        //    else
        //    {
        //        string filename = Path.GetFileName(file.FileName);
        //        int filesize = file.ContentLength;//获取上传文件的大小单位为字节byte
        //        string fileEx = System.IO.Path.GetExtension(filename);//获取上传文件的扩展名
        //        string NoFileName = System.IO.Path.GetFileNameWithoutExtension(filename);//获取无扩展名的文件名
        //        string FileType = ".xls,.xlsx";//定义上传文件的类型字符串

        //        FileName = NoFileName + DateTime.Now.ToString("yyyyMMddhhmmss") + fileEx;
        //        if (!FileType.Contains(fileEx))
        //        {
        //            ViewBag.error = "文件类型不对，只能导入xls和xlsx格式的文件";
        //            return View();
        //        }
        //        string path = AppDomain.CurrentDomain.BaseDirectory + "uploads/";
        //        savePath = Path.Combine(path, FileName);
        //        file.SaveAs(savePath);
        //    }

        //    //string result = string.Empty;
        //    /*
        //    //2003（Microsoft.Jet.Oledb.4.0）
        //    string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
        //    //2010（Microsoft.ACE.OLEDB.12.0）
        //    string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
        //    */
        //    string strConn;
        //    strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + savePath + ";" + "Extended Properties=Excel 8.0";
        //    OleDbConnection conn = new OleDbConnection(strConn);
        //    conn.Open();
        //    OleDbDataAdapter myCommand = new OleDbDataAdapter("select * from [Sheet1$]", strConn);
        //    DataSet myDataSet = new DataSet();
        //    try
        //    {
        //        myCommand.Fill(myDataSet, "ExcelInfo");
        //    }
        //    catch (Exception ex)
        //    {
        //        ViewBag.error = ex.Message;
        //        return View();
        //    }
        //    DataTable table = myDataSet.Tables["ExcelInfo"].DefaultView.ToTable();

        //    //引用事务机制，出错时，事物回滚
        //    using (TransactionScope transaction = new TransactionScope())
        //    {
        //        SqlConnection connection = new SqlConnection(connectionString);
        //        string sql = "delete from Saleorders";
        //        connection.Open();
        //        SqlCommand cmd = new SqlCommand(sql, connection);
        //        cmd.ExecuteNonQuery(); 
        //        for (int i = 0; i < table.Rows.Count; i++)
        //        {
        //            string insertSql = "insert into Saleorders(OrderNum,GoodsId,GoodsName,OrderQuantity,Payment,OrderTime,JdPrice,OrderAmount,SettlementAmount,BalancePayment,NeedAmount,OrderStatus,OrderType,SingleAccount,CustomerName,CustomerAddress,TelPhone,OrderRemarks,InvoiceType,InvoicesHead,InvoicesContent,MerchantRemark,MerchantRemarkRate,FreightAmount,PaymentConfirmTime,VATInvoice,TaxpayerIdentificationNum,BankAccount,Bank,RegistrationPhone,RegisteredAddress) values(@OrderNum,@GoodsId,@GoodsName,@OrderQuantity,@Payment,@OrderTime,@JdPrice,@OrderAmount,@SettlementAmount,@BalancePayment,@NeedAmount,@OrderStatus,@OrderType,@SingleAccount,@CustomerName,@CustomerAddress,@TelPhone,@OrderRemarks,@InvoiceType,@InvoicesHead,@InvoicesContent,@MerchantRemark,@MerchantRemarkRate,@FreightAmount,@PaymentConfirmTime,@VATInvoice,@TaxpayerIdentificationNum,@BankAccount,@Bank,@RegistrationPhone,@RegisteredAddress)";
        //            SqlCommand command = new SqlCommand(insertSql, connection);
        //            //command.CommandTimeout = 360;
        //            command.Parameters.Add("@OrderNum",SqlDbType.NVarChar).Value=table.Rows[i][0].ToString().Trim();
        //            command.Parameters.Add("@GoodsId",SqlDbType.NVarChar).Value=table.Rows[i][1].ToString().Trim();
        //            command.Parameters.Add("@GoodsName",SqlDbType.NVarChar).Value=table.Rows[i][2].ToString().Trim();
        //            command.Parameters.Add("@OrderQuantity",SqlDbType.Int).Value=string.IsNullOrEmpty(table.Rows[i][3].ToString().Trim())?0:int.Parse(table.Rows[i][3].ToString().Trim());
        //            command.Parameters.Add("@Payment",SqlDbType.NVarChar).Value=table.Rows[i][4].ToString().Trim();
        //            command.Parameters.Add("@OrderTime", SqlDbType.DateTime).Value = string.IsNullOrEmpty(table.Rows[i][5].ToString().Trim()) ? DateTime.MaxValue : DateTime.Parse(table.Rows[i][5].ToString().Trim());
        //            command.Parameters.Add("@JdPrice", SqlDbType.Decimal).Value = string.IsNullOrEmpty(table.Rows[i][6].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][6].ToString().Trim());
        //            command.Parameters.Add("@OrderAmount", SqlDbType.Decimal).Value = string.IsNullOrEmpty(table.Rows[i][7].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][7].ToString().Trim());
        //            command.Parameters.Add("@SettlementAmount", SqlDbType.Decimal).Value = string.IsNullOrEmpty(table.Rows[i][8].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][8].ToString().Trim());
        //            command.Parameters.Add("@BalancePayment", SqlDbType.Decimal).Value = string.IsNullOrEmpty(table.Rows[i][9].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][9].ToString().Trim());
        //            command.Parameters.Add("@NeedAmount", SqlDbType.Decimal).Value = string.IsNullOrEmpty(table.Rows[i][10].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][10].ToString().Trim());
        //            command.Parameters.Add("@OrderStatus",SqlDbType.NVarChar).Value=table.Rows[i][11].ToString().Trim();
        //            command.Parameters.Add("@OrderType",SqlDbType.NVarChar).Value=table.Rows[i][12].ToString().Trim();
        //            command.Parameters.Add("@SingleAccount",SqlDbType.NVarChar).Value=table.Rows[i][13].ToString().Trim();
        //            command.Parameters.Add("@CustomerName",SqlDbType.NVarChar).Value=table.Rows[i][14].ToString().Trim();
        //            command.Parameters.Add("@CustomerAddress",SqlDbType.NVarChar).Value=table.Rows[i][15].ToString().Trim();
        //            command.Parameters.Add("@TelPhone",SqlDbType.NVarChar).Value=table.Rows[i][16].ToString().Trim();
        //            command.Parameters.Add("@OrderRemarks",SqlDbType.NVarChar).Value=table.Rows[i][17].ToString().Trim();
        //            command.Parameters.Add("@InvoiceType",SqlDbType.NVarChar).Value=table.Rows[i][18].ToString().Trim();
        //            command.Parameters.Add("@InvoicesHead",SqlDbType.NVarChar).Value=table.Rows[i][19].ToString().Trim();
        //            command.Parameters.Add("@InvoicesContent",SqlDbType.NVarChar).Value=table.Rows[i][20].ToString().Trim();
        //            command.Parameters.Add("@MerchantRemark", SqlDbType.NVarChar).Value = table.Rows[i][21].ToString().Trim();
        //            command.Parameters.Add("@MerchantRemarkRate", SqlDbType.NVarChar).Value = table.Rows[i][22].ToString().Trim();
        //            command.Parameters.Add("@FreightAmount", SqlDbType.Decimal).Value = string.IsNullOrEmpty(table.Rows[i][23].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][23].ToString().Trim());
        //            command.Parameters.Add("@PaymentConfirmTime", SqlDbType.DateTime).Value = string.IsNullOrEmpty(table.Rows[i][24].ToString().Trim()) ?DateTime.MaxValue:DateTime.Parse(table.Rows[i][24].ToString().Trim());
        //            var VATInvoice = table.Rows[i][25].ToString().Trim();
        //            string[] VATInvoiceSplit = VATInvoice.Split(',');
        //            command.Parameters.Add("@VATInvoice", SqlDbType.Bit).Value = string.IsNullOrEmpty(VATInvoice) ? false : true;
        //            command.Parameters.Add("@TaxpayerIdentificationNum", SqlDbType.NVarChar).Value = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[0]) ? string.Empty : VATInvoiceSplit[0].Substring(VATInvoiceSplit[0].IndexOf(':')+1);
        //            command.Parameters.Add("@BankAccount", SqlDbType.NVarChar).Value = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[1]) ? string.Empty : VATInvoiceSplit[1].Substring(VATInvoiceSplit[1].IndexOf(':') + 1);
        //            command.Parameters.Add("@Bank", SqlDbType.NVarChar).Value = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[2]) ? string.Empty : VATInvoiceSplit[2].Substring(VATInvoiceSplit[2].IndexOf(':') + 1);
        //            command.Parameters.Add("@RegistrationPhone", SqlDbType.NVarChar).Value = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[3]) ? string.Empty : VATInvoiceSplit[3].Substring(VATInvoiceSplit[3].IndexOf(':') + 1);
        //            command.Parameters.Add("@RegisteredAddress", SqlDbType.NVarChar).Value = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[4]) ? string.Empty : VATInvoiceSplit[4].Substring(VATInvoiceSplit[4].IndexOf(':') + 1);
        //            command.ExecuteNonQuery();    
        //        }
        //        transaction.Complete();
        //        conn.Close();   
        //    }
        //    ViewBag.Success = "导入成功";
        //    return View();
        //}


        [HttpPost]
        public ActionResult Index(HttpPostedFileBase filebase, string sourceSelect)
        {
            if (string.IsNullOrEmpty(sourceSelect) || sourceSelect == "0")
            {
                ViewBag.error = "请先选择数据来源";
                return View();
            }
            HttpPostedFileBase file = Request.Files["files"];
            string FileName;
            string savePath;
            if (file == null || file.ContentLength <= 0)
            {
                ViewBag.error = "文件不能为空";
                return View();
            }
            else
            {
                string filename = Path.GetFileName(file.FileName);
                int filesize = file.ContentLength;//获取上传文件的大小单位为字节byte
                string fileEx = System.IO.Path.GetExtension(filename);//获取上传文件的扩展名
                string NoFileName = System.IO.Path.GetFileNameWithoutExtension(filename);//获取无扩展名的文件名
                string FileType = ".xls,.xlsx";//定义上传文件的类型字符串

                FileName = NoFileName + DateTime.Now.ToString("yyyyMMddhhmmss") + fileEx;
                if (!FileType.Contains(fileEx))
                {
                    ViewBag.error = "文件类型不对，只能导入xls和xlsx格式的文件";
                    return View();
                }
                string path = AppDomain.CurrentDomain.BaseDirectory + "uploads/";
                savePath = Path.Combine(path, FileName);
                file.SaveAs(savePath);
            }

            //string result = string.Empty;
            /*
            //2003（Microsoft.Jet.Oledb.4.0）
            string strConn = string.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
            //2010（Microsoft.ACE.OLEDB.12.0）
            string strConn = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'", excelFilePath);
            */
            string strConn;
            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + savePath + ";" + "Extended Properties=Excel 8.0";
            OleDbConnection conn = new OleDbConnection(strConn);
            conn.Open();
            OleDbDataAdapter myCommand = new OleDbDataAdapter("select * from [Sheet1$]", strConn);
            DataSet myDataSet = new DataSet();
            try
            {
                myCommand.Fill(myDataSet, "ExcelInfo");
            }
            catch (Exception ex)
            {
                ViewBag.error = ex.Message;
                return View();
            }
            DataTable table = myDataSet.Tables["ExcelInfo"].DefaultView.ToTable();

            //引用事务机制，出错时，事物回滚
            using (TransactionScope transaction = new TransactionScope())
            {
                SqlConnection connection = new SqlConnection(connectionString);
                string sql = "delete from Saleorders";
                connection.Open();
                SqlCommand cmd = new SqlCommand(sql, connection);
                cmd.ExecuteNonQuery();
                // ...用foreach把tab中数据添加到数据库
                DataTable newDT = new DataTable();
                newDT.Columns.Add("OrderNum", typeof(string));
                newDT.Columns.Add("GoodsId", typeof(string));
                newDT.Columns.Add("GoodsName", typeof(string));
                newDT.Columns.Add("OrderQuantity", typeof(int));
                newDT.Columns.Add("Payment", typeof(string));
                newDT.Columns.Add("OrderTime", typeof(DateTime));
                newDT.Columns.Add("JdPrice", typeof(decimal));
                newDT.Columns.Add("OrderAmount", typeof(decimal));
                newDT.Columns.Add("SettlementAmount", typeof(decimal));
                newDT.Columns.Add("BalancePayment", typeof(decimal));
                newDT.Columns.Add("NeedAmount", typeof(decimal));
                newDT.Columns.Add("OrderStatus", typeof(string));
                newDT.Columns.Add("OrderType", typeof(string));
                newDT.Columns.Add("SingleAccount", typeof(string));
                newDT.Columns.Add("CustomerName", typeof(string));
                newDT.Columns.Add("CustomerAddress", typeof(string));
                newDT.Columns.Add("TelPhone", typeof(string));
                newDT.Columns.Add("OrderRemarks", typeof(string));
                newDT.Columns.Add("InvoiceType", typeof(string));
                newDT.Columns.Add("InvoicesHead", typeof(string));
                newDT.Columns.Add("InvoicesContent", typeof(string));
                newDT.Columns.Add("MerchantRemark", typeof(string));
                newDT.Columns.Add("MerchantRemarkRate", typeof(string));
                newDT.Columns.Add("FreightAmount", typeof(decimal));
                newDT.Columns.Add("PaymentConfirmTime", typeof(DateTime));
                newDT.Columns.Add("VATInvoice", typeof(bool));
                newDT.Columns.Add("TaxpayerIdentificationNum", typeof(string));
                newDT.Columns.Add("BankAccount", typeof(string));
                newDT.Columns.Add("Bank", typeof(string));
                newDT.Columns.Add("RegistrationPhone", typeof(string));
                newDT.Columns.Add("RegisteredAddress", typeof(string));
                for (int i = 0; i < table.Rows.Count ; i++)
                {
                    DataRow newRow = newDT.NewRow();
                    newRow["OrderNum"] = table.Rows[i][0].ToString().Trim();
                    newRow["GoodsId"] = table.Rows[i][1].ToString().Trim();
                    newRow["GoodsName"] = table.Rows[i][2].ToString().Trim();
                    newRow["OrderQuantity"] = string.IsNullOrEmpty(table.Rows[i][3].ToString().Trim()) ? 0 : int.Parse(table.Rows[i][3].ToString().Trim());
                    newRow["Payment"] = table.Rows[i][4].ToString().Trim();
                    newRow["OrderTime"] = string.IsNullOrEmpty(table.Rows[i][5].ToString().Trim()) ? DateTime.MaxValue : DateTime.Parse(table.Rows[i][5].ToString().Trim());
                    newRow["JdPrice"] = string.IsNullOrEmpty(table.Rows[i][6].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][6].ToString().Trim());
                    newRow["OrderAmount"] = string.IsNullOrEmpty(table.Rows[i][7].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][7].ToString().Trim());
                    newRow["SettlementAmount"] = string.IsNullOrEmpty(table.Rows[i][8].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][8].ToString().Trim());
                    newRow["BalancePayment"] = string.IsNullOrEmpty(table.Rows[i][9].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][9].ToString().Trim());
                    newRow["NeedAmount"] = string.IsNullOrEmpty(table.Rows[i][10].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][10].ToString().Trim());
                    newRow["OrderStatus"] = table.Rows[i][11].ToString().Trim();
                    newRow["OrderType"] = table.Rows[i][12].ToString().Trim();
                    newRow["SingleAccount"] = table.Rows[i][13].ToString().Trim();
                    newRow["CustomerName"] = table.Rows[i][14].ToString().Trim();
                    newRow["CustomerAddress"] = table.Rows[i][15].ToString().Trim();
                    newRow["TelPhone"] = table.Rows[i][16].ToString().Trim();
                    newRow["OrderRemarks"] = table.Rows[i][17].ToString().Trim();
                    newRow["InvoiceType"] = table.Rows[i][18].ToString().Trim();
                    newRow["InvoicesHead"] = table.Rows[i][19].ToString().Trim();
                    newRow["InvoicesContent"] = table.Rows[i][20].ToString().Trim();
                    newRow["MerchantRemark"] = table.Rows[i][21].ToString().Trim();
                    newRow["MerchantRemarkRate"] = table.Rows[i][22].ToString().Trim();
                    newRow["FreightAmount"] = string.IsNullOrEmpty(table.Rows[i][23].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][23].ToString().Trim());
                    newRow["PaymentConfirmTime"] = string.IsNullOrEmpty(table.Rows[i][24].ToString().Trim()) ? DateTime.MaxValue : DateTime.Parse(table.Rows[i][24].ToString().Trim());
                    var VATInvoice = table.Rows[i][25].ToString().Trim();
                    string[] VATInvoiceSplit = VATInvoice.Split(',');
                    newRow["VATInvoice"] = string.IsNullOrEmpty(VATInvoice) ? false : true;
                    newRow["TaxpayerIdentificationNum"] = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[0]) ? string.Empty : VATInvoiceSplit[0].Substring(VATInvoiceSplit[0].IndexOf(':') + 1);
                    newRow["BankAccount"] = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[1]) ? string.Empty : VATInvoiceSplit[1].Substring(VATInvoiceSplit[1].IndexOf(':') + 1);
                    newRow["Bank"] = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[2]) ? string.Empty : VATInvoiceSplit[2].Substring(VATInvoiceSplit[2].IndexOf(':') + 1);
                    newRow["RegistrationPhone"] = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[3]) ? string.Empty : VATInvoiceSplit[3].Substring(VATInvoiceSplit[3].IndexOf(':') + 1);
                    newRow["RegisteredAddress"] = string.IsNullOrEmpty(VATInvoice) ? string.Empty : string.IsNullOrEmpty(VATInvoiceSplit[4]) ? string.Empty : VATInvoiceSplit[4].Substring(VATInvoiceSplit[4].IndexOf(':') + 1);
                    newDT.Rows.Add(newRow);
                }
                SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);
                sqlbulkcopy.DestinationTableName = "Saleorders";//数据库中的表名
                sqlbulkcopy.ColumnMappings.Add("OrderNum", "OrderNum");
                sqlbulkcopy.ColumnMappings.Add("GoodsId", "GoodsId");
                sqlbulkcopy.ColumnMappings.Add("GoodsName", "GoodsName");
                sqlbulkcopy.ColumnMappings.Add("OrderQuantity", "OrderQuantity");
                sqlbulkcopy.ColumnMappings.Add("Payment", "Payment");
                sqlbulkcopy.ColumnMappings.Add("OrderTime", "OrderTime");
                sqlbulkcopy.ColumnMappings.Add("JdPrice", "JdPrice");
                sqlbulkcopy.ColumnMappings.Add("OrderAmount", "OrderAmount");
                sqlbulkcopy.ColumnMappings.Add("SettlementAmount", "SettlementAmount");
                sqlbulkcopy.ColumnMappings.Add("BalancePayment", "BalancePayment");
                sqlbulkcopy.ColumnMappings.Add("NeedAmount", "NeedAmount");
                sqlbulkcopy.ColumnMappings.Add("OrderStatus", "OrderStatus");
                sqlbulkcopy.ColumnMappings.Add("OrderType", "OrderType");
                sqlbulkcopy.ColumnMappings.Add("SingleAccount", "SingleAccount");
                sqlbulkcopy.ColumnMappings.Add("CustomerName", "CustomerName");
                sqlbulkcopy.ColumnMappings.Add("CustomerAddress", "CustomerAddress");
                sqlbulkcopy.ColumnMappings.Add("TelPhone", "TelPhone");
                sqlbulkcopy.ColumnMappings.Add("OrderRemarks", "OrderRemarks");
                sqlbulkcopy.ColumnMappings.Add("InvoiceType", "InvoiceType");
                sqlbulkcopy.ColumnMappings.Add("InvoicesHead", "InvoicesHead");
                sqlbulkcopy.ColumnMappings.Add("InvoicesContent", "InvoicesContent");
                sqlbulkcopy.ColumnMappings.Add("MerchantRemark", "MerchantRemark");
                sqlbulkcopy.ColumnMappings.Add("MerchantRemarkRate", "MerchantRemarkRate");
                sqlbulkcopy.ColumnMappings.Add("FreightAmount", "FreightAmount");
                sqlbulkcopy.ColumnMappings.Add("PaymentConfirmTime", "PaymentConfirmTime");
                sqlbulkcopy.ColumnMappings.Add("VATInvoice", "VATInvoice");
                sqlbulkcopy.ColumnMappings.Add("TaxpayerIdentificationNum", "TaxpayerIdentificationNum");
                sqlbulkcopy.ColumnMappings.Add("BankAccount", "BankAccount");
                sqlbulkcopy.ColumnMappings.Add("Bank", "Bank");
                sqlbulkcopy.ColumnMappings.Add("RegistrationPhone", "RegistrationPhone");
                sqlbulkcopy.ColumnMappings.Add("RegisteredAddress", "RegisteredAddress");
                sqlbulkcopy.WriteToServer(newDT);                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                     
                conn.Close();
                transaction.Complete();  
            }
            ViewBag.Success = "导入成功";
            return View();
        }

        [HttpPost]
        public ActionResult ExportToExcel(DateTime? beginDateTime, DateTime? endDateTime)
        {
            DataTable dt=new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("JDSaleOrderProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt=ds.Tables[0];
            }
            if (dt.Rows.Count > 0)
            {
                VoidExportToExcel("Index", dt, Response);
                return RedirectToAction("Index/", "Home");
            }
            else
            {
                return RedirectToAction("Index/", "Home");
            }
        }

        /// <summary>
        /// 导出
        /// </summary>
        /// <param name="actionTarget"></param>
        /// <param name="expData"></param>
        /// <param name="bs"></param>
        public void VoidExportToExcel(string actionTarget, DataTable expData, HttpResponseBase bs)
        {
            var grid = new GridView { DataSource = expData };
            grid.DataBind();
            bs.ClearContent();
            bs.ContentType = "application/ms-excel";
            bs.Charset = "UTF-8";
            bs.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8");
            bs.AddHeader("content-disposition", "attachment; filename=京东数据处理.xls");
            string strStyle = "<style>td{mso-number-format:\"\\@\";}</style>";
            var sw = new StringWriter();
            var htw = new HtmlTextWriter(sw);
            sw.WriteLine(strStyle);
            grid.RenderControl(htw);
            bs.Write(sw.ToString());
            bs.End();
        }

        #endregion

        public string GetData()
        {
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("JDSaleOrderProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            return JsonTableHelper.ToJson(dt);
            
        }
    }

    /// <summary>
    /// 写成DataTable的扩展方法是这样
    /// </summary>
    public static class JsonTableHelper
    {
        /// <summary> 
        /// 返回对象序列化 
        /// </summary> 
        /// <param name="obj">源对象</param> 
        /// <returns>json数据</returns> 
        public static string ToJson(object obj)
        {
            JavaScriptSerializer serialize = new JavaScriptSerializer();
            return serialize.Serialize(obj);
        }

        /// <summary> 
        /// 控制深度 
        /// </summary> 
        /// <param name="obj">源对象</param> 
        /// <param name="recursionDepth">深度</param> 
        /// <returns>json数据</returns> 
        public static string ToJson(object obj, int recursionDepth)
        {
            JavaScriptSerializer serialize = new JavaScriptSerializer();
            serialize.RecursionLimit = recursionDepth;
            return serialize.Serialize(obj);
        }

        /// <summary> 
        /// DataTable转为json 
        /// </summary> 
        /// <param name="dt">DataTable</param> 
        /// <returns>json数据</returns> 
        public static string ToJson(this DataTable dt)
        {
            List<object> dic = new List<object>();

            foreach (DataRow dr in dt.Rows)
            {
                Dictionary<string, object> result = new Dictionary<string, object>();

                foreach (DataColumn dc in dt.Columns)
                {
                    result.Add(dc.ColumnName, dr[dc].ToString());
                }
                dic.Add(result);
            }
            return ToJson(dic);
        }
    }
}