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

            #region 使用SqlBulkCopy实现批量导入数据库（该方法可以重构）
            switch (sourceSelect)
            {
                #region 京东收款
                case "1":
                    //引用事务机制，出错时，事物回滚
                    using (TransactionScope transaction = new TransactionScope())
                    {
                        SqlConnection connection = new SqlConnection(connectionString);
                        string sql = "delete from JdReceivables";
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sql, connection);
                        cmd.ExecuteNonQuery();
                        DataTable newDT = new DataTable();
                        newDT.Columns.Add("OrderNum", typeof(string));
                        newDT.Columns.Add("BusinessOccurrenceTime", typeof(DateTime));
                        newDT.Columns.Add("BusinessNum", typeof(string));
                        newDT.Columns.Add("MerchantTypeName", typeof(string));
                        newDT.Columns.Add("BusinessTypeName", typeof(string));
                        newDT.Columns.Add("FeeTypeName", typeof(string));
                        newDT.Columns.Add("Amount", typeof(int));
                        newDT.Columns.Add("GoodsTotalAmount", typeof(decimal));
                        newDT.Columns.Add("SettlementAmount", typeof(decimal));
                        newDT.Columns.Add("ProductCode", typeof(string));
                        newDT.Columns.Add("ProductName", typeof(string));
                        newDT.Columns.Add("MerchantDeals", typeof(decimal));
                        newDT.Columns.Add("DeductionRate", typeof(decimal));
                        newDT.Columns.Add("Remark", typeof(string));
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            DataRow newRow = newDT.NewRow();
                            newRow["OrderNum"] = table.Rows[i][0].ToString().Trim();
                            newRow["BusinessOccurrenceTime"] = string.IsNullOrEmpty(table.Rows[i][1].ToString().Trim()) ? DateTime.MaxValue : DateTime.Parse(table.Rows[i][1].ToString().Trim());
                            newRow["BusinessNum"] = table.Rows[i][2].ToString().Trim();
                            newRow["MerchantTypeName"] = table.Rows[i][3].ToString().Trim();
                            newRow["BusinessTypeName"] = table.Rows[i][4].ToString().Trim();
                            newRow["FeeTypeName"] = table.Rows[i][5].ToString().Trim();
                            newRow["Amount"] = string.IsNullOrEmpty(table.Rows[i][6].ToString().Trim()) ? 0 : int.Parse(table.Rows[i][6].ToString().Trim());
                            newRow["GoodsTotalAmount"] = string.IsNullOrEmpty(table.Rows[i][7].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][7].ToString().Trim());
                            newRow["SettlementAmount"] = string.IsNullOrEmpty(table.Rows[i][8].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][8].ToString().Trim());
                            newRow["ProductCode"] = table.Rows[i][9].ToString().Trim();
                            newRow["ProductName"] = table.Rows[i][10].ToString().Trim();
                            newRow["MerchantDeals"] = string.IsNullOrEmpty(table.Rows[i][11].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][11].ToString().Trim());
                            newRow["DeductionRate"] = string.IsNullOrEmpty(table.Rows[i][12].ToString().Trim()) || table.Rows[i][12].ToString()=="null"? 0 : decimal.Parse(table.Rows[i][12].ToString().Trim());
                            newRow["Remark"] = table.Rows[i][13].ToString().Trim();
                            newDT.Rows.Add(newRow);
                        }
                        SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);
                        sqlbulkcopy.DestinationTableName = "JdReceivables";//数据库中的表名
                        foreach (System.Data.DataColumn k in newDT.Columns)
                        {
                            sqlbulkcopy.ColumnMappings.Add(k.ColumnName.ToString(), k.ColumnName.ToString());
                        }
                        sqlbulkcopy.WriteToServer(newDT);
                        conn.Close();
                        transaction.Complete();
                    }
                    break;
                #endregion
                #region 银联收款
                case "2":
                    using (TransactionScope transaction = new TransactionScope())
                    {
                        SqlConnection connection = new SqlConnection(connectionString);
                        string sql = "delete from CupReceivables";
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sql, connection);
                        cmd.ExecuteNonQuery();
                        DataTable newDT = new DataTable();
                        newDT.Columns.Add("LiquidationDate", typeof(DateTime));
                        newDT.Columns.Add("TransactionDate", typeof(DateTime));
                        newDT.Columns.Add("TerminalNum", typeof(string));
                        newDT.Columns.Add("TransactionAmount", typeof(decimal));
                        newDT.Columns.Add("LiquidationAmount", typeof(decimal));
                        newDT.Columns.Add("Fee", typeof(decimal));
                        newDT.Columns.Add("SerialNumber", typeof(string));
                        newDT.Columns.Add("TransactionType", typeof(string));
                        newDT.Columns.Add("ReferenceNum", typeof(string));
                        newDT.Columns.Add("CardNumber", typeof(string));
                        newDT.Columns.Add("CardType", typeof(string));
                        newDT.Columns.Add("IssuingBank", typeof(string));
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            DataRow newRow = newDT.NewRow();
                            newRow["LiquidationDate"] = string.IsNullOrEmpty(table.Rows[i][0].ToString().Trim()) ? DateTime.MaxValue : DateTime.ParseExact(table.Rows[i][0].ToString().Trim(), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                            newRow["TransactionDate"] = string.IsNullOrEmpty(table.Rows[i][2].ToString().Trim()) ? DateTime.MaxValue : DateTime.Parse(table.Rows[i][2].ToString().Trim());
                            newRow["TerminalNum"] = table.Rows[i][3].ToString().Trim();
                            newRow["TransactionAmount"] = string.IsNullOrEmpty(table.Rows[i][4].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][4].ToString().Trim());
                            newRow["LiquidationAmount"] = string.IsNullOrEmpty(table.Rows[i][5].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][5].ToString().Trim());
                            newRow["Fee"] = string.IsNullOrEmpty(table.Rows[i][6].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][6].ToString().Trim());
                            newRow["SerialNumber"] = table.Rows[i][7].ToString().Trim();
                            newRow["TransactionType"] = table.Rows[i][8].ToString().Trim();
                            newRow["ReferenceNum"] = table.Rows[i][9].ToString().Trim();
                            newRow["CardNumber"] = table.Rows[i][10].ToString().Trim();
                            newRow["CardType"] = table.Rows[i][11].ToString().Trim();
                            newRow["IssuingBank"] = table.Rows[i][12].ToString().Trim();
                            newDT.Rows.Add(newRow);
                        }
                        SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);
                        sqlbulkcopy.DestinationTableName = "CupReceivables";//数据库中的表名
                        foreach (System.Data.DataColumn k in newDT.Columns)
                        {
                            sqlbulkcopy.ColumnMappings.Add(k.ColumnName.ToString(), k.ColumnName.ToString());
                        }
                        sqlbulkcopy.WriteToServer(newDT);
                        conn.Close();
                        transaction.Complete();
                    }
                    break;
                #endregion
                #region 支付宝收款
                case "3":
                    using (TransactionScope transaction = new TransactionScope())
                    {
                        SqlConnection connection = new SqlConnection(connectionString);
                        string sql = "delete from AlipayReceivables";
                        connection.Open();
                        SqlCommand cmd = new SqlCommand(sql, connection);
                        cmd.ExecuteNonQuery();
                        DataTable newDT = new DataTable();
                        newDT.Columns.Add("AccountsSerialNum", typeof(string));
                        newDT.Columns.Add("BusinessSerialNum", typeof(string));
                        newDT.Columns.Add("MerchantOrderNumber", typeof(string));
                        newDT.Columns.Add("ProductName", typeof(string));
                        newDT.Columns.Add("OccurrenceTime", typeof(DateTime));
                        newDT.Columns.Add("OtherAccount", typeof(string));
                        newDT.Columns.Add("IncomeAmount", typeof(decimal));
                        newDT.Columns.Add("ExpenditureAmount", typeof(decimal));
                        newDT.Columns.Add("AccountBalance", typeof(decimal));
                        newDT.Columns.Add("TradingChannels", typeof(string));
                        newDT.Columns.Add("BusinessType", typeof(string));
                        newDT.Columns.Add("Remark", typeof(string));
                        for (int i = 0; i < table.Rows.Count; i++)
                        {
                            DataRow newRow = newDT.NewRow();
                            newRow["AccountsSerialNum"] = table.Rows[i][0].ToString().Trim();
                            newRow["BusinessSerialNum"] = table.Rows[i][1].ToString().Trim();
                            newRow["MerchantOrderNumber"] = table.Rows[i][2].ToString().Trim();
                            newRow["ProductName"] = table.Rows[i][3].ToString().Trim();
                            newRow["OccurrenceTime"] = string.IsNullOrEmpty(table.Rows[i][4].ToString().Trim()) ? DateTime.MaxValue : DateTime.Parse(table.Rows[i][4].ToString().Trim());
                            newRow["OtherAccount"] = table.Rows[i][5].ToString().Trim();
                            newRow["IncomeAmount"] = string.IsNullOrEmpty(table.Rows[i][6].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][6].ToString().Trim());
                            newRow["ExpenditureAmount"] = string.IsNullOrEmpty(table.Rows[i][7].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][7].ToString().Trim());
                            newRow["AccountBalance"] = string.IsNullOrEmpty(table.Rows[i][8].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][8].ToString().Trim());
                            newRow["TradingChannels"] = table.Rows[i][9].ToString().Trim();
                            newRow["BusinessType"] = table.Rows[i][10].ToString().Trim();
                            newRow["Remark"] = table.Rows[i][11].ToString().Trim();
                            newDT.Rows.Add(newRow);
                        }
                        SqlBulkCopy sqlbulkcopy = new SqlBulkCopy(connectionString, SqlBulkCopyOptions.UseInternalTransaction);
                        sqlbulkcopy.DestinationTableName = "AlipayReceivables";//数据库中的表名
                        foreach (System.Data.DataColumn k in newDT.Columns)
                        {
                            sqlbulkcopy.ColumnMappings.Add(k.ColumnName.ToString(), k.ColumnName.ToString());
                        }
                        sqlbulkcopy.WriteToServer(newDT);
                        conn.Close();
                        transaction.Complete();
                    }
                    break;
                #endregion
                default:
                    ViewBag.error = "请先选择数据来源";
                    return View();
            }
            #endregion
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
            //using (SqlConnection connection = new SqlConnection(connectionString))
            //{
            //    SqlCommand cm = new SqlCommand("JDSaleOrderProcedure", connection);
            //    cm.CommandType = CommandType.StoredProcedure;
            //    DataSet ds = new DataSet();
            //    SqlDataAdapter ad = new SqlDataAdapter(cm);
            //    ad.Fill(ds, "SaleOrder");
            //    dt = ds.Tables[0];
            //}
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