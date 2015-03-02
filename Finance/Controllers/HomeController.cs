using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Text.RegularExpressions;
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
                if (!System.IO.File.Exists(path))
                {

                    System.IO.FileStream fsnew = System.IO.File.Create(path);
                    fsnew.Close();
                }
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
                            newRow["BusinessOccurrenceTime"] = string.IsNullOrEmpty(table.Rows[i][1].ToString().Trim()) ? DateTime.Now : DateTime.Parse(table.Rows[i][1].ToString().Trim());
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
                            newRow["DeductionRate"] = string.IsNullOrEmpty(table.Rows[i][12].ToString().Trim()) || table.Rows[i][12].ToString() == "null" ? 0 : decimal.Parse(table.Rows[i][12].ToString().Trim());
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
                            newRow["LiquidationDate"] = string.IsNullOrEmpty(table.Rows[i][0].ToString().Trim()) ? DateTime.Now : DateTime.ParseExact(table.Rows[i][0].ToString().Trim(), "yyyyMMdd", System.Globalization.CultureInfo.CurrentCulture);
                            newRow["TransactionDate"] = string.IsNullOrEmpty(table.Rows[i][2].ToString().Trim()) ? DateTime.Now : DateTime.Parse(table.Rows[i][2].ToString().Trim());
                            newRow["TerminalNum"] = table.Rows[i][3].ToString().Trim();
                            newRow["TransactionAmount"] = string.IsNullOrEmpty(table.Rows[i][4].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][4].ToString().Trim());
                            newRow["LiquidationAmount"] = string.IsNullOrEmpty(table.Rows[i][5].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][5].ToString().Trim());
                            newRow["Fee"] = string.IsNullOrEmpty(table.Rows[i][6].ToString().Trim()) ? 0 : decimal.Parse(table.Rows[i][6].ToString().Trim());
                            newRow["SerialNumber"] = table.Rows[i][7].ToString().Trim();
                            newRow["TransactionType"] = table.Rows[i][8].ToString().Trim();
                            newRow["ReferenceNum"] = "201" + table.Rows[i][9].ToString().Trim();
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
                            newRow["OccurrenceTime"] = string.IsNullOrEmpty(table.Rows[i][4].ToString().Trim()) ? DateTime.Now : DateTime.Parse(table.Rows[i][4].ToString().Trim());
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
                #region
                case "4":
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
                        for (int i = 0; i < table.Rows.Count; i++)
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

        public ActionResult About()
        {
            return View();
        }

        public string GetData()
        {
            var beginDateTime = string.IsNullOrEmpty(Request.QueryString["beginDateTime"]) ? DateTime.Now.AddHours(-24) : DateTime.Parse(Request.QueryString["beginDateTime"]);
            var endDateTime = string.IsNullOrEmpty(Request.QueryString["endDateTime"]) ? DateTime.Now.AddHours(-24) : DateTime.Parse(Request.QueryString["endDateTime"]);
            int? bInvoiceMade = null;
            if (Request.QueryString["bInvoiceMade"] == "1")
            {
                bInvoiceMade = 1;
            }
            else if (Request.QueryString["bInvoiceMade"] == "2")
            {
                bInvoiceMade = 0;
            }
            else
            {
                bInvoiceMade = null;
            }
            int? bLogic = null;
            if (Request.QueryString["bLogic"] == "1")
            {
                bLogic = 1;
            }
            else if (Request.QueryString["bLogic"] == "2")
            {
                bLogic = 0;
            }
            else
            {
                bLogic = null;
            }
            int? bFinanceAccounted = null;
            if (Request.QueryString["bFinanceAccounted"] == "1")
            {
                bFinanceAccounted = 1;
            }
            else if (Request.QueryString["bFinanceAccounted"] == "2")
            {
                bFinanceAccounted = 0;
            }
            else
            {
                bFinanceAccounted = null;
            }
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("InvoiceFinanceProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                cm.Parameters.Add("@beginDateTime", SqlDbType.DateTime).Value = beginDateTime;
                cm.Parameters.Add("@endDateTime", SqlDbType.DateTime).Value = endDateTime;
                cm.Parameters.Add("@bInvoiceMade", SqlDbType.Bit).Value = bInvoiceMade;
                cm.Parameters.Add("@bLogic", SqlDbType.Bit).Value = bLogic;
                cm.Parameters.Add("@bFinanceAccounted", SqlDbType.Bit).Value = bFinanceAccounted;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            return JsonTableHelper.ToJson(dt);

        }

        /// <summary>
        /// 获取数据核对结果
        /// </summary>
        /// <returns></returns>
        public string GetDataCheckData()
        {
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("DataCheckProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            return JsonTableHelper.ToJson(dt);

        }

        /// <summary>
        /// 获取数据查看结果
        /// </summary>
        /// <returns></returns>
        public string GetDataViewData()
        {
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("DataViewProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            return JsonTableHelper.ToJson(dt);

        }

        [HttpPost]
        public ActionResult ExportToExcel()
        {
            var beginDateTime = string.IsNullOrEmpty(Request["beginDateTimeStr"]) ? DateTime.Now.AddHours(-24) : DateTime.Parse(Request["beginDateTimeStr"]);
            var endDateTime = string.IsNullOrEmpty(Request["endDateTimeStr"]) ? DateTime.Now.AddHours(-24) : DateTime.Parse(Request["endDateTimeStr"]);
            int? bInvoiceMade = null;
            if (Request["bInvoiceMade"] == "1")
            {
                bInvoiceMade = 1;
            }
            else if (Request["bInvoiceMade"] == "2")
            {
                bInvoiceMade = 0;
            }
            else
            {
                bInvoiceMade = null;
            }
            int? bLogic = null;
            if (Request["bLogic"] == "1")
            {
                bLogic = 1;
            }
            else if (Request["bLogic"] == "2")
            {
                bLogic = 0;
            }
            else
            {
                bLogic = null;
            }
            int? bFinanceAccounted = null;
            if (Request["bFinanceAccounted"] == "1")
            {
                bFinanceAccounted = 1;
            }
            else if (Request["bFinanceAccounted"] == "2")
            {
                bFinanceAccounted = 0;
            }
            else
            {
                bFinanceAccounted = null;
            }
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("FinanceAccountedProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                cm.Parameters.Add("@beginDateTime", SqlDbType.DateTime).Value = beginDateTime;
                cm.Parameters.Add("@endDateTime", SqlDbType.DateTime).Value = endDateTime;
                cm.Parameters.Add("@bInvoiceMade", SqlDbType.Bit).Value = bInvoiceMade;
                cm.Parameters.Add("@bLogic", SqlDbType.Bit).Value = bLogic;
                cm.Parameters.Add("@bFinanceAccounted", SqlDbType.Bit).Value = bFinanceAccounted;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            if (dt.Rows.Count > 0)
            {
                //VoidExportToExcel("Index", dt, Response, "财务入账.xls");
                Dictionary<int, int> mergeCellNums = new Dictionary<int, int>();
                mergeCellNums.Add(3, 1);
                DataTable2Excel(dt, null, "财务入账", mergeCellNums, 0);
                using (SqlConnection connectionStr = new SqlConnection(connectionString))
                {
                    connectionStr.Open();
                    SqlCommand updataCom = new SqlCommand("UpdataBFinanceAccounted", connectionStr);
                    updataCom.CommandType = CommandType.StoredProcedure;
                    updataCom.Parameters.Add("@beginDateTime", SqlDbType.DateTime).Value = beginDateTime;
                    updataCom.Parameters.Add("@endDateTime", SqlDbType.DateTime).Value = endDateTime;
                    updataCom.Parameters.Add("@bInvoiceMade", SqlDbType.Bit).Value = bInvoiceMade;
                    updataCom.Parameters.Add("@bLogic", SqlDbType.Bit).Value = bLogic;
                    updataCom.Parameters.Add("@bFinanceAccounted", SqlDbType.Bit).Value = bFinanceAccounted;
                    updataCom.ExecuteNonQuery();
                    connectionStr.Close();
                }
                return RedirectToAction("Index/", "Home");
            }
            else
            {
                return RedirectToAction("Index/", "Home");
            }
        }


        /// <summary>
        /// 导出选中数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ExportSelectToExcel()
        {
            DataTable dt = new DataTable();
            string TradeNoListStr = "," + Request["TradeNoList"] + ","; ;
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("InvoiceFinanceExportProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                cm.Parameters.Add("@TradeNoList", SqlDbType.Text).Value = TradeNoListStr;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            if (dt.Rows.Count > 0)
            {
                VoidExportToExcel("Index", dt, Response, "财务发票.xls");
                using (SqlConnection connectionStr = new SqlConnection(connectionString))
                {
                    connectionStr.Open();
                    SqlCommand updataCom = new SqlCommand("UpdataTradeListProcedure", connectionStr);
                    updataCom.CommandType = CommandType.StoredProcedure;
                    updataCom.Parameters.Add("@TradeNoList", SqlDbType.Text).Value = TradeNoListStr;
                    updataCom.ExecuteNonQuery();
                    connectionStr.Close();
                }
                return RedirectToAction("Index/", "Home");
            }
            else
            {
                return RedirectToAction("Index/", "Home");
            }
        }

        /// <summary>
        /// 导出所有数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ExportAllDataToExcel()
        {
            var beginDateTime = string.IsNullOrEmpty(Request["beginDateTimeStr"]) ? DateTime.Now.AddHours(-24) : DateTime.Parse(Request["beginDateTimeStr"]);
            var endDateTime = string.IsNullOrEmpty(Request["endDateTimeStr"]) ? DateTime.Now.AddHours(-24) : DateTime.Parse(Request["endDateTimeStr"]);
            int? bInvoiceMade = null;
            if (Request["bInvoiceMade"] == "1")
            {
                bInvoiceMade = 1;
            }
            else if (Request["bInvoiceMade"] == "2")
            {
                bInvoiceMade = 0;
            }
            else
            {
                bInvoiceMade = null;
            }
            int? bLogic = null;
            if (Request["bLogic"] == "1")
            {
                bLogic = 1;
            }
            else if (Request["bLogic"] == "2")
            {
                bLogic = 0;
            }
            else
            {
                bLogic = null;
            }
            int? bFinanceAccounted = null;
            if (Request["bFinanceAccounted"] == "1")
            {
                bFinanceAccounted = 1;
            }
            else if (Request["bFinanceAccounted"] == "2")
            {
                bFinanceAccounted = 0;
            }
            else
            {
                bFinanceAccounted = null;
            }
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("InvoiceFinanceExportProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                cm.Parameters.Add("@beginDateTime", SqlDbType.DateTime).Value = beginDateTime;
                cm.Parameters.Add("@endDateTime", SqlDbType.DateTime).Value = endDateTime;
                cm.Parameters.Add("@bInvoiceMade", SqlDbType.Bit).Value = bInvoiceMade;
                cm.Parameters.Add("@bLogic", SqlDbType.Bit).Value = bLogic;
                cm.Parameters.Add("@bFinanceAccounted", SqlDbType.Bit).Value = bFinanceAccounted;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            if (dt.Rows.Count > 0)
            {
                Dictionary<int, int> mergeCellNums = new Dictionary<int, int>();   
                mergeCellNums.Add(0, 1);  
                //VoidExportToExcel("Index", dt, Response, "财务发票.xls");
                DataTable2Excel(dt, null, "财务发票", mergeCellNums, 0);
                using (SqlConnection connectionStr = new SqlConnection(connectionString))
                {
                    connectionStr.Open();
                    SqlCommand updataCom = new SqlCommand("UpdataTradeListProcedure", connectionStr);
                    updataCom.CommandType = CommandType.StoredProcedure;
                    updataCom.Parameters.Add("@beginDateTime", SqlDbType.DateTime).Value = beginDateTime;
                    updataCom.Parameters.Add("@endDateTime", SqlDbType.DateTime).Value = endDateTime;
                    updataCom.Parameters.Add("@bInvoiceMade", SqlDbType.Bit).Value = bInvoiceMade;
                    updataCom.Parameters.Add("@bLogic", SqlDbType.Bit).Value = bLogic;
                    updataCom.Parameters.Add("@bFinanceAccounted", SqlDbType.Bit).Value = bFinanceAccounted;
                    updataCom.ExecuteNonQuery();
                    connectionStr.Close();
                }
                return RedirectToAction("Index/", "Home");
            }
            else
            {
                return RedirectToAction("Index/", "Home");
            }
        }

        /// <summary>
        /// 导出核对异常数据
        /// </summary>
        /// <returns></returns>
        [HttpPost]
        public ActionResult ExportAbnormalDataToExcel()
        {
            DataTable dt = new DataTable();
            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand cm = new SqlCommand("DataCheckAbnormalDataExportProcedure", connection);
                cm.CommandType = CommandType.StoredProcedure;
                DataSet ds = new DataSet();
                SqlDataAdapter ad = new SqlDataAdapter(cm);
                ad.Fill(ds, "SaleOrder");
                dt = ds.Tables[0];
            }
            if (dt.Rows.Count > 0)
            {
                VoidExportToExcel("Index", dt, Response, "核对异常数据.xls");
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
        public void VoidExportToExcel(string actionTarget, DataTable expData, HttpResponseBase bs, string filename)
        {
            var grid = new GridView { DataSource = expData };
            grid.DataBind();
            bs.ClearContent();
            bs.ContentType = "application/ms-excel";
            bs.Charset = "UTF-8";
            bs.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8");
            bs.AddHeader("content-disposition", string.Format("attachment; filename={0}", filename));
            string strStyle = "<style>td{mso-number-format:\"\\@\";}</style>";
            var sw = new StringWriter();
            var htw = new HtmlTextWriter(sw);
            sw.WriteLine(strStyle);
            grid.RenderControl(htw);
            bs.Write(sw.ToString());
            bs.End();
        }

        /// <summary>  
        /// 描述：把DataTable内容导出excel并返回客户端   
        /// </summary>  
        /// <param name="dtData"></param>  
        /// <param name="header"></param>  
        /// <param name="fileName"></param>  
        /// <param name="mergeCellNums">要合并的列索引字典 格式：列索引-合并模式(合并模式 1 合并相同项、2 合并空项、3 合并相同项及空项)</param>  
        /// <param name="mergeKey">作为合并项的标记列索引</param>  
        public static void DataTable2Excel(System.Data.DataTable dtData, TableCell[] header, string fileName, Dictionary<int, int> mergeCellNums, int? mergeKey)
        {
            GridView gvExport = null;
            // 当前对话   
            System.Web.HttpContext curContext = System.Web.HttpContext.Current;
            // IO用于导出并返回excel文件   
            System.IO.StringWriter strWriter = null;
            System.Web.UI.HtmlTextWriter htmlWriter = null;

            if (dtData != null)
            {
                // 设置编码和附件格式   
                curContext.Response.ContentType = "application/vnd.ms-excel";
                curContext.Response.ContentEncoding = System.Text.Encoding.GetEncoding("gb2312");
                curContext.Response.Charset = "gb2312";
                if (!string.IsNullOrEmpty(fileName))
                {
                    //处理中文名乱码问题  
                    fileName = System.Web.HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8);
                    curContext.Response.AppendHeader("Content-Disposition", ("attachment;filename=" + (fileName.ToLower().EndsWith(".xls") ? fileName : fileName + ".xls")));
                }
                // 导出excel文件   
                strWriter = new System.IO.StringWriter();
                htmlWriter = new System.Web.UI.HtmlTextWriter(strWriter);

                // 重新定义一个无分页的GridView   
                gvExport = new System.Web.UI.WebControls.GridView();
                gvExport.DataSource = dtData.DefaultView;
                gvExport.AllowPaging = false;
                //优化导出数据显示，如身份证、12-1等显示异常问题  
                gvExport.RowDataBound += new System.Web.UI.WebControls.GridViewRowEventHandler(dgExport_RowDataBound);

                gvExport.DataBind();
                //处理表头  
                if (header != null && header.Length > 0)
                {
                    gvExport.HeaderRow.Cells.Clear();
                    gvExport.HeaderRow.Cells.AddRange(header);
                }
                //合并单元格  
                if (mergeCellNums != null && mergeCellNums.Count > 0)
                {
                    foreach (int cellNum in mergeCellNums.Keys)
                    {
                        MergeRows(gvExport, cellNum, mergeCellNums[cellNum], mergeKey);
                    }
                }

                // 返回客户端   
                gvExport.RenderControl(htmlWriter);
                curContext.Response.Clear();
                curContext.Response.Write("<meta http-equiv=\"content-type\" content=\"application/ms-excel; charset=gb2312\"/>" + strWriter.ToString());
                curContext.Response.End();
            }
        }
        /// <summary>  
        /// 描述：行绑定事件  
        /// </summary>  
        /// <param name="sender"></param>  
        /// <param name="e"></param>  
        protected static void dgExport_RowDataBound(object sender, GridViewRowEventArgs e)
        {
            if (e.Row.RowType == DataControlRowType.DataRow)
            {
                foreach (TableCell cell in e.Row.Cells)
                {
                    //优化导出数据显示，如身份证、12-1等显示异常问题  
                    if (Regex.IsMatch(cell.Text.Trim(), @"^\d{12,}$") || Regex.IsMatch(cell.Text.Trim(), @"^\d+[-]\d+$"))
                    {
                        cell.Attributes.Add("style", "vnd.ms-excel.numberformat:@");
                    }
                }
            }
        }

        /// <summary>     
        /// 描述：合并GridView列中相同的行   
        /// </summary>     
        /// <param   name="gvExport">GridView对象</param>     
        /// <param   name="cellNum">需要合并的列</param>     
        /// <param name="mergeMode">合并模式 1 合并相同项、2 合并空项、3 合并相同项及空项</param>  
        /// <param name="mergeKey">作为合并项的标记列索引</param>  
        public static void MergeRows(GridView gvExport, int cellNum, int mergeMode, int? mergeKey)
        {
            int i = 0, rowSpanNum = 1;
            System.Drawing.Color alterColor = System.Drawing.Color.LightGray;
            while (i < gvExport.Rows.Count - 1)
            {
                GridViewRow gvr = gvExport.Rows[i];
                for (++i; i < gvExport.Rows.Count; i++)
                {
                    GridViewRow gvrNext = gvExport.Rows[i];
                    if ((!mergeKey.HasValue || (mergeKey.HasValue && (gvr.Cells[mergeKey.Value].Text.Equals(gvrNext.Cells[mergeKey.Value].Text) || " ".Equals(gvrNext.Cells[mergeKey.Value].Text)))) && ((mergeMode == 1 && gvr.Cells[cellNum].Text == gvrNext.Cells[cellNum].Text) || (mergeMode == 2 && " ".Equals(gvrNext.Cells[cellNum].Text.Trim())) || (mergeMode == 3 && (gvr.Cells[cellNum].Text == gvrNext.Cells[cellNum].Text || " ".Equals(gvrNext.Cells[cellNum].Text.Trim())))))
                    {
                        gvrNext.Cells[cellNum].Visible = false;
                        rowSpanNum++;
                        gvrNext.BackColor = gvr.BackColor;
                    }
                    else
                    {
                        gvr.Cells[cellNum].RowSpan = rowSpanNum;
                        rowSpanNum = 1;
                        break;
                    }
                    if (i == gvExport.Rows.Count - 1)
                    {
                        gvr.Cells[cellNum].RowSpan = rowSpanNum;
                        if (mergeKey.HasValue && cellNum == mergeKey.Value)
                        {
                            if (alterColor == System.Drawing.Color.White)
                                gvr.BackColor = System.Drawing.Color.LightGray;
                        }
                    }
                }
            }
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