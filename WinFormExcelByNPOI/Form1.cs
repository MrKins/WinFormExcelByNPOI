using Mirabeau.MsSql.Library;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using WinFormExcelByNPOI.Common;

namespace WinFormExcelByNPOI
{
    public partial class Form1 : Form
    {
        private FileInfo fileExcelLoad;
        private string fileNameLoad;

        public Form1()
        {
            InitializeComponent();
        }

        private void ImportExcel(FileInfo file)
        {
            string fileName = file.FullName;

            /*You konw that? Crash is AWESOME*/
            //try

            /*读取文件*/
            using (FileStream fileStream = File.OpenRead(fileName))
            {
                HSSFWorkbook workBook = new HSSFWorkbook(fileStream);
                /*读取第一个sheet，零基*/
                ISheet sheet = workBook.GetSheetAt(0);

                /*通常第一行是标题行，从第二行开始进行循环*/
                for (int j = 1; j <= sheet.LastRowNum; j++)
                {
                    /*找到Excel对应行数的行*/
                    IRow dataRow = sheet.GetRow(j);

                    string shopID = dataRow.GetCell(2).ToString().Trim();//读取第j行的第3列数据(零基)
                    string level = dataRow.GetCell(6).ToString().Trim();//读取第j行的第7列数据(零基)

                    SetCustomer(shopID, level);//调用本地接口函数
                }
            }
        }

        private void ExportExcel()
        {
            List<Customer> listCustomer = GetCustomer("");

            /*新建表；新建Sheet并命名；设定cellStyle*/
            HSSFWorkbook book = new HSSFWorkbook();
            ISheet sheet1 = book.CreateSheet("Sheet1");
            IRow headerRow4Sheet1 = sheet1.CreateRow(0);
            ExcelCellStyle cellStyle = new ExcelCellStyle(book);

            ICell cell;

            /*设定标题行*/
            cell = headerRow4Sheet1.CreateCell(0);
            cell.CellStyle = cellStyle.style;
            cell.SetCellValue("客户代码");

            cell = headerRow4Sheet1.CreateCell(1);
            cell.CellStyle = cellStyle.style;
            cell.SetCellValue("客户名");

            cell = headerRow4Sheet1.CreateCell(2);
            cell.CellStyle = cellStyle.style;
            cell.SetCellValue("客户地址");

            /*设定内容行*/
            int sheet1RowID = 1;//从表格的第二2行开始进行循环

            foreach (Customer customer in listCustomer)
            {
                IRow r = sheet1.CreateRow(sheet1RowID);

                cell = r.CreateCell(0); cell.SetCellValue(customer.customerID); cell.CellStyle = cellStyle.style;
                cell = r.CreateCell(1); cell.SetCellValue(customer.customerName); cell.CellStyle = cellStyle.style;
                cell = r.CreateCell(2); cell.SetCellValue(customer.customerAddress); cell.CellStyle = cellStyle.style;

                sheet1RowID = sheet1RowID + 1;
            }

            /*单元格长度格式化*/
            ChangeStyle(book, sheet1);

            /*IO流输出保存*/
            string saveFileName = "客户列表导出Excel";
            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.DefaultExt = "xls";
            saveDialog.Filter = "Excel文件|*.xls";
            saveDialog.FileName = saveFileName;
            saveDialog.ShowDialog();
            saveFileName = saveDialog.FileName;

            MemoryStream ms = new MemoryStream();
            book.Write(ms);
            FileStream file = new FileStream(saveFileName, FileMode.Create);
            book.Write(file);
            file.Close();
            book = null;
            ms.Close();
            ms.Dispose();
        }

        private void ChangeStyle(IWorkbook hssfworkbook, ISheet sheet)
        {
            for (int columnNum = 0; columnNum <= sheet.GetRow(0).LastCellNum + 12; columnNum++) //columnNum为列的数量
            {
                int columnWidth = sheet.GetColumnWidth(columnNum) / 256; //获取当前列宽度
                for (int rowNum = 0; rowNum <= sheet.LastRowNum; rowNum++) //在这一列上循环行
                {
                    IRow currentRow = sheet.GetRow(rowNum);
                    if (currentRow != null)
                    {
                        ICell currentCell = currentRow.GetCell(columnNum);
                        if (currentCell != null)
                        {
                            int length = Encoding.Default.GetBytes(currentCell.ToString()).Length;
                            //单元格的宽度
                            if (columnWidth < length + 1)
                            {
                                columnWidth = length + 1;
                            }
                            //若当前单元格内容宽度大于列宽，则调整列宽为当前单元格宽度，后面的+1是我人为的操作
                        }
                    }
                }
                if (columnWidth > 255)
                {
                    columnWidth = 255;//由于最大宽度是255，所以这里需要判断下，否则会报错
                }
                sheet.SetColumnWidth(columnNum, columnWidth * 256);//设置最终宽度
            }
        }

        public List<Customer> GetCustomer(string customerID)
        {
            List<Customer> listCustomer = new List<Customer>();

            string CMD = "GetCustomer";

            var parameters = new List<SqlParameter>
            {
                customerID.CreateSqlParameter("customerID")
            };

            DataSet dataSet = DatabaseHelper.ExecuteDataSet(GlobalVar.connectionString, CommandType.StoredProcedure, CMD, parameters);
            DataTable dataTable = dataSet.Tables[0];

            foreach (DataRow dataRow in dataTable.Rows)
            {
                Customer customer = new Customer();

                customer.customerID = dataRow["CustomerID"].ToString().Trim();
                customer.customerName = dataRow["CustomerName"].ToString().Trim();
                customer.customerAddress = dataRow["CustomerAddress"].ToString().Trim();

                listCustomer.Add(customer);
            }

            return listCustomer;
        }

        private void SetCustomer(string customerName, string customerAddress)
        {

            string CMD = "SetCustomer";

            var parameters = new List<SqlParameter>
            {
                customerName.CreateSqlParameter("customerName"),
                customerAddress.CreateSqlParameter("customerAddress")
            };

            DatabaseHelper.ExecuteNonQuery(GlobalVar.connectionString, CommandType.StoredProcedure, CMD, parameters);
        }

        private void btnChoseFile_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//该值确定是否可以选择多个文件
            dialog.Title = "请选择文件";
            dialog.Filter = "Excel文件(*.xls)|*.xls";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                fileNameLoad = dialog.FileName;
                fileExcelLoad = new FileInfo(fileNameLoad);
                tbFilePath.Text = fileNameLoad;
            }
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            ImportExcel(fileExcelLoad);
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            ExportExcel();
        }
    }
}
