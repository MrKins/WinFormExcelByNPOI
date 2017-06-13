using NPOI.SS.UserModel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WinFormExcelByNPOI.Common
{
    public class ExcelCellStyle
    {
        public ICellStyle style { get; }
        public ICellStyle styleLeft { get; }
        public ICellStyle style4LemonChiffon { get; }
        public ICellStyle style4CornflowerBlue { get; }
        public ICellStyle style4LightGreen { get; }
        public ICellStyle style4Lime { get; }
        public ICellStyle cellStyleDate { get; }
        public ICellStyle cellStyleSmallDate { get; }
        public ICellStyle cellStyleNumber4LemonChiffon { get; }
        public ICellStyle cellStyleNumber4CornflowerBlue { get; }
        public ICellStyle cellStyleNumber4LightGreen { get; }
        public ICellStyle cellStyleNumber4Lime { get; }
        public ICellStyle cellStyleNumber { get; }
        public ICellStyle cellStyleDouble { get; }
        public ICellStyle cellStyleCurrency { get; }
        public ICellStyle cellStyleCurrency4LemonChiffon { get; }
        public ICellStyle cellStyleCurrency4CornflowerBlue { get; }
        public ICellStyle cellStyleCurrency4LightGreen { get; }
        public ICellStyle cellStyleCurrency4Lime { get; }

        public ExcelCellStyle(NPOI.HSSF.UserModel.HSSFWorkbook book)
        {
            /*新建名为style的（普通）CellStyle，水平垂直居中*/
            style = book.CreateCellStyle();
            style.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            style.BorderBottom = BorderStyle.Thin;
            style.BorderTop = BorderStyle.Thin;
            style.BorderLeft = BorderStyle.Thin;
            style.BorderRight = BorderStyle.Thin;

            /*新建名为styleLeft的（普通）CellStyle，水平垂直居左*/
            styleLeft = book.CreateCellStyle();
            styleLeft.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;
            styleLeft.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            styleLeft.BorderBottom = BorderStyle.Thin;
            styleLeft.BorderTop = BorderStyle.Thin;
            styleLeft.BorderLeft = BorderStyle.Thin;
            styleLeft.BorderRight = BorderStyle.Thin;

            /*新建名为style4LemonChiffon的（普通）CellStyle，水平垂直居中*/
            style4LemonChiffon = book.CreateCellStyle();
            style4LemonChiffon.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style4LemonChiffon.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            style4LemonChiffon.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LemonChiffon.Index;
            style4LemonChiffon.FillPattern = FillPattern.SolidForeground;
            style4LemonChiffon.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LemonChiffon.Index;
            style4LemonChiffon.BorderBottom = BorderStyle.Thin;
            style4LemonChiffon.BorderTop = BorderStyle.Thin;
            style4LemonChiffon.BorderLeft = BorderStyle.Thin;
            style4LemonChiffon.BorderRight = BorderStyle.Thin;

            /*新建名为style4CornflowerBlue的（普通）CellStyle，水平垂直居中*/
            style4CornflowerBlue = book.CreateCellStyle();
            style4CornflowerBlue.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style4CornflowerBlue.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            style4CornflowerBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            style4CornflowerBlue.FillPattern = FillPattern.SolidForeground;
            style4CornflowerBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            style4CornflowerBlue.BorderBottom = BorderStyle.Thin;
            style4CornflowerBlue.BorderTop = BorderStyle.Thin;
            style4CornflowerBlue.BorderLeft = BorderStyle.Thin;
            style4CornflowerBlue.BorderRight = BorderStyle.Thin;

            /*新建名为style4LightGreen的（普通）CellStyle，水平垂直居中*/
            style4LightGreen = book.CreateCellStyle();
            style4LightGreen.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style4LightGreen.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            style4LightGreen.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            style4LightGreen.FillPattern = FillPattern.SolidForeground;
            style4LightGreen.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            style4LightGreen.BorderBottom = BorderStyle.Thin;
            style4LightGreen.BorderTop = BorderStyle.Thin;
            style4LightGreen.BorderLeft = BorderStyle.Thin;
            style4LightGreen.BorderRight = BorderStyle.Thin;

            /*新建名为style4Lime的（普通）CellStyle，水平垂直居中*/
            style4Lime = book.CreateCellStyle();
            style4Lime.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            style4Lime.VerticalAlignment = NPOI.SS.UserModel.VerticalAlignment.Center;
            style4Lime.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            style4Lime.FillPattern = FillPattern.SolidForeground;
            style4Lime.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            style4Lime.BorderBottom = BorderStyle.Thin;
            style4Lime.BorderTop = BorderStyle.Thin;
            style4Lime.BorderLeft = BorderStyle.Thin;
            style4Lime.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleDate的（日期用）CellStyle，水平居中，日期格式化为yyyy-MM-dd HH:mm:ss*/
            cellStyleDate = book.CreateCellStyle();
            IDataFormat formatDate = book.CreateDataFormat();
            cellStyleDate.DataFormat = formatDate.GetFormat("yyyy-MM-dd HH:mm:ss");
            cellStyleDate.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            cellStyleDate.BorderBottom = BorderStyle.Thin;
            cellStyleDate.BorderTop = BorderStyle.Thin;
            cellStyleDate.BorderLeft = BorderStyle.Thin;
            cellStyleDate.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleSmallDate的（日期用）CellStyle，水平居中，日期格式化为yyyy-MM-dd*/
            cellStyleSmallDate = book.CreateCellStyle();
            IDataFormat formatSmallDate = book.CreateDataFormat();
            cellStyleSmallDate.DataFormat = formatSmallDate.GetFormat("yyyy-MM-dd");
            cellStyleSmallDate.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;
            cellStyleSmallDate.BorderBottom = BorderStyle.Thin;
            cellStyleSmallDate.BorderTop = BorderStyle.Thin;
            cellStyleSmallDate.BorderLeft = BorderStyle.Thin;
            cellStyleSmallDate.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleNumber的（千位分割数字）CellStyle，水平居中，整数化*/
            cellStyleNumber = book.CreateCellStyle();
            IDataFormat formatNumber = book.CreateDataFormat();
            cellStyleNumber.DataFormat = formatNumber.GetFormat("#,##0");
            cellStyleNumber.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleNumber.BorderBottom = BorderStyle.Thin;
            cellStyleNumber.BorderTop = BorderStyle.Thin;
            cellStyleNumber.BorderLeft = BorderStyle.Thin;
            cellStyleNumber.BorderRight = BorderStyle.Thin;

            /*cellStyleCurrency（千位分割货币）CellStyle，水平居中，整数化*/
            cellStyleCurrency = book.CreateCellStyle();
            IDataFormat formatCurrency = book.CreateDataFormat();
            cellStyleCurrency.DataFormat = formatCurrency.GetFormat("#,##0.00");
            cellStyleCurrency.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleCurrency.BorderBottom = BorderStyle.Thin;
            cellStyleCurrency.BorderTop = BorderStyle.Thin;
            cellStyleCurrency.BorderLeft = BorderStyle.Thin;
            cellStyleCurrency.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleDouble的（折扣用）CellStyle，水平居中，整数化*/
            cellStyleDouble = book.CreateCellStyle();
            IDataFormat formatDoubel = book.CreateDataFormat();
            cellStyleDouble.DataFormat = formatDoubel.GetFormat("0.00");
            cellStyleDouble.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleDouble.BorderBottom = BorderStyle.Thin;
            cellStyleDouble.BorderTop = BorderStyle.Thin;
            cellStyleDouble.BorderLeft = BorderStyle.Thin;
            cellStyleDouble.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleNumber4LemonChiffon的（千位分割数字用）CellStyle，颜色LemonChiffon，水平剧中*/
            cellStyleNumber4LemonChiffon = book.CreateCellStyle();
            IDataFormat formatNumber4LemonChiffon = book.CreateDataFormat();
            cellStyleNumber4LemonChiffon.DataFormat = formatNumber4LemonChiffon.GetFormat("#,##0");
            cellStyleNumber4LemonChiffon.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleNumber4LemonChiffon.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LemonChiffon.Index;
            cellStyleNumber4LemonChiffon.FillPattern = FillPattern.SolidForeground;
            cellStyleNumber4LemonChiffon.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LemonChiffon.Index;
            cellStyleNumber4LemonChiffon.BorderBottom = BorderStyle.Thin;
            cellStyleNumber4LemonChiffon.BorderTop = BorderStyle.Thin;
            cellStyleNumber4LemonChiffon.BorderLeft = BorderStyle.Thin;
            cellStyleNumber4LemonChiffon.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleNumber4CornflowerBlue的（千位分割数字用）CellStyle，颜色CornflowerBlue，水平剧中*/
            cellStyleNumber4CornflowerBlue = book.CreateCellStyle();
            IDataFormat formatNumber4CornflowerBlue = book.CreateDataFormat();
            cellStyleNumber4CornflowerBlue.DataFormat = formatNumber4CornflowerBlue.GetFormat("#,##0");
            cellStyleNumber4CornflowerBlue.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleNumber4CornflowerBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            cellStyleNumber4CornflowerBlue.FillPattern = FillPattern.SolidForeground;
            cellStyleNumber4CornflowerBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            cellStyleNumber4CornflowerBlue.BorderBottom = BorderStyle.Thin;
            cellStyleNumber4CornflowerBlue.BorderTop = BorderStyle.Thin;
            cellStyleNumber4CornflowerBlue.BorderLeft = BorderStyle.Thin;
            cellStyleNumber4CornflowerBlue.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleNumber4LightGreen的（千位分割数字用）CellStyle，颜色LightGreen，水平剧中*/
            cellStyleNumber4LightGreen = book.CreateCellStyle();
            IDataFormat formatNumber4LightGreen = book.CreateDataFormat();
            cellStyleNumber4LightGreen.DataFormat = formatNumber4LightGreen.GetFormat("#,##0");
            cellStyleNumber4LightGreen.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleNumber4LightGreen.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            cellStyleNumber4LightGreen.FillPattern = FillPattern.SolidForeground;
            cellStyleNumber4LightGreen.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            cellStyleNumber4LightGreen.BorderBottom = BorderStyle.Thin;
            cellStyleNumber4LightGreen.BorderTop = BorderStyle.Thin;
            cellStyleNumber4LightGreen.BorderLeft = BorderStyle.Thin;
            cellStyleNumber4LightGreen.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleNumber4Lime的（千位分割数字用）CellStyle，颜色Lime，水平剧中*/
            cellStyleNumber4Lime = book.CreateCellStyle();
            IDataFormat formatNumber4Lime = book.CreateDataFormat();
            cellStyleNumber4Lime.DataFormat = formatNumber4Lime.GetFormat("#,##0");
            cellStyleNumber4Lime.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleNumber4Lime.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            cellStyleNumber4Lime.FillPattern = FillPattern.SolidForeground;
            cellStyleNumber4Lime.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            cellStyleNumber4Lime.BorderBottom = BorderStyle.Thin;
            cellStyleNumber4Lime.BorderTop = BorderStyle.Thin;
            cellStyleNumber4Lime.BorderLeft = BorderStyle.Thin;
            cellStyleNumber4Lime.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleCurrency4LemonChiffon的（千位分割数字用）CellStyle，颜色LemonChiffon，水平剧中*/
            cellStyleCurrency4LemonChiffon = book.CreateCellStyle();
            IDataFormat formatCurrency4LemonChiffon = book.CreateDataFormat();
            cellStyleCurrency4LemonChiffon.DataFormat = formatCurrency4LemonChiffon.GetFormat("#,##0.00");
            cellStyleCurrency4LemonChiffon.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleCurrency4LemonChiffon.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LemonChiffon.Index;
            cellStyleCurrency4LemonChiffon.FillPattern = FillPattern.SolidForeground;
            cellStyleCurrency4LemonChiffon.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LemonChiffon.Index;
            cellStyleCurrency4LemonChiffon.BorderBottom = BorderStyle.Thin;
            cellStyleCurrency4LemonChiffon.BorderTop = BorderStyle.Thin;
            cellStyleCurrency4LemonChiffon.BorderLeft = BorderStyle.Thin;
            cellStyleCurrency4LemonChiffon.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleCurrency4CornflowerBlue的（千位分割数字用）CellStyle，颜色CornflowerBlue，水平剧中*/
            cellStyleCurrency4CornflowerBlue = book.CreateCellStyle();
            IDataFormat formatCurrency4CornflowerBlue = book.CreateDataFormat();
            cellStyleCurrency4CornflowerBlue.DataFormat = formatCurrency4CornflowerBlue.GetFormat("#,##0.00");
            cellStyleCurrency4CornflowerBlue.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleCurrency4CornflowerBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            cellStyleCurrency4CornflowerBlue.FillPattern = FillPattern.SolidForeground;
            cellStyleCurrency4CornflowerBlue.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.CornflowerBlue.Index;
            cellStyleCurrency4CornflowerBlue.BorderBottom = BorderStyle.Thin;
            cellStyleCurrency4CornflowerBlue.BorderTop = BorderStyle.Thin;
            cellStyleCurrency4CornflowerBlue.BorderLeft = BorderStyle.Thin;
            cellStyleCurrency4CornflowerBlue.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleCurrency4LightGreen的（千位分割数字用）CellStyle，颜色LightGreen，水平剧中*/
            cellStyleCurrency4LightGreen = book.CreateCellStyle();
            IDataFormat formatCurrency4LightGreen = book.CreateDataFormat();
            cellStyleCurrency4LightGreen.DataFormat = formatCurrency4LightGreen.GetFormat("#,##0.00");
            cellStyleCurrency4LightGreen.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleCurrency4LightGreen.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            cellStyleCurrency4LightGreen.FillPattern = FillPattern.SolidForeground;
            cellStyleCurrency4LightGreen.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.LightGreen.Index;
            cellStyleCurrency4LightGreen.BorderBottom = BorderStyle.Thin;
            cellStyleCurrency4LightGreen.BorderTop = BorderStyle.Thin;
            cellStyleCurrency4LightGreen.BorderLeft = BorderStyle.Thin;
            cellStyleCurrency4LightGreen.BorderRight = BorderStyle.Thin;

            /*新建名为cellStyleCurrency4Lime的（千位分割数字用）CellStyle，颜色Lime，水平剧中*/
            cellStyleCurrency4Lime = book.CreateCellStyle();
            IDataFormat formatCurrency4Lime = book.CreateDataFormat();
            cellStyleCurrency4Lime.DataFormat = formatCurrency4Lime.GetFormat("#,##0.00");
            cellStyleCurrency4Lime.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;
            cellStyleCurrency4Lime.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            cellStyleCurrency4Lime.FillPattern = FillPattern.SolidForeground;
            cellStyleCurrency4Lime.FillForegroundColor = NPOI.HSSF.Util.HSSFColor.Lime.Index;
            cellStyleCurrency4Lime.BorderBottom = BorderStyle.Thin;
            cellStyleCurrency4Lime.BorderTop = BorderStyle.Thin;
            cellStyleCurrency4Lime.BorderLeft = BorderStyle.Thin;
            cellStyleCurrency4Lime.BorderRight = BorderStyle.Thin;
        }
    }
}
