using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace Axerrio.Utilities.Excel
{
    /// <summary>
    /// http://www.dofactory.com/net/chain-of-responsibility-design-pattern
    /// </summary>
    public class ExcelExport : IExcelExport
    {
        private ICellStyle _currencyStyle;
        private ICellStyle _dateStyle;

        public async Task<byte[]> ConvertJsonToExcelAsync(ExcelTemplate excelTemplate)
        {
            return await Task.Run(() => ConvertToBytes(excelTemplate));
        }

        private byte[] ConvertToBytes(ExcelTemplate excelTemplate)
        {
            IWorkbook workbook;

            switch (excelTemplate.ExcelType)
            {
                case ExcelTemplateType.Xsl:
                    workbook = new HSSFWorkbook();
                    break;
                case ExcelTemplateType.Xslx:
                default:
                    workbook = new XSSFWorkbook();
                    break;
            }

            //Create the styles in Excel
            var _currencyFormat = workbook.CreateDataFormat().GetFormat(excelTemplate.CurrencyFormat);
            var _dateFormat = workbook.CreateDataFormat().GetFormat(excelTemplate.DateFormat);

            _currencyStyle = workbook.CreateCellStyle();
            _dateStyle = workbook.CreateCellStyle();

            _currencyStyle.DataFormat = _currencyFormat;
            _dateStyle.DataFormat = _dateFormat;

            ISheet excelSheet = workbook.CreateSheet(excelTemplate.Name);
            int rowCount = 0;
            
            //Create each row and each supplied column
            foreach (var rowTemplate in excelTemplate.Rows)
            {
                int columnCount = 0;
                var row = excelSheet.CreateRow(rowCount++);

                if (rowTemplate.Columns != null)
                {
                    foreach (var templateColumn in rowTemplate.Columns)
                    {
                        var column = row.CreateCell(columnCount++);
                        ApplyValue(column, templateColumn.Type, templateColumn.Value);
                    }
                }
            }

            //Autosize columns
            var columnsNumber = excelTemplate.Rows.Where(r => r.Columns != null).Max(r => r.Columns.Count);
            for (int i = 0; i <= columnsNumber; i++)
            {
                excelSheet.AutoSizeColumn(i);

                var width = excelSheet.GetColumnWidth(i);
                width += (width / 14) ;
                excelSheet.SetColumnWidth(i,width);
            }

            //Return workbook as byte
            byte[] result = null;
            using (MemoryStream ms = new MemoryStream())
            {
                workbook.Write(ms);
                result = ms.ToArray();
            }

            return result;
        }

        private void ApplyValue(ICell cell, ExcelTemplateDataType dataType, string value)
        {
            try
            {
                switch (dataType)
                {
                    case ExcelTemplateDataType.Date:
                        DateTime date;
                        if (DateTime.TryParse(value, out date))
                        {
                            cell.CellStyle = _dateStyle;
                            cell.SetCellValue(date);
                        }
                        else
                            goto default;
                        break;
                    case ExcelTemplateDataType.Money:
                        double money;
                        if (double.TryParse(value, out money))
                        {
                            cell.CellStyle = _currencyStyle;
                            cell.SetCellValue(money);
                        }
                        else
                            goto default;
                        break;
                    case ExcelTemplateDataType.Numeric:
                        double numeric;
                        if (double.TryParse(value, out numeric))
                        {
                            cell.SetCellValue(numeric);
                            cell.SetCellType(CellType.Numeric);
                        }
                        else
                            goto default;
                        break;
                    case ExcelTemplateDataType.String:
                    default:
                        cell.SetCellValue(value);
                        break;
                }
            }
            catch (Exception ex)
            {
                cell.SetCellValue(value);
            }

        }

        
    }
}
