using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Axerrio.Utilities.Excel
{
    public interface IExcelExport
    {
        Task<byte[]> ConvertJsonToExcelAsync(ExcelTemplate excelTemplate);
    }
}
