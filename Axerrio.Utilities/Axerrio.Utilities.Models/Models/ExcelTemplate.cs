using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Axerrio.Utilities
{
    public class ExcelTemplate
    {
        public ExcelTemplateType ExcelType { get; set; }
        public string Name{ get; set; }
        public string CurrencyFormat { get; set; }
        public string DateFormat { get; set; }
        public List<ExcelTemplateRow> Rows { get; set; }
    }
}
