using System;
using System.IO;
using Axerrio.Utilities;
using Axerrio.Utilities.Excel;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using Newtonsoft.Json;

namespace Axerio.Utilities.Excel.Test
{
    [TestClass]
    public class ExcelTest
    {
        private readonly string _testJson = @"{
    'excelType': 'Xsl',
    'name':'Prijslijstnaam',
    'currencyFormat': '$#.00',
    'dateFormat': 'yyyy-MM-dd HH:mm:ss',
    'rows': [
        {
        'columns': [   
                        {'value':'Product'},
                        {'value':'Omschrijving'},
                        {'value':'Datum'},
                        {'value':'Prijs'},
                        {'value':'Verpakking'},
                        {'value':'Aantal'}
                    ]
        },
        {},        
        {
            'columns': [
                {
                    'value': 'Dit is een minder lange'
                },
                {                    
                    'value': 'Dit is een hele lange tekst heel erg lang heeeel erg lang lang lang lang lang lang la la la land'
                },
                {
                    'value': '2016-12-10 17:00:00',
                    'type': 'Date'
                },
                {
                    'value': '1204,44',
                    'type': 'Money'
                },
                {
                    'value': '12 x 40'
                },
                {
                    'value': '40',
                    'type': 'Numeric'
                }
            ]
        }  ,        
        {
            'columns': [
                {
                    'value': 'Dit is een minder lange asdaasd '
                },
                {                    
                    'value': 'Dit is een hele lange tekst heel erg lang heeeel erg lang lang lang lang lang lang la la la land asdasd asd asd asd asd '
                },
                {
                    'value': '2016-12-11 17:00:00',
                    'type': 'Date'
                },
                {
                    'value': '1224,44',
                    'type': 'Money'
                },
                {
                    'value': '10 x 40'
                },
                {
                    'value': '50',
                    'type': 'Numeric'
                }
            ]
        }     
    ]
}";


        [TestMethod]
        public void TestMethod_Test()
        {
            var excelTemplate = JsonConvert.DeserializeObject<ExcelTemplate>(_testJson);

            IExcelExport jsonToExcel = new ExcelExport();
            var byteTask = jsonToExcel.ConvertJsonToExcelAsync(excelTemplate);
            byteTask.Wait();

            using (MemoryStream memStream = new MemoryStream(byteTask.Result))
            using (FileStream stream = new FileStream(@"C:\Temp\Test.xls", FileMode.Create, FileAccess.Write))
            {
                memStream.WriteTo(stream);
            }
        }
    }
}
