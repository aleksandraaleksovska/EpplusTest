using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using System.IO;
using System.Linq;

namespace ConsoleApp2
{
    class Program
    {
        static void Main(string[] args)
        {            
            CreatePivot();
        }

        private static bool CreatePivot()
        {
            //open sales excel
            string salesPath = @"C:\Users\AleksandraAleksovska\Desktop\Doc\python\SalesData.xlsx";
            FileInfo fileInfo = new FileInfo(salesPath);
            ExcelPackage wrkBookSalesData = new ExcelPackage(fileInfo);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            
            //add new sheet
            wrkBookSalesData.Workbook.Worksheets.Add("Pivot");
            
            ExcelWorksheet salesSheet = wrkBookSalesData.Workbook.Worksheets["Sales"];
            ExcelWorksheet pivotSheet = wrkBookSalesData.Workbook.Worksheets["Pivot"];

            //define the data range on the source sheet
            var dataRange = salesSheet.Cells[salesSheet.Dimension.Address];

            //create the pivot table
            var pivotTable = pivotSheet.PivotTables.Add(pivotSheet.Cells["B2"], dataRange, "PivotTable");

            //label field
            var units = pivotTable.DataFields.Add(pivotTable.Fields["UNITS"]);
            units.Function = DataFieldFunctions.Sum;

            //data fields
            var field = pivotTable.DataFields.Add(pivotTable.Fields["FFA € (1.8 %)"]);
            field.Name = "Count of Column FFA";
            field.Function = DataFieldFunctions.Count;

            field = pivotTable.DataFields.Add(pivotTable.Fields["Gross €"]);
            field.Name = "Sum of Column Gross";
            field.Function = DataFieldFunctions.Sum;
            field.Format = "0.00";            

            field = pivotTable.DataFields.Add(pivotTable.Fields["Carrier"]);
            field.Name = "Sum of Column Carrier";
            field.Function = DataFieldFunctions.Sum;
            field.Format = "€#,##0.00";


            wrkBookSalesData.Save();


            return true;
        }
    }
}
