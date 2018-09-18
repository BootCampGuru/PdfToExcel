using System;
using System.Data;
using System.IO;
using Excel;
using SampleConvert.Helper;
using SautinSoft;

namespace Sample
{
    class Sample
    {
        static void Main(string[] args)
        {
            string pathToPdf = @"c:\pdf\test.pdf";
            string pathToExcel = Path.ChangeExtension(pathToPdf, ".xls");

            // Convert PDF file to Excel file
            SautinSoft.PdfFocus f = new SautinSoft.PdfFocus();
            
	    	// 'true' = Convert all data to spreadsheet (tabular and even textual).
            // 'false' = Skip textual data and convert only tabular (tables) data.
            f.ExcelOptions.ConvertNonTabularDataToSpreadsheet = true;

            // 'true'  = Preserve original page layout.
            // 'false' = Place tables before text.
            f.ExcelOptions.PreservePageLayout = true;

            f.OpenPdf(pathToPdf);

            if (f.PageCount > 0)
            {
                int result = f.ToExcel(pathToExcel);
                
                //Open a produced Excel workbook
                if (result==0)
                {
                  //  System.Diagnostics.Process.Start(pathToExcel);
                }
            }

            //  ConvertToText(@"C:\pdf\test.xls");
            var excelFilePath = @"C:\pdf\test.xls";
            string output = Path.ChangeExtension(excelFilePath, ".csv");
            ExcelFileHelper.SaveAsCsv(@"C:\pdf\test.xls", output);
        }

        private static void ConvertToText(string filePath)
        {
            FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);

            // Reading from a binary Excel file ('97-2003 format; *.xls)
            //IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader(stream);

            // Reading from a OpenXml Excel file (2007 format; *.xlsx)
             IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);

            // DataSet - The result of each spreadsheet will be created in the result.Tables
            DataSet result = excelReader.AsDataSet();

            // Free resources (IExcelDataReader is IDisposable)
            excelReader.Close();
        }
    }
}
