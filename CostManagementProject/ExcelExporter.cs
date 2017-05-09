using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using CostManagementProject.Models;

namespace CostManagementProject
{
    public class ExcelExportData
    {
        public IEnumerable<YearGrowth> YearGrowthData { get; set; }
        public IEnumerable<YearGrowth> YearGrowthCriteriaData { get; set; }
        public IEnumerable<YearGrowth> YearGrowthRateData { get; set; }
        public IEnumerable<YearGrowth> YearSpirmanData { get; set; }
        public IEnumerable<FehnerCompare> FehnerData { get; set; }
        public IEnumerable<CompanyScale> ScaleRateData { get; set; }

        public ExcelExportData()
        {
            YearGrowthData = new List<YearGrowth>();
            YearGrowthCriteriaData = new List<YearGrowth>();
            YearGrowthRateData = new List<YearGrowth>();
            YearSpirmanData = new List<YearGrowth>();
            FehnerData = new List<FehnerCompare>();
            ScaleRateData = new List<CompanyScale>();
        }
    }

    public static class ExcelExporter
    {
        public static void Export(ExcelExportData data, string filePath) 
        {
            object misValue = System.Reflection.Missing.Value;

            Application xlApp = new Application();
            Workbook xlWorkBook = xlApp.Workbooks.Add(misValue);
            List<Worksheet> xlWorksheets = new List<Worksheet>();

            xlWorksheets.Add(FillYearGrowthData((Worksheet)xlWorkBook.Worksheets.get_Item(1), data.YearGrowthData));
            xlWorksheets.Add(FillYearGrowthCriteriaData((Worksheet)xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]), data.YearGrowthCriteriaData));
            xlWorksheets.Add(FillYearGrowthRateData((Worksheet)xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]), data.YearGrowthRateData));
            xlWorksheets.Add(FillYearSpirmanData((Worksheet)xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]), data.YearSpirmanData));
            xlWorksheets.Add(FillFehnerData((Worksheet)xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]), data.FehnerData));
            xlWorksheets.Add(FillScaleRateData((Worksheet)xlWorkBook.Worksheets.Add(After: xlWorkBook.Sheets[xlWorkBook.Sheets.Count]), data.ScaleRateData));

            try
            {
                xlWorkBook.SaveAs(filePath, XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                xlApp.Quit();

                foreach (var sheet in xlWorksheets)
                {
                    ReleaseObject(sheet);
                }

                ReleaseObject(xlWorkBook);
                ReleaseObject(xlApp);
            }
        }

        private static Worksheet FillYearGrowthData(Worksheet sheet, IEnumerable<YearGrowth> data)
        {
            sheet.Name = "Основні індикатори";
            sheet.Columns[1].ColumnWidth = 50;

            sheet.Cells[1, 1] = @"Рік";
            sheet.Cells[2, 1] = @"Чистий прибуток\збиток";
            sheet.Cells[3, 1] = @"Чистий дохід від реалізації";
            sheet.Cells[4, 1] = @"Собівартість";
            sheet.Cells[5, 1] = @"Середньорічна вартість активів";
            sheet.Cells[6, 1] = @"Середньорічна вартість основних засобів";
            sheet.Cells[7, 1] = @"Середньорічна вартість оборотних активів";
            sheet.Cells[8, 1] = @"Середньоспискова чисельність працівників";

            var currentIndex = 2;
            foreach (var dataItem in data)
            {
                sheet.Columns[currentIndex].ColumnWidth = 15;

                sheet.Cells[1, currentIndex] = dataItem.Year;
                sheet.Cells[2, currentIndex] = dataItem.NetProfit;
                sheet.Cells[3, currentIndex] = dataItem.SalesNetIncome;
                sheet.Cells[4, currentIndex] = dataItem.Cost;
                sheet.Cells[5, currentIndex] = dataItem.AverageAssets;
                sheet.Cells[6, currentIndex] = dataItem.AverageFixedAssets;
                sheet.Cells[7, currentIndex] = dataItem.AverageCurrentAssets;
                sheet.Cells[8, currentIndex] = dataItem.EmployeeCount;

                currentIndex++;
            }

            return sheet;
        }

        private static Worksheet FillYearGrowthCriteriaData(Worksheet sheet, IEnumerable<YearGrowth> data)
        {
            sheet.Name = "Темпи зростання (ТЗ)";
            sheet.Columns[1].ColumnWidth = 50;

            sheet.Cells[1, 1] = @"Рік";
            sheet.Cells[2, 1] = @"Чистий прибуток\збиток";
            sheet.Cells[3, 1] = @"Чистий дохід від реалізації";
            sheet.Cells[4, 1] = @"Собівартість";
            sheet.Cells[5, 1] = @"Середньорічна вартість активів";
            sheet.Cells[6, 1] = @"Середньорічна вартість основних засобів";
            sheet.Cells[7, 1] = @"Середньорічна вартість оборотних активів";
            sheet.Cells[8, 1] = @"Середньоспискова чисельність працівників";

            var currentIndex = 2;
            foreach (var dataItem in data)
            {
                sheet.Columns[currentIndex].ColumnWidth = 15;

                sheet.Cells[1, currentIndex] = dataItem.Year;
                sheet.Cells[2, currentIndex] = dataItem.NetProfit;
                sheet.Cells[3, currentIndex] = dataItem.SalesNetIncome;
                sheet.Cells[4, currentIndex] = dataItem.Cost;
                sheet.Cells[5, currentIndex] = dataItem.AverageAssets;
                sheet.Cells[6, currentIndex] = dataItem.AverageFixedAssets;
                sheet.Cells[7, currentIndex] = dataItem.AverageCurrentAssets;
                sheet.Cells[8, currentIndex] = dataItem.EmployeeCount;

                currentIndex++;
            }

            return sheet;
        }

        private static Worksheet FillYearGrowthRateData(Worksheet sheet, IEnumerable<YearGrowth> data)
        {
            sheet.Name = "Рейтингова оцінка ТЗ";
            sheet.Columns[1].ColumnWidth = 50;

            sheet.Cells[1, 1] = @"Рік";
            sheet.Cells[2, 1] = @"Чистий прибуток\збиток";
            sheet.Cells[3, 1] = @"Чистий дохід від реалізації";
            sheet.Cells[4, 1] = @"Собівартість";
            sheet.Cells[5, 1] = @"Середньорічна вартість активів";
            sheet.Cells[6, 1] = @"Середньорічна вартість основних засобів";
            sheet.Cells[7, 1] = @"Середньорічна вартість оборотних активів";
            sheet.Cells[8, 1] = @"Середньоспискова чисельність працівників";

            var currentIndex = 2;
            foreach (var dataItem in data)
            {
                sheet.Columns[currentIndex].ColumnWidth = 15;

                sheet.Cells[1, currentIndex] = dataItem.Year;
                sheet.Cells[2, currentIndex] = dataItem.NetProfit;
                sheet.Cells[3, currentIndex] = dataItem.SalesNetIncome;
                sheet.Cells[4, currentIndex] = dataItem.Cost;
                sheet.Cells[5, currentIndex] = dataItem.AverageAssets;
                sheet.Cells[6, currentIndex] = dataItem.AverageFixedAssets;
                sheet.Cells[7, currentIndex] = dataItem.AverageCurrentAssets;
                sheet.Cells[8, currentIndex] = dataItem.EmployeeCount;

                currentIndex++;
            }

            return sheet;
        }

        private static Worksheet FillYearSpirmanData(Worksheet sheet, IEnumerable<YearGrowth> data)
        {
            sheet.Name = "Кореляція Спірмана";
            sheet.Columns[1].ColumnWidth = 50;
            
            sheet.Cells[1, 1] = @"Рік";
            sheet.Cells[2, 1] = @"Чистий прибуток\збиток";
            sheet.Cells[3, 1] = @"Чистий дохід від реалізації";
            sheet.Cells[4, 1] = @"Собівартість";
            sheet.Cells[5, 1] = @"Середньорічна вартість активів";
            sheet.Cells[6, 1] = @"Середньорічна вартість основних засобів";
            sheet.Cells[7, 1] = @"Середньорічна вартість оборотних активів";
            sheet.Cells[8, 1] = @"Середньоспискова чисельність працівників";
            sheet.Cells[9, 1] = @"Коефіцієнт Спірмана";

            var currentIndex = 2;
            foreach (var dataItem in data)
            {
                sheet.Columns[currentIndex].ColumnWidth = 15;

                sheet.Cells[1, currentIndex] = dataItem.Year;
                sheet.Cells[2, currentIndex] = dataItem.NetProfit;
                sheet.Cells[3, currentIndex] = dataItem.SalesNetIncome;
                sheet.Cells[4, currentIndex] = dataItem.Cost;
                sheet.Cells[5, currentIndex] = dataItem.AverageAssets;
                sheet.Cells[6, currentIndex] = dataItem.AverageFixedAssets;
                sheet.Cells[7, currentIndex] = dataItem.AverageCurrentAssets;
                sheet.Cells[8, currentIndex] = dataItem.EmployeeCount;
                sheet.Cells[9, currentIndex] = dataItem.SpiermanCoef;

                currentIndex++;
            }

            return sheet;
        }

        private static Worksheet FillFehnerData(Worksheet sheet, IEnumerable<FehnerCompare> data)
        {
            sheet.Name = "Коефіцієнти збігів Фехнера";
            sheet.Columns[1].ColumnWidth = 75;

            sheet.Cells[1, 1] = @"Рік";
            sheet.Cells[2, 1] = @"Чистий прибуток // Середньорічна вартість оборотних активів";
            sheet.Cells[3, 1] = @"Чистий прибуток // Чистий дохід від реалізації";
            sheet.Cells[4, 1] = @"Чистий прибуток // Середньорічна вартість активів";
            sheet.Cells[5, 1] = @"Чистий прибуток // Середньорічна вартість основних активів";
            sheet.Cells[6, 1] = @"Чистий прибуток // Собівартість";
            sheet.Cells[7, 1] = @"Чистий прибуток // Середньоспискова чисельність працівників";
            sheet.Cells[8, 1] = @"Середньорічна вартість оборотних активів // Чистий дохід від реалізації";
            sheet.Cells[9, 1] = @"Середньорічна вартість оборотних активів // Середньорічна вартість активів";
            sheet.Cells[10, 1] = @"Середньорічна вартість оборотних активів // Середньорічна вартість основних засобів";
            sheet.Cells[11, 1] = @"Середньорічна вартість оборотних активів // Собівартість";
            sheet.Cells[12, 1] = @"Середньорічна вартість оборотних активів // Середньоспискова чисельність працівників";
            sheet.Cells[13, 1] = @"Чистий дохід від реалізації // Середньорічна вартість активів";
            sheet.Cells[14, 1] = @"Чистий дохід від реалізації // Середньорічна вартість основних засобів";
            sheet.Cells[15, 1] = @"Чистий дохід від реалізації // Собівартість";
            sheet.Cells[16, 1] = @"Чистий дохід від реалізації // Середньоспискова чисельність працівників";
            sheet.Cells[17, 1] = @"Середньорічна вартість активів // Середньорічна вартість основних засобів";
            sheet.Cells[18, 1] = @"Середньорічна вартість активів // Собівартість";
            sheet.Cells[19, 1] = @"Середньорічна вартість активів // Середньоспискова чисельність працівників";
            sheet.Cells[20, 1] = @"Середньорічна вартість основних засобів // Собівартість";
            sheet.Cells[21, 1] = @"Середньорічна вартість основних засобів // Середньоспискова чисельність працівників";
            sheet.Cells[22, 1] = @"Собівартість // Середньоспискова чисельність працівників";
            sheet.Cells[23, 1] = @"Сума позитивних і негативних співвідношень між рангами";
            sheet.Cells[24, 1] = @"Коефіцієнт збігів Фехнера";


            var currentIndex = 2;
            foreach (var dataItem in data)
            {
                sheet.Columns[currentIndex].ColumnWidth = 15;

                sheet.Cells[1, currentIndex] = dataItem.Year;
                sheet.Cells[2, currentIndex] = dataItem.First;
                sheet.Cells[3, currentIndex] = dataItem.Second;
                sheet.Cells[4, currentIndex] = dataItem.Third;
                sheet.Cells[5, currentIndex] = dataItem.Fourth;
                sheet.Cells[6, currentIndex] = dataItem.Fifth;
                sheet.Cells[7, currentIndex] = dataItem.Sixth;
                sheet.Cells[8, currentIndex] = dataItem.Seventh;
                sheet.Cells[9, currentIndex] = dataItem.Eighth;
                sheet.Cells[10, currentIndex] =dataItem.Ninth;
                sheet.Cells[11, currentIndex] = dataItem.Tenth;
                sheet.Cells[12, currentIndex] = dataItem.Eleventh;
                sheet.Cells[13, currentIndex] = dataItem.Twelfth;
                sheet.Cells[14, currentIndex] = dataItem.Thirteenth;
                sheet.Cells[15, currentIndex] = dataItem.Fourteenth;
                sheet.Cells[16, currentIndex] = dataItem.Fifteenth;
                sheet.Cells[17, currentIndex] = dataItem.Sixteenth;
                sheet.Cells[18, currentIndex] = dataItem.Seventeenth;
                sheet.Cells[19, currentIndex] = dataItem.Eighteenth;
                sheet.Cells[20, currentIndex] = dataItem.Nineteenth;
                sheet.Cells[21, currentIndex] = dataItem.Twentieth;
                sheet.Cells[22, currentIndex] = dataItem.TwentyFirst;
                sheet.Cells[23, currentIndex] = dataItem.RangSum;
                sheet.Cells[24, currentIndex] = dataItem.FahnerCoef;

                currentIndex++;
            }

            return sheet;
        }

        private static Worksheet FillScaleRateData(Worksheet sheet, IEnumerable<CompanyScale> data)
        {
            sheet.Name = "Рівні масштабності розвитку";
            sheet.Columns[1].ColumnWidth = 30;
            sheet.Columns[2].ColumnWidth = 30;
            sheet.Columns[3].ColumnWidth = 30;
            sheet.Columns[4].ColumnWidth = 30;

            sheet.Cells[1, 1] = @"Аналізовані періоди, роки".ToUpper();
            sheet.Cells[1, 2] = @"Коефіцієнт Спірмана".ToUpper();
            sheet.Cells[1, 3] = @"Коефіцієнт Фехнера".ToUpper();
            sheet.Cells[1, 4] = @"Рівент масштабності підприємства".ToUpper();

            var currentIndex = 2;
            foreach (var dataItem in data)
            {
                sheet.Cells[currentIndex, 1] = dataItem.Year;
                sheet.Cells[currentIndex, 2] = dataItem.SpirmanCoef;
                sheet.Cells[currentIndex, 3] = dataItem.FehnerCoef;
                sheet.Cells[currentIndex, 4] = dataItem.ScaleRate;

                currentIndex++;
            }

            return sheet;
        }

        private static void ReleaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;

                Debug.Write(ex.Message);
                if (ex.InnerException != null)
                {
                    Debug.Write(ex.InnerException.Message);
                }
            }
            finally
            {
                GC.Collect();
            }
        }
    }
}
