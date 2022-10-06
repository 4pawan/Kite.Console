using Newtonsoft.Json;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Zerodha.Excel;

namespace Kite.Console
{
    public class IntraDayReport
    {
        public static void GenerateReport()
        {
            DataTable result = ReadResultFromExcelFile();
            var _5minReport = Read5minReport();
            DataTable updated = AddIntraDayReportToTable(_5minReport, result);
            Excelhelper.CreateExcel(updated);
        }
        public static DataTable AddIntraDayReportToTable(List<Candles> reports, DataTable dt)
        {

            foreach (DataRow dsrow in dt.Rows)
            {
                DateTime date = DateTime.ParseExact(dsrow[0].ToString(), Constant.DateFormat, null);
                var dayEntries = reports.Where(r => r.Date.Date == date);

                var max = dayEntries.First(r => r.High == dayEntries.Max(r => r.High));
                var min = dayEntries.First(r => r.Low == dayEntries.Min(r => r.Low));                            
               
                //var _10AM = dayEntries.Where(r => r.Date.ToLocalTime() == "10");
                
                dsrow[17] = max.Date.ToShortTimeString();   // DayhighReachedAt
                dsrow[18] = min.Date.ToShortTimeString();   // DaylowReachedAt

            }
            return dt;
        }

        public static List<Candles> Read5minReport()
        {
            string json = File.ReadAllText(Constant.PATH._5minChartPath);
            var data = JsonConvert.DeserializeObject<Response>(json);
            List<Candles> candleList = new List<Candles>();

            foreach (List<object> c in data.data.candles)
            {
                var candle = new Candles();
                var _date = DateTime.Parse(Convert.ToString(c[0]));
                candle.Date = _date;
                double Open = Convert.ToDouble(c[1]);
                double High = Convert.ToDouble(c[2]);
                double Low = Convert.ToDouble(c[3]);
                double Close = Convert.ToDouble(c[4]);
                candle.DateFormated = _date.ToString(Constant.DateFormat);
                candle.Open = Open;
                candle.High = High;
                candle.Low = Low;
                candle.Close = Close;
                candle.Volume = long.Parse(c[5].ToString());
                candleList.Add(candle);
            }
            return candleList;
        }

        static DataTable ReadResultFromExcelFile()
        {
            IWorkbook workbook;
            using (var stream = new FileStream("Result.xlsx", FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(stream); // XSSFWorkbook for XLSX
            }

            var sheet = workbook.GetSheetAt(0); // zero-based index of your target sheet
            var dataTable = new DataTable(sheet.SheetName);

            // write the header row
            var headerRow = sheet.GetRow(0);
            foreach (var headerCell in headerRow)
            {
                dataTable.Columns.Add(headerCell.ToString());
            }

            // write the rest
            for (int i = 1; i < sheet.PhysicalNumberOfRows; i++)
            {
                var sheetRow = sheet.GetRow(i);
                var dtRow = dataTable.NewRow();
                dtRow.ItemArray = dataTable.Columns
                    .Cast<DataColumn>()
                    .Select(c => sheetRow.GetCell(c.Ordinal, MissingCellPolicy.CREATE_NULL_AS_BLANK).ToString())
                    .ToArray();
                dataTable.Rows.Add(dtRow);
            }
            return dataTable;
        }
    }
}
