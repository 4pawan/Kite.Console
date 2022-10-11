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

                if (!dayEntries.Any())
                    continue;

                if (dsrow[21].ToString() == "0" || string.IsNullOrEmpty(dsrow[21].ToString()))
                {
                    var max = dayEntries.First(r => r.High == dayEntries.Max(r => r.High));
                    var min = dayEntries.First(r => r.Low == dayEntries.Min(r => r.Low));

                    var _10AM = dayEntries.FirstOrDefault(r => r.Date.ToShortTimeString() == "10:00 AM");
                    var _10_30AM = dayEntries.FirstOrDefault(r => r.Date.ToShortTimeString() == "10:30 AM");
                    var _1PM = dayEntries.FirstOrDefault(r => r.Date.ToShortTimeString() == "01:00 PM");
                    var _2PM = dayEntries.FirstOrDefault(r => r.Date.ToShortTimeString() == "02:00 PM");
                    var _2_25PM = dayEntries.FirstOrDefault(r => r.Date.ToShortTimeString() == "02:25 PM");

                    if (_10AM != null)
                        dsrow[16] = _10AM.Close;
                    if (_10_30AM != null)
                        dsrow[17] = _10_30AM.Close;
                    if (_1PM != null)
                        dsrow[18] = _1PM.Close;
                    if (_2PM != null)
                        dsrow[19] = _2PM.Close;
                    if (_2_25PM != null)
                        dsrow[20] = _2_25PM.Close;

                    dsrow[21] = max.Date.ToShortTimeString();   // DayhighReachedAt
                    dsrow[22] = min.Date.ToShortTimeString();   // DaylowReachedAt

                    UpdateExpirayDateForWeek(dayEntries, dsrow);

                }
            }
            return dt;
        }

        private static void UpdateExpirayDateForWeek(IEnumerable<Candles> dayEntries, DataRow dsrow)
        {
            var lastweekDay = dayEntries.FirstOrDefault(d => d.Date.DayOfWeek == DayOfWeek.Thursday);

            if (lastweekDay != null)
            {
                dsrow[23] = 1;
            }
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
