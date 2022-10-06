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
            List<String> columns = new List<string>();
            int columnIndex = 0;
            foreach (System.Data.DataColumn column in dt.Columns)
            {
                if (column.ColumnName == "IsLowerTailLarger")
                    continue;

                columns.Add(column.ColumnName);
                columnIndex++;
            }

            int rowIndex = 1;
            foreach (DataRow dsrow in dt.Rows)
            {
                int cellIndex = 0;
                foreach (string col in columns)
                {
                    var d = dsrow[col];
                    cellIndex++;
                }
                rowIndex++;
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
                double DayLowToHigh = High - Low;
                double PrevDayClose = candleList.Any() ? candleList.Last().Close : 0;
                candle.DateFormated = _date.ToString(Constant.DateFormat);
                candle.Open = Open;
                candle.High = High;
                candle.Low = Low;
                candle.Close = Close;
                candle.Volume = long.Parse(c[5].ToString());
                candle.DayLowToHigh = DayLowToHigh;
                candle.PrevDayClose = PrevDayClose;
                candle.Gap = Open - PrevDayClose;
                candle.CentHighFrmY = ((High - PrevDayClose) / PrevDayClose) * 100;
                candle.CentLowFrmY = ((PrevDayClose - Low) / PrevDayClose) * 100;
                candle.CentCloseFrmY = ((Close - Open) / Open) * 100;
                candle.DayCentLowToHigh = (DayLowToHigh / Low) * 100;
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
