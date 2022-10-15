using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Kite.Console;
using Newtonsoft.Json;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Zerodha.Excel
{
    public class Excelhelper
    {
        public static void ExportToExcel()
        {
            string json = ReadJson();
            List<Candles> candleList = FormatJsonToObject(json);

            DataTable dt = ObjectToDataTable(candleList.OrderByDescending(c => c.Date).ToList());
            dt.Columns.Remove("Date");
            CreateExcel(dt);
        }

        static string ReadJson()
        {
            return File.ReadAllText(Constant.PATH.ExcelPath);
        }

        public static void CreateExcel(DataTable table)
        {
            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");
                excelSheet.CreateFreezePane(0, 1);
                int rowCount = table.Rows.Count;

                List<String> columns = new List<string>();
                IRow row = excelSheet.CreateRow(0);
                int columnIndex = 0;

                foreach (System.Data.DataColumn column in table.Columns)
                {
                    if (column.ColumnName == "IsLowerTailLarger")
                        continue;

                    columns.Add(column.ColumnName);
                    row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                    columnIndex++;
                }
                int rowIndex = 1;
                foreach (DataRow dsrow in table.Rows)
                {
                    row = excelSheet.CreateRow(rowIndex);
                    int cellIndex = 0;
                    foreach (string col in columns)
                    {
                        if (cellIndex == 0)
                        {
                            DateTime date = DateTime.ParseExact(dsrow[col].ToString(), Constant.DateFormat, CultureInfo.InvariantCulture);
                            var cell = row.CreateCell(cellIndex);
                            if (IsMonday(date))
                            {
                                ICellStyle backGroundColorStyle = workbook.CreateCellStyle();
                                short colorBlue = HSSFColor.Grey25Percent.Index;
                                backGroundColorStyle.FillForegroundColor = colorBlue;
                                backGroundColorStyle.FillPattern = FillPattern.SolidForeground;
                                cell.CellStyle = backGroundColorStyle;
                            }
                            if (dsrow[23].ToString() == "1")
                            {
                                ICellStyle backGroundColorStyle = workbook.CreateCellStyle();
                                short colorBlue = HSSFColor.LightGreen.Index;
                                backGroundColorStyle.FillForegroundColor = colorBlue;
                                backGroundColorStyle.FillPattern = FillPattern.SolidForeground;
                                cell.CellStyle = backGroundColorStyle;
                            }

                            cell.SetCellValue(dsrow[col].ToString());
                        }
                        else if (cellIndex == 21 || cellIndex == 22)
                        {
                            row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                        }
                        else
                        {
                            double.TryParse(Convert.ToString(dsrow[col]), NumberStyles.Number, CultureInfo.InvariantCulture, out double rowVal);
                            row.CreateCell(cellIndex).SetCellValue(rowVal);
                        }
                        cellIndex++;
                    }

                    rowIndex++;
                }
                workbook.Write(fs);
            }

        }

        static List<Candles> FormatJsonToObject(string json)
        {
            var data = JsonConvert.DeserializeObject<Response>(json);
            List<Candles> candleList = new List<Candles>();

            // get all formating done with all calculations
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
                candle.HighFrmY = High - PrevDayClose;
                candle.LowFrmY = Low - PrevDayClose;
                candle.CloseFrmY = Close - PrevDayClose;
                candle.CentHighFrmY = ((High - PrevDayClose) / PrevDayClose) * 100;
                candle.CentLowFrmY = ((Low - PrevDayClose) / PrevDayClose) * 100;
                candle.CentCloseFrmY = ((Close - PrevDayClose) / PrevDayClose) * 100;
                candle.DayCentLowToHigh = (DayLowToHigh / Low) * 100;
                candleList.Add(candle);
            }

            return candleList;
        }

        static DataTable ObjectToDataTable(List<Candles> candleList)
        {
            DataTable table = (DataTable)JsonConvert.DeserializeObject(JsonConvert.SerializeObject(candleList), (typeof(DataTable)));
            return table;
        }
        static void SetTailProperty(Candles candle)
        {

        }

        static bool IsMonday(DateTime date)
        {
            return date.DayOfWeek == DayOfWeek.Monday;
        }
    }
}
