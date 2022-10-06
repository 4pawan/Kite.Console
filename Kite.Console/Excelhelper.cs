using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using NPOI.HSSF.Util;
using NPOI.SS.UserModel;
using NPOI.SS.Util;
using NPOI.XSSF.UserModel;

namespace Zerodha.Excel
{
    public class Excelhelper
    {
        private static string dateFormat = "dd-MM-yyyy:hh:mm";
        private static string VolThreshold = "3000000";
        private static string lowToHightThreshold = "10";
        private static string lowToHightThresholdCent = "2";

        public static void ExportToExcel(string key)
        {
            string json = ReadJson();
            List<Candles> candleList = FormatJsonToObject(json);

            DataTable dt = ObjectToDataTable(candleList.OrderByDescending(c => c.Date).ToList());
            dt.Columns.Remove("Date");
            CreateExcel(dt);
        }

        static string ReadJson()
        {
            string path = @"C:\\Users\\Pawan\\source\\repos\\Kite.Console\\Kite.Console\\input\\M.json";
            return File.ReadAllText(path);
        }

        static void CreateExcel(DataTable table)
        {
            using (var fs = new FileStream("Result.xlsx", FileMode.Create, FileAccess.Write))
            {
                IWorkbook workbook = new XSSFWorkbook();
                ISheet excelSheet = workbook.CreateSheet("Sheet1");
                excelSheet.CreateFreezePane(0, 1);
                int rowCount = table.Rows.Count;
                //Show high Volume activity
                XSSFSheetConditionalFormatting condFormating = (XSSFSheetConditionalFormatting)excelSheet.SheetConditionalFormatting;
                XSSFConditionalFormattingRule cfVolRed =
                    (XSSFConditionalFormattingRule)condFormating.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, VolThreshold);

                XSSFPatternFormatting fillRed = (XSSFPatternFormatting)cfVolRed.CreatePatternFormatting();
                fillRed.FillBackgroundColor = IndexedColors.Red.Index;
                fillRed.FillPattern = FillPattern.SolidForeground;
                CellRangeAddress[] cfRangeVol = { CellRangeAddress.ValueOf($"F2:F{rowCount}") };
                condFormating.AddConditionalFormatting(cfRangeVol, cfVolRed);

                //Show Volatility
                XSSFConditionalFormattingRule cfVolatilityYellow =
                    (XSSFConditionalFormattingRule)condFormating.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, lowToHightThresholdCent);
                XSSFPatternFormatting fillyellow = (XSSFPatternFormatting)cfVolatilityYellow.CreatePatternFormatting();
                fillyellow.FillBackgroundColor = IndexedColors.LightOrange.Index;
                fillyellow.FillPattern = FillPattern.SolidForeground;
                CellRangeAddress[] cfRangelowToHighInCent = { CellRangeAddress.ValueOf($"M2:M{rowCount}") };
                condFormating.AddConditionalFormatting(cfRangelowToHighInCent, cfVolatilityYellow);

                XSSFConditionalFormattingRule cfVolatilitylight =
                    (XSSFConditionalFormattingRule)condFormating.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, lowToHightThreshold);
                XSSFPatternFormatting fill_light = (XSSFPatternFormatting)cfVolatilitylight.CreatePatternFormatting();
                fill_light.FillBackgroundColor = IndexedColors.Aqua.Index;
                fill_light.FillPattern = FillPattern.SolidForeground;
                CellRangeAddress[] cfRangelightlowToHighInCent = { CellRangeAddress.ValueOf($"G2:G{rowCount}") };
                condFormating.AddConditionalFormatting(cfRangelightlowToHighInCent, cfVolatilitylight);

                // show Gap
                XSSFConditionalFormattingRule cfGapPositive =
                    (XSSFConditionalFormattingRule)condFormating.CreateConditionalFormattingRule(ComparisonOperator.GreaterThanOrEqual, "4");
                XSSFPatternFormatting fill_gap = (XSSFPatternFormatting)cfGapPositive.CreatePatternFormatting();
                short colorGap = IndexedColors.LightTurquoise.Index;
                fill_gap.FillBackgroundColor = colorGap;
                fill_gap.FillPattern = FillPattern.SolidForeground;

                XSSFConditionalFormattingRule cfGapNegative =
                    (XSSFConditionalFormattingRule)condFormating.CreateConditionalFormattingRule(ComparisonOperator.LessThanOrEqual, "-4");
                XSSFPatternFormatting fill_gap2 = (XSSFPatternFormatting)cfGapNegative.CreatePatternFormatting();
                fill_gap2.FillBackgroundColor = colorGap;
                fill_gap2.FillPattern = FillPattern.SolidForeground;
                CellRangeAddress[] cfRangeGap = { CellRangeAddress.ValueOf($"N2:N{rowCount}") };
                condFormating.AddConditionalFormatting(cfRangeGap, cfGapPositive, cfGapNegative);

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
                            DateTime date = DateTime.ParseExact(dsrow[col].ToString(), dateFormat, CultureInfo.InvariantCulture);
                            var cell = row.CreateCell(cellIndex);
                            if (IsMonday(date))
                            {
                                ICellStyle backGroundColorStyle = workbook.CreateCellStyle();
                                short colorBlue = HSSFColor.Grey25Percent.Index;
                                backGroundColorStyle.FillForegroundColor = colorBlue;
                                backGroundColorStyle.FillPattern = FillPattern.SolidForeground;
                                cell.CellStyle = backGroundColorStyle;
                            }
                            cell.SetCellValue(dsrow[col].ToString());
                        }
                        else
                        {
                            row.CreateCell(cellIndex).SetCellValue(Convert.ToString(dsrow[col]));
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
                candle.DateFormated = _date.ToString(dateFormat);
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
