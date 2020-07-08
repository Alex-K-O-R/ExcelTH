using GemBox.Spreadsheet;

namespace TableHandlers
{
    public static class CellFormatsGembox
    {
        public static void StandartInlineCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            var borders = x.Style.Borders;
            x.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromArgb(105, 105, 105), LineStyle.Thin);
            //borders[XlBordersIndex.xlInsideHorizontal].Color = XlRgbColor.rgbDarkSlateGrey;
            //borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;
            //borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            x.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            x.Style.Font.Weight = ExcelFont.NormalWeight;
            x.Style.Font.Name = "Calibri";
            x.Style.Font.Size = 201;
            x.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
            
            //x.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
            //IF 1 then Value must be between ExcelFont.MinWeight and ExcelFont.MaxWeight.
           
        }

        public static void StandartInlineRecordNumberCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            x.Style.HorizontalAlignment = HorizontalAlignmentStyle.Center;
            x.Style.Font.Color = SpreadsheetColor.FromArgb(105, 105, 105);
        }

        public static void StandartInlineNumberCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            x.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
            x.Style.Font.Weight = ExcelFont.NormalWeight;
            x.Style.Font.Size = 200;
            x.Style.Font.Name = "Calibri";
            x.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
            x.Style.NumberFormat = "@";
        }


        public static void StandartHeaderGroupCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            x.Style.Borders.SetBorders(MultipleBorders.All, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
            x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(218, 227, 250), SpreadsheetColor.FromName(ColorName.LightBlue));
            x.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
            x.Style.Font.Weight = ExcelFont.MaxWeight;
            x.Style.Font.Name = "Calibri";
            x.Style.Font.Size = 220;
            x.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
        }




        public static void ReportDataOfElectricityMeterConsumption_tariffz(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            if (dataRow != null)
            {
                int mean_index = 10;
                if (dataRow.Length > mean_index && (string)dataRow[mean_index] == "Сумма")
                    x.Offset(mean_index).Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromName(ColorName.Orange), SpreadsheetColor.FromName(ColorName.Orange));
            }
        }








        public static class Demo3
        {
            public static void GeneralHeaderFormat(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                x.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                x.Style.Font.Weight = ExcelFont.BoldWeight;
                x.Style.Font.Name = "Calibri";
                x.Style.Font.Size = 212;
                x.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
            }
            public static void HeaderCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(112, 48, 160), SpreadsheetColor.FromArgb(112, 48, 160));
            }

            public static void HeaderGroup1Cell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(218, 227, 250), SpreadsheetColor.FromArgb(218, 227, 250));
            }

            public static void HeaderGroup2Cell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(146, 208, 80), SpreadsheetColor.FromArgb(146, 208, 80));
            }

            public static void HeaderGroup3Cell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(0, 112, 192), SpreadsheetColor.FromArgb(0, 112, 192));
            }

            public static void StandartInlineCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                x.Style.Font.Weight = ExcelFont.NormalWeight;
                x.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                x.Style.Font.Name = "Calibri";
                x.Style.Font.Size = 212;
                x.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
            }
       
        }



        



        public static class Demo1
        {
            public static void SCHeaderCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.Borders.SetBorders(MultipleBorders.Outside, SpreadsheetColor.FromName(ColorName.Black), LineStyle.Thin);
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(81, 118, 232), SpreadsheetColor.FromArgb(81, 118, 232));
                x.Style.VerticalAlignment = VerticalAlignmentStyle.Top;
                x.Style.Font.Weight = ExcelFont.MaxWeight;
                x.Style.Font.Name = "Arial";
                x.Style.Font.Size = 201;
                x.Style.Font.Color = SpreadsheetColor.FromName(ColorName.Black);
            }
            public static void SurvivabilityGroupCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(3, 163, 5), SpreadsheetColor.FromArgb(3, 163, 5));
            }

            public static void PowerGroupCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(158, 49, 223), SpreadsheetColor.FromArgb(158, 49, 223));
            }

            public static void StrengthGroupCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(107, 127, 188), SpreadsheetColor.FromArgb(107, 127, 188));
            }

            public static void CostGroupCell(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(255, 148, 57), SpreadsheetColor.FromArgb(255, 148, 57));
            }

            public static void EmptyWeaponStyle(GemboxTableHandler.Table.Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                if (dataRow != null)
                {
                    //Link to description COLUMN INDEX ["Description" is 8th]; TODO: fast way to link by column name
                    int mean_index = 8;
                    if (dataRow.Length > mean_index && dataRow[mean_index] == null)
                        x.Offset(mean_index).Resize(4).Style.FillPattern.SetPattern(FillPatternStyle.Solid, SpreadsheetColor.FromArgb(255, 188, 192), SpreadsheetColor.FromArgb(255, 188, 192)); 
                    //paint next 4 columns
                }
            }
        }
        /**/

    }
}
