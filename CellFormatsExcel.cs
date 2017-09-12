using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;

namespace TableHandlers
{
    public static class CellFormatsExcel
    {
        public static void StandartInlineCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            var borders = x.Borders;
            x.BorderAround(XlLineStyle.xlContinuous, XlBorderWeight.xlThin, XlColorIndex.xlColorIndexNone, XlRgbColor.rgbDarkGray);
            borders[XlBordersIndex.xlInsideHorizontal].Color = XlRgbColor.rgbDarkSlateGrey;
            borders[XlBordersIndex.xlInsideHorizontal].Weight = XlBorderWeight.xlThin;
            borders[XlBordersIndex.xlInsideHorizontal].LineStyle = XlLineStyle.xlContinuous;
            
            x.Font.Bold = false;
            x.VerticalAlignment = XlVAlign.xlVAlignTop;
            x.Font.Size = 10;
            x.Font.Name = "Calibri";
            x.Font.Color = XlRgbColor.rgbBlack;
        }

        public static void StandartInlineRecordNumberCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            x.HorizontalAlignment = XlHAlign.xlHAlignCenter;
            x.Font.Color = XlRgbColor.rgbDimGray;
        }

        public static void StandartInlineNumberCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            x.BorderAround(XlLineStyle.xlContinuous, 2, XlColorIndex.xlColorIndexNone, XlRgbColor.rgbBlack);
            x.Font.Bold = false;
            x.Font.Size = 10;
            x.Font.Name = "Calibri";
            x.Font.Color = XlRgbColor.rgbBlack;
            x.NumberFormat = "@";
        }


        public static void StandartHeaderGroupCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
        {
            x.BorderAround(XlLineStyle.xlContinuous, 2, XlColorIndex.xlColorIndexNone, XlRgbColor.rgbBlack);
            x.Interior.Color = 0xFAE3DA;
            x.VerticalAlignment = XlVAlign.xlVAlignTop;
            x.Font.Bold = true;
            x.Font.Name = "Calibri";
            x.Font.Size = 11;
            x.Font.Color = XlRgbColor.rgbBlack;
        }


        public static class Demo3
        {
            public static void GeneralHeaderFormat(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.BorderAround(XlLineStyle.xlContinuous, 2, XlColorIndex.xlColorIndexNone, XlRgbColor.rgbBlack);
                x.VerticalAlignment = XlVAlign.xlVAlignTop;
                x.Font.Bold = true;
                x.Font.Name = "Calibri";
                x.Font.Size = 11;
                x.Font.Color = XlRgbColor.rgbBlack;
            }
            public static void HeaderCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0xA03070;
            }

            public static void HeaderGroup1Cell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0xFAE3DA;
            }

            public static void HeaderGroup2Cell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0x50D092;
            }

            public static void HeaderGroup3Cell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0xC07000;
            }

            public static void StandartInlineCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.BorderAround(XlLineStyle.xlContinuous, 2, XlColorIndex.xlColorIndexNone, XlRgbColor.rgbBlack);
                x.Font.Bold = false;
                x.VerticalAlignment = XlVAlign.xlVAlignTop;
                x.Font.Size = 10;
                x.Font.Name = "Calibri";
                x.Font.Color = XlRgbColor.rgbBlack;
            }

        }



        public static class Demo1 {
            public static void SCHeaderCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.BorderAround(XlLineStyle.xlContinuous, 2, XlColorIndex.xlColorIndexNone, XlRgbColor.rgbBlack);
                x.Interior.Color = 0xe87651;
                x.VerticalAlignment = XlVAlign.xlVAlignTop;
                x.Font.Bold = true;
                x.Font.Name = "Arial";
                x.Font.Size = 10;
                x.Font.Color = XlRgbColor.rgbBlack;
            }
            public static void SurvivabilityGroupCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0x05a303;
            }

            public static void PowerGroupCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0xdf319e;
            }

            public static void StrengthGroupCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0xbc7f6b;
            }

            public static void CostGroupCell(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                x.Interior.Color = 0x3994ff;
            }

            public static void EmptyWeaponStyle(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
            {
                if (dataRow != null)
                {
                    //Link to description COLUMN INDEX ["Description" is 8th]; TODO: fast way to link by column name
                    int mean_index = 8;
                    if (dataRow.Length > mean_index && dataRow[mean_index] == null)
                        x.Offset(0, mean_index).Resize(1, 4).Interior.Color = 0xc0bcff; //paint next 4 columns
                }
            }
        }
    }
}
