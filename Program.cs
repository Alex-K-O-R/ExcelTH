using GemBox.Spreadsheet;
using NetOffice.ExcelApi;
using NetOffice.ExcelApi.Enums;
using System;
using System.Linq;
using TableHandlers;
using System.IO;

namespace Main
{
    class Program
    {
        public class Logger : TableHandlers.Logger
        {
            private string filePathAndName = String.Empty;

            public Logger(string filePathAndName = null) : base()
            {
                if (filePathAndName != null) this.filePathAndName = filePathAndName+".txt";
            }
            public override void Write(string txt)
            {
                if (filePathAndName != String.Empty) { 
                    if (File.GetLastWriteTime(filePathAndName) < DateTime.Now.AddDays(-2) || new FileInfo(filePathAndName).Length > 1024*8192){
                        System.IO.File.WriteAllText(filePathAndName, string.Empty);
                    }
                    File.AppendAllText(filePathAndName, System.Environment.NewLine+DateTime.Now.ToString()+": "+txt);
                }
            }
        }


        private static string Get_Full_Output_Path(string fileName) { return Path.GetPathRoot(Environment.SystemDirectory) + @"Reports_example\"+fileName; }
        static void Main(string[] args)
        {
            Console.WriteLine("This is a demo of ExcelTH project.");
            Console.WriteLine(@"Two type of files will be created in your SYSTEM_DISK:\Reports_example\ directory.");
            Console.WriteLine(@"xl*.xlsx is a table made up with help of NetOffice lib.");
            Console.WriteLine(@"gb*.xlsx is a table made up with help of Gembox lib.");
            for (int i = 0; i < 20; i++) Console.Write("-");
            Console.WriteLine("");
            Console.WriteLine("OPTIONS:");
            Console.WriteLine("(1) Create headers demo using NetOffice lib");
            Console.WriteLine("(2) Create headers demo using Gembox lib");
            Console.WriteLine("(3) Create rows demo using NetOffice lib");
            Console.WriteLine("(4) Create rows demo using Gembox lib");

            var a = Console.ReadLine();

            if (a == "1" || a == "2")
            {
                var generated_data_as_2_dimensional_object_array = TableHandlers.ReportData.LoadDataDemo1();
                if (a == "1") GenerateDemoUsingNetOffice(generated_data_as_2_dimensional_object_array);
                if (a == "2") GenerateDemoUsingGembox(generated_data_as_2_dimensional_object_array);
            }

            if (a == "3" || a == "4")
            {
                var generated_data_as_2_dimensional_object_array = TableHandlers.ReportData.LoadDataDemo1();
                if (a == "3") GenerateDemoUsingNetOffice(generated_data_as_2_dimensional_object_array, 1);
                if (a == "4") GenerateDemoUsingGembox(generated_data_as_2_dimensional_object_array, 1);
            }
            
            Console.ReadKey();
        }


        public static void GenerateDemoUsingNetOffice(object[][] data, int demoNum = 0 , string fileName = "xl12345.xlsx")
        {
            string worksheetName = "Demo NetOffice";

            Console.WriteLine("Building excel report using NetOffice.");
            Console.WriteLine("Launching background Excel process.");
            var excelType = Type.GetTypeFromCLSID(new Guid(@"{00020812-0000-0000-C000-000000000046}"));
            var excel = Activator.CreateInstance(excelType);
            var xlApp = new Application(new NetOffice.COMObject(excel));

            xlApp.Application.DisplayAlerts = false;

            xlApp.Workbooks.Add();
            (xlApp.Workbooks.Last() as Workbook).Activate();
            var wb = xlApp.ActiveWorkbook;

            Console.WriteLine("Workbook created, creating worksheet.");
            wb.Worksheets.Add();
            var WORK_SHEET = (wb.Worksheets.Last() as Worksheet);
            if (WORK_SHEET != null)
            {
                for (int i = 0; i <= wb.Worksheets.Count; i++)
                {
                    (wb.Worksheets.ElementAt(0) as Worksheet).Delete();
                }

                WORK_SHEET.Name = worksheetName;

                Console.WriteLine("Worksheet's ready. Building report.");

                if (demoNum == 0) CreateDemoHeadersTableWithNetOffice(ref WORK_SHEET);
                if (demoNum == 1)
                {
                    CreateSCTableWithNetOffice(ref WORK_SHEET);
                    WORK_SHEET.UsedRange.Columns.AutoFit();
                }
            }

            Console.WriteLine("Report filled. Saving and closing background excel process.");
            wb.SaveAs(Get_Full_Output_Path(fileName), XlFileFormat.xlOpenXMLWorkbook, null,
                null, false, false, XlSaveAsAccessMode.xlNoChange,
                XlSaveConflictResolution.xlUserResolution, true,
                null, null, null);
            wb.Close();
            xlApp.Application.Quit();

            //BRUTEFORCE
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0) PK.Kill();
            }

            Console.WriteLine("Job's done. Press any key.");
        }



        private static void CreateSCTableWithNetOffice(ref Worksheet WORK_SHEET)
        {
            var table = new TableHandlers.ExcelTableHandler.Table(new Logger(Get_Full_Output_Path("NetExcel table_demo log")));

            //Set specific fill for headers [using .Formatter; priority 1 is higher than 0]
            var styleH = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.SCHeaderCell, 0);

            //Set specific fill for rows [using .Formatter; priority 1 is higher than 0]
            var styleC = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.StandartInlineCell, 0);
            var styleEmptyWeaponOutline = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.EmptyWeaponStyle, 1);

            //Some extra-feature that hopefully will be modified for good. It's usable for now, but lacks of modification-proof
            //Specific formats for the part of entire row -- goes AFTER all cells in specific row are formatted
            table.addRowFormat(styleEmptyWeaponOutline);

            //Link new columns to array of data by indexes [addColumn(TITLE, DATA_ARRAY_INDEX)], add styling
            table.addColumn("#num [SRC]", 10).addFormat(styleH)
                .addCellFormat(new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.StandartInlineCell, 0))
                    .addCellFormat(new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.StandartInlineRecordNumberCell, 1));
            //  VS
            table.addColumn("[NextEmpty]", 10).addFormat(styleH)
                .addCellFormat(new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.StandartInlineCell, 0))
                    .addCellFormat(new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.StandartInlineRecordNumberCell, 1))
                        .EnableAutoMergeNextEmptyValues();
            //Enable merging modes for column: 
            //    -treat next same values in data array column as one and merge them;
            //    -treat all next empty values that do follow after non-empty value in data array column as last non-empty value and merge them;
            //    -break ALL COLUMNS merging rules when value is changed [provides infoblocks of multirow data]
            table.addColumn("Info [SRC]", 14).addFormat(styleH).addCellFormat(styleC);
            table.addColumn("[MergeSame]", 14).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();

            table.addColumn("Unit [SRC]", 0).addFormat(styleH).addCellFormat(styleC);
            table.addColumn("[NE + BlockDelimeter]", 0).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeNextEmptyValues().UseRowDataAsInfoBlockDelimeter();

            table.addColumn("Health Points", 1).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Shield Points", 2).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Description", 3).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();
            table.addColumn("Damage", 4).addCellFormat(styleC).addFormat(styleH).EnableAutoMergeSameValues();
            table.addColumn("Range", 5).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();
            table.addColumn("Cooldown", 6).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();
            table.addColumn("Minerals", 7).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Vespene", 8).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Resources", 9).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();

            //Set specific fill for different groups [using .Formatter; priority 1 is higher than 0]
            var styleSVA = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.SurvivabilityGroupCell, 1);
            var stylePwr = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.PowerGroupCell, 1);
            var styleCst = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.CostGroupCell, 1);
            var styleSgh = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.StrengthGroupCell, 1);
            var styleGeneralHeaderGroup = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo1.SCHeaderCell, 0);

            //Adding groups of headers
            table.addHeaderGroup("Strength", "Health Points", "Shield Points",
                        "Description", "Damage", "Range", "Cooldown").addFormat(styleSgh).addFormat(styleGeneralHeaderGroup);
            table.addHeaderGroup("Survivability", "Health Points", "Shield Points").addFormat(styleSVA).addFormat(styleGeneralHeaderGroup);
            table.addHeaderGroup("Powers", "Description", "Damage", "Range", "Cooldown").addFormat(stylePwr).addFormat(styleGeneralHeaderGroup);
            table.addHeaderGroup("Cost", "Minerals", "Vespene", "Resources").addFormat(styleCst).addFormat(styleGeneralHeaderGroup);



            WORK_SHEET.Cells[2, 2].Value = "Row cells have 3 settings available that can be accessed either by methods or by settings of a specific column.";
            WORK_SHEET.Range(WORK_SHEET.Cells[2, 2], WORK_SHEET.Cells[2, 14]).Merge();
            WORK_SHEET.Cells[3, 2].Value = "Those are: AutoMergeSameValues = [bool]; AutoMergeNextEmptyValues = [bool]; IsBlockDelimeter = [bool]";
            WORK_SHEET.Range(WORK_SHEET.Cells[3, 2], WORK_SHEET.Cells[3, 14]).Merge();
            WORK_SHEET.Cells[4, 2].Value = "Purpose: to control merge of column's cells content.";
            WORK_SHEET.Range(WORK_SHEET.Cells[4, 2], WORK_SHEET.Cells[4, 14]).Merge();
            WORK_SHEET.Cells[6, 2].Value = "While AutoMergeSameValues and AutoMergeNextEmptyValues are obvious, IsBlockDelimeter is not.";
            WORK_SHEET.Range(WORK_SHEET.Cells[6, 2], WORK_SHEET.Cells[6, 14]).Merge();
            WORK_SHEET.Cells[7, 2].Value = "IsBlockDelimeter is used to prevent merging any cells in current row with a previous row that is logically related to another information group (keyValue=>data / Infoblock structure).";
            WORK_SHEET.Range(WORK_SHEET.Cells[7, 2], WORK_SHEET.Cells[7, 18]).Merge();
            WORK_SHEET.Cells[8, 2].Value = "For better understanding, see table below. It has columns with source data marked with [SRC] key and ongoing column with some setting is enabled.";
            WORK_SHEET.Range(WORK_SHEET.Cells[8, 2], WORK_SHEET.Cells[8, 14]).Merge();
            //Specifying margins: coordinates of the left upper row [new Pivot(RowNumber, ColNumber)]
            table.SetPivot(new TableHandlers.ExcelTableHandler.Table.Pivot(9, 2));

            var data = ReportData.LoadDataDemo1();
            table.Draw(ref WORK_SHEET, ref data);
        }




        private static void CreateDemoHeadersTableWithNetOffice(ref Worksheet WORK_SHEET)
        {
            var table = new TableHandlers.ExcelTableHandler.Table(new Logger(Get_Full_Output_Path("NetExcel header_demo log")));

            //Setting global header style for text
            var HGeneral = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo3.GeneralHeaderFormat, 0);

            //Setting specific fill for different groups and total [using .Formatter priority 1, higher than HGeneral 0]
            var HCell = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo3.HeaderCell, 1);
            var HGr1 = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo3.HeaderGroup1Cell, 1);
            var HGr2 = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo3.HeaderGroup2Cell, 1);
            var HGr3 = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo3.HeaderGroup3Cell, 1);

            var Cell = new TableHandlers.ExcelTableHandler.Table.Formatter(CellFormatsExcel.Demo3.StandartInlineCell, 0);

            table.addColumn("HD1", 0).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues().UseRowDataAsInfoBlockDelimeter();
            table.addColumn("HD2", 1).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("HD3", 2).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("HD4", 3).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues();
                                                                                        
            table.addHeaderGroup("GR1", "HD1").addFormat(HGeneral).addFormat(HGr1);
            table.addHeaderGroup("GR2", "HD1", "HD3", "HD4").addFormat(HGeneral).addFormat(HGr2);
            table.addHeaderGroup("GR2.5", "HD3").addFormat(HGeneral).addFormat(HGr1);
            table.addHeaderGroup("GR3", "HD4").addFormat(HGeneral).addFormat(HGr3);

            var data = ReportData.LoadDataDemo3();
            WORK_SHEET.Cells[2, 2].Value = "Headers have 5 settings accessed via TABLE().Settings";
            WORK_SHEET.Cells[3, 2].Value = "Settings // FixFirstNColumns = [int amount] and FixHeaderRows = [bool]";
            WORK_SHEET.Cells[4, 2].Value = "Those are used to freeze some of vertical or horizontal areas of worksheet";
            table.Settings.FixFirstNColumns = 1;
            table.Settings.FixHeaderRows = true;
            table.Settings.UseSmartResetForFixedColsAndRows = true;
            table.SetPivot(new TableHandlers.ExcelTableHandler.Table.Pivot(6, 4));
            table.Draw(ref WORK_SHEET, ref data);

            table.Settings.BubbleColumnCaptions = true;
            WORK_SHEET.Cells[2, table.getRightBottomCorner().lcCol+3].Value = "Settings // UseSmartResetForFixedColsAndRows = [bool] and BubbleColumnCaptions = [bool]";
            WORK_SHEET.Cells[3, table.getRightBottomCorner().lcCol+3].Value = "SmartReset will cause the effect when only first drawn table freeze panes while single instance of ExcelTH is applied to draw";
            WORK_SHEET.Cells[4, table.getRightBottomCorner().lcCol + 3].Value = "BubbleColumnCaptions rules whether captions should be closer to the top or to the bottom of header's area";
            table.SetPivot(new TableHandlers.ExcelTableHandler.Table.Pivot(7, table.getRightBottomCorner().lcCol+3));
            table.Draw(ref WORK_SHEET, ref data);

            table.Settings.DrawHeaders = false;
            WORK_SHEET.Cells[table.getRightBottomCorner().lcRow + 3, table.pivot.lcCol].Value = "Settings // DrawHeaders = [bool]";
            WORK_SHEET.Cells[table.getRightBottomCorner().lcRow + 4, table.pivot.lcCol].Value = "DrawHeaders is set to false if you want to hide all of the headers and header groups";
            WORK_SHEET.Cells[table.getRightBottomCorner().lcRow + 5, table.pivot.lcCol].Value = "That's it. Enjoy!";
            table.SetPivot(new TableHandlers.ExcelTableHandler.Table.Pivot(table.getRightBottomCorner().lcRow+6, table.pivot.lcCol));
            table.Draw(ref WORK_SHEET, ref data);
        }



















        

        public static void GenerateDemoUsingGembox(object[][] data, int demoNum = 0, string fileName = "gb12345.xlsx")
        {
            string worksheetName = "Demo Gembox";

            Console.WriteLine("Building excel report using Gembox.");
            Console.WriteLine("Creating workbook.");
            SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
            ExcelFile wb = new ExcelFile();

            Console.WriteLine("Workbook created, creating worksheet.");
            wb.Worksheets.Add(worksheetName);
            for(int i=0; i<=wb.Worksheets.Count; i++) { wb.Worksheets.Remove(0); }
            wb.Worksheets.Add(worksheetName);
            ExcelWorksheet WORK_SHEET = wb.Worksheets[0];
            if (WORK_SHEET != null)
            {
                Console.WriteLine("Worksheet's ready. Building report.");
                //CreateSCTableWithGembox();
                if (demoNum == 0) CreateDemoHeadersTableWithGembox(ref WORK_SHEET);
                if (demoNum == 1)
                {
                    CreateSCTableWithGembox(ref WORK_SHEET);
                    int columnCount = WORK_SHEET.CalculateMaxUsedColumns();
                    for (int i = 0; i < columnCount; i++)
                    {
                        WORK_SHEET.Columns[i].AutoFit(1, WORK_SHEET.Rows[1], WORK_SHEET.Rows[WORK_SHEET.Rows.Count - 1]);
                    }
                }
                
            }

            Console.WriteLine("Report filled. Saving.");
            wb.Save(Get_Full_Output_Path(fileName), SaveOptions.XlsxDefault);

            Console.WriteLine("Job's done. Press any key.");
        }



        private static void CreateSCTableWithGembox(ref ExcelWorksheet WORK_SHEET)
        {
            var table = new TableHandlers.GemboxTableHandler.Table();

            //Set specific fill for headers [using .Formatter; priority 1 is higher than 0]
            var styleH = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.SCHeaderCell, 0);

            //Set specific fill for rows [using .Formatter; priority 1 is higher than 0]
            var styleC = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.StandartInlineCell, 0);
            var styleEmptyWeaponOutline = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.EmptyWeaponStyle, 1);

            //Some extra-feature that hopefully will be modified for good. It's usable for now, but lacks of modification-proof
            //Specific formats for the part of entire row -- goes AFTER all cells in specific row are formatted
            table.addRowFormat(styleEmptyWeaponOutline);

            //Link new columns to array of data by indexes [addColumn(TITLE, DATA_ARRAY_INDEX)], add styling
            table.addColumn("#num [SRC]", 10).addFormat(styleH)
                .addCellFormat(new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.StandartInlineCell, 0))
                    .addCellFormat(new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.StandartInlineRecordNumberCell, 1));
            //  VS
            table.addColumn("[NextEmpty]", 10).addFormat(styleH)
                .addCellFormat(new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.StandartInlineCell, 0))
                    .addCellFormat(new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.StandartInlineRecordNumberCell, 1))
                        .EnableAutoMergeNextEmptyValues();
            //Enable merging modes for column: 
            //    -treat next same values in data array column as one and merge them;
            //    -treat all next empty values that do follow after non-empty value in data array column as last non-empty value and merge them;
            //    -break ALL COLUMNS merging rules when value is changed [provides infoblocks of multirow data]
            table.addColumn("Info [SRC]", 14).addFormat(styleH).addCellFormat(styleC);
            table.addColumn("[MergeSame]", 14).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();

            table.addColumn("Unit [SRC]", 0).addFormat(styleH).addCellFormat(styleC);
            table.addColumn("[NE + BlockDelimeter]", 0).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeNextEmptyValues().UseRowDataAsInfoBlockDelimeter();

            table.addColumn("Health Points", 1).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Shield Points", 2).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Description", 3).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();
            table.addColumn("Damage", 4).addCellFormat(styleC).addFormat(styleH).EnableAutoMergeSameValues();
            table.addColumn("Range", 5).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();
            table.addColumn("Cooldown", 6).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues();
            table.addColumn("Minerals", 7).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Vespene", 8).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("Resources", 9).addFormat(styleH).addCellFormat(styleC).EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();

            //Set specific fill for different groups [using .Formatter; priority 1 is higher than 0]
            var styleSVA = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.SurvivabilityGroupCell, 1);
            var stylePwr = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.PowerGroupCell, 1);
            var styleCst = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.CostGroupCell, 1);
            var styleSgh = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.StrengthGroupCell, 1);
            var styleGeneralHeaderGroup = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo1.SCHeaderCell, 0);

            //Adding groups of headers
            table.addHeaderGroup("Strength", "Health Points", "Shield Points",
                        "Description", "Damage", "Range", "Cooldown").addFormat(styleSgh).addFormat(styleGeneralHeaderGroup);
            table.addHeaderGroup("Survivability", "Health Points", "Shield Points").addFormat(styleSVA).addFormat(styleGeneralHeaderGroup);
            table.addHeaderGroup("Powers", "Description", "Damage", "Range", "Cooldown").addFormat(stylePwr).addFormat(styleGeneralHeaderGroup);
            table.addHeaderGroup("Cost", "Minerals", "Vespene", "Resources").addFormat(styleCst).addFormat(styleGeneralHeaderGroup);



            WORK_SHEET.Cells[2, 2].Value = "Row cells have 3 settings available that can be accessed either by methods or by settings of a specific column.";
            WORK_SHEET.Cells.GetSubrangeAbsolute(2, 2, 2, 14).Merged = true;
            WORK_SHEET.Cells[3, 2].Value = "Those are: AutoMergeSameValues = [bool]; AutoMergeNextEmptyValues = [bool]; IsBlockDelimeter = [bool]";
            WORK_SHEET.Cells.GetSubrangeAbsolute(3, 2, 3, 14).Merged = true;
            WORK_SHEET.Cells[4, 2].Value = "Purpose: to control merge of column's cells content.";
            WORK_SHEET.Cells.GetSubrangeAbsolute(4, 2, 4, 14).Merged = true;
            WORK_SHEET.Cells[6, 2].Value = "While AutoMergeSameValues and AutoMergeNextEmptyValues are obvious, IsBlockDelimeter is not.";
            WORK_SHEET.Cells.GetSubrangeAbsolute(6, 2, 6, 14).Merged = true;
            WORK_SHEET.Cells[7, 2].Value = "IsBlockDelimeter is used to prevent merging any cells in current row with a previous row that is logically related to another information group (keyValue=>data / Infoblock structure).";
            WORK_SHEET.Cells.GetSubrangeAbsolute(7, 2, 7, 18).Merged = true;
            WORK_SHEET.Cells[8, 2].Value = "For better understanding, see table below. It has columns with source data marked with [SRC] key and ongoing column with some setting is enabled.";
            WORK_SHEET.Cells.GetSubrangeAbsolute(8, 2, 8, 14).Merged = true;
            //Specifying margins: coordinates of the left upper row [new Pivot(RowNumber, ColNumber)]
            table.SetPivot(new TableHandlers.GemboxTableHandler.Table.Pivot(9, 2));

            var data = ReportData.LoadDataDemo1();
            table.Draw(ref WORK_SHEET, ref data);

        }


        private static void CreateDemoHeadersTableWithGembox(ref ExcelWorksheet WORK_SHEET)
        {
            var table = new TableHandlers.GemboxTableHandler.Table();

            //Setting global header style for text
            var HGeneral = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo3.GeneralHeaderFormat, 0);

            //Setting specific fill for different groups and total [using .Formatter priority 1, higher than HGeneral 0]
            var HCell = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo3.HeaderCell, 1);
            var HGr1 = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo3.HeaderGroup1Cell, 1);
            var HGr2 = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo3.HeaderGroup2Cell, 1);
            var HGr3 = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo3.HeaderGroup3Cell, 1);

            var Cell = new TableHandlers.GemboxTableHandler.Table.Formatter(CellFormatsGembox.Demo3.StandartInlineCell, 0);

            table.addColumn("HD1", 0).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues().UseRowDataAsInfoBlockDelimeter();
            table.addColumn("HD2", 1).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("HD3", 2).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues().EnableAutoMergeNextEmptyValues();
            table.addColumn("HD4", 3).addFormat(HGeneral).addFormat(HCell).addCellFormat(Cell);//.EnableAutoMergeSameValues();

            table.addHeaderGroup("GR1", "HD1").addFormat(HGeneral).addFormat(HGr1);
            table.addHeaderGroup("GR2", "HD1", "HD3", "HD4").addFormat(HGeneral).addFormat(HGr2);
            table.addHeaderGroup("GR2.5", "HD3").addFormat(HGeneral).addFormat(HGr1);
            table.addHeaderGroup("GR3", "HD4").addFormat(HGeneral).addFormat(HGr3);

            var data = ReportData.LoadDataDemo3();
            WORK_SHEET.Cells[1, 1].Value = "Headers have 5 settings accessed via TABLE().Settings";
            WORK_SHEET.Cells[2, 1].Value = "Settings // FixFirstNColumns = [int amount] and FixHeaderRows = [bool]";
            WORK_SHEET.Cells[3, 1].Value = "Those are used to freeze some of vertical or horizontal areas of worksheet";
            table.Settings.FixFirstNColumns = 1;
            table.Settings.FixHeaderRows = true;
            table.Settings.UseSmartResetForFixedColsAndRows = true;
            table.SetPivot(new TableHandlers.GemboxTableHandler.Table.Pivot(5, 3));
            table.Draw(ref WORK_SHEET, ref data);

            table.Settings.BubbleColumnCaptions = true;
            WORK_SHEET.Cells[1, table.getRightBottomCorner().lcCol + 3].Value = "Settings // UseSmartResetForFixedColsAndRows = [bool] and BubbleColumnCaptions = [bool]";
            WORK_SHEET.Cells[2, table.getRightBottomCorner().lcCol + 3].Value = "SmartReset will cause the effect when only first drawn table freeze panes while single instance of ExcelTH is applied to draw";
            WORK_SHEET.Cells[3, table.getRightBottomCorner().lcCol + 3].Value = "BubbleColumnCaptions rules whether captions should be closer to the top or to the bottom of header's area";
            table.SetPivot(new TableHandlers.GemboxTableHandler.Table.Pivot(6, table.getRightBottomCorner().lcCol + 3));
            table.Draw(ref WORK_SHEET, ref data);

            table.Settings.DrawHeaders = false;
            WORK_SHEET.Cells[table.getRightBottomCorner().lcRow + 3, table.pivot.lcCol].Value = "Settings // DrawHeaders = [bool]";
            WORK_SHEET.Cells[table.getRightBottomCorner().lcRow + 4, table.pivot.lcCol].Value = "DrawHeaders is set to false if you want to hide all of the headers and header groups";
            WORK_SHEET.Cells[table.getRightBottomCorner().lcRow + 5, table.pivot.lcCol].Value = "That's it. Enjoy!";
            table.SetPivot(new TableHandlers.GemboxTableHandler.Table.Pivot(table.getRightBottomCorner().lcRow + 6, table.pivot.lcCol));
            table.Draw(ref WORK_SHEET, ref data);
        }

    }
}
