using NetOffice.ExcelApi;
using System;
using System.Linq;

namespace TableHandlers
{
    /*
      Made by Alex K.
         */
    public class ExcelTableHandler
    {
        public class Table
        {
            public Table()
            {
                this.Settings = new _Settings();
                this.RowFormatRules = new Row();
            }

            public abstract class Formattable
            {
                private Table.Formatter[] formats;
                public Formattable addFormat(Formatter f){
                    if (formats != null) { 
                        var current_format_with_same_priority = formats.Where(x => x!=null && f != null && x.priority == f.priority).Select(x=>x).FirstOrDefault();
                        if (current_format_with_same_priority == null) { 
                            Utilities.add<Formatter>(ref formats, f);
                        } else
                        {
                            current_format_with_same_priority = f;
                        }
                    } else Utilities.add<Formatter>(ref formats, f);
                    return this;
                }

                public void Format(Range x, object[] dataRow = null, object GlobalConditionsObject = null)
                {
                    if (this.formats != null)
                    {
                        var orderedFormatList = this.formats.OrderBy(f => f.priority).ToArray();
                        for (int i = 0; i < orderedFormatList.Length; i++)
                        {
                            orderedFormatList[i].ContiniousFormat(x, dataRow, GlobalConditionsObject);
                        }
                    }
                }
            }

            public class HeaderGroup : Table.Formattable
            {
                public int level = 0; //CLOSEST TO ACTUAL HEADERS
                public string Caption;
                public int[] colNumbers;
                public int[] groupNumbers;

                public HeaderGroup(string name=""){
                    this.Caption = name;
                }
            }

            protected HeaderGroup[] headerGroups;
            protected Column[] headers;
            public Row RowFormatRules;
            public Pivot pivot;
            private Worksheet xs;
            private string IsMyBirthDay = "[09011990]";
            private int lastDrawnNumber = 0;
            protected bool? isFirstInstanceDrawn = null;
            public _Settings Settings;
            //private Formatter BasicHeaderFormatter;

            private void Reset() { this.lastDrawnNumber = 0; if (this.Settings.UseSmartResetForFixedColsAndRows) { if (this.isFirstInstanceDrawn == true) { this.Settings.FixFirstNColumns = 0; this.Settings.FixHeaderRows = false; this.isFirstInstanceDrawn = false; } } }
            public Pivot getRightBottomCorner()
            {
                return new Pivot(this.pivot.lcRow+lastDrawnNumber, this.pivot.lcCol + this.headers.Length - 1);
            }

            public class _Settings
            {
                public bool DrawHeaders = true;
                public bool BubbleColumnCaptions = false;
                public bool FixHeaderRows = false;
                public int FixFirstNColumns = 0;
                public bool UseSmartResetForFixedColsAndRows = true;
            }
            

            public class Pivot
            {
                public int lcRow;
                public int lcCol;

                public Pivot(int leftCornerRow, int LeftCornerColumn)
                {
                    this.lcRow = leftCornerRow;
                    this.lcCol = LeftCornerColumn;
                }
            }

            public Table addRowFormat(Formatter f)
            {
                RowFormatRules.addFormat(f);
                return this;
            }

            public class Column : Table.Formattable
            {
                public _Settings Settings;
                public class _Settings
                {
                    public bool AutoMergeSameValues = false;
                    public bool AutoMergeNextEmptyValues = false;
                    //принудительно отделяет ряды от последующих
                    public bool IsBlockDelimeter = false;
                }

                public string Caption;
                public int rowDataIndex;
                public Row ColumnRowFormats;

                
                public Column EnableAutoMergeSameValues() {this.Settings.AutoMergeSameValues = true; return this;}
                public Column EnableAutoMergeNextEmptyValues() { this.Settings.AutoMergeNextEmptyValues = true; return this; }
                public Column UseRowDataAsInfoBlockDelimeter() { this.Settings.IsBlockDelimeter = true; return this; }

                public Column(int rowDataIndex, string caption = ""){
                    this.Caption = caption;
                    this.rowDataIndex = rowDataIndex;
                    this.ColumnRowFormats = new Row();
                    this.Settings = new _Settings();
                }

                public Column addCellFormat(Formatter f)
                {
                    ColumnRowFormats.addFormat(f);
                    return this;
                }

                public new Column addFormat(Formatter f)
                {
                    base.addFormat(f);
                    return this;
                }
            }

            public class Row : Table.Formattable
            {
                
                //public object[] data;
                //public object[] previousRowData;
            }

            public class Formatter
            {
                public int priority;
                public Action<Range, Range> OpenFormat;
                public Action<Range, object[], object> ContiniousFormat;
                public Action<Range, Range> CloseFormat;

                public Formatter(Action<Range, object[], object> ContiniousFormat, int priority = 0)
                {
                    this.priority = priority;
                    this.ContiniousFormat = ContiniousFormat;
                }

                public Formatter SetContiniousFormat(Action<Range, object[], object> ContiniousFormat, int priority = 0)
                {
                    this.ContiniousFormat = ContiniousFormat;
                    this.priority = priority;
                    return this;
                }
            }

            public Column addColumn(string caption, int rowDataIndex)
            {
                Utilities.add<Column>(ref headers, new Column(rowDataIndex, caption));
                return headers[headers.Length - 1];
            }


            public HeaderGroup addHeaderGroup(string groupName, params string[] colNames)
            {
                Utilities.add<HeaderGroup>(ref headerGroups, new HeaderGroup(groupName));
                var hg = headerGroups[headerGroups.Length - 1];
                hg.level = getHighestHeaderLevelInGroups(colNames)+1;
                
                foreach (var col in colNames)
                {
                    Utilities.add<int>(ref hg.colNumbers, this.getColumnIndexByName(col));
                }

                return hg;
            }

            private int getHighestHeaderLevelInGroups(params string[] colNames)
            {
                var maxIndx = 0;
                foreach (var col in colNames)
                {
                    var colIndx = getColumnIndexByName(col);
                    
                    foreach (var hg in headerGroups)
                    {
                        if ((hg.colNumbers!=null)&&((Array.IndexOf(hg.colNumbers, colIndx))!=-1)&&(maxIndx<hg.level)) maxIndx = hg.level;
                    }
                }
                
                return maxIndx;
            }


            private int getColumnIndexByName(string colName){
                return Array.IndexOf(this.headers, this.headers.Where(x=>(x!=null)&&(x.Caption==colName)).FirstOrDefault());   
            }


            public Table SetPivot(Pivot pvt)
            {
                this.pivot = pvt;
                return this;
            }

            public Table SetPivot(int leftCornerRow, int LeftCornerColumn)
            {
                this.pivot = new Pivot(leftCornerRow, LeftCornerColumn);
                return this;
            }

            /*
            public void addBasicHeaderStyle(Formatter f)
            {
                f.priority = -1;
                this.BasicHeaderFormatter = f;
                for (int i = 0; i < this.headers.Length; i--)
                {
                    if(this.headers[i].ColumnRowFormats)
                }
            }
            */
            public void Draw(ref Worksheet xs, ref object[][] data)
            {
                if (this.pivot != null)
                {
                    this.xs = xs;

                    this.Reset();

                    if (this.Settings.FixHeaderRows || this.Settings.FixFirstNColumns > 0)
                    {
                        this.xs.Application.ActiveWindow.FreezePanes = false;
                        this.xs.Application.ActiveWindow.Split = false;
                    }
                    if (this.Settings.DrawHeaders) this.DrawHeaders();

                    if (this.Settings.FixHeaderRows || this.Settings.FixFirstNColumns > 0)
                    {
                        this.xs.Application.ActiveWindow.FreezePanes = true;
                        this.xs.Application.ActiveWindow.Split = true;
                    }

                    if (data!=null) this.DrawData(ref data);

                    if(this.Settings.UseSmartResetForFixedColsAndRows) if(this.isFirstInstanceDrawn == null) this.isFirstInstanceDrawn = true;
                }
            }
            /**
             * 
             *                    if (this.headers[j].Settings.IsFixed)
                    {
                        rng(lastDrawnNumber, j, lastDrawnNumber + DataArr.Length - 1, j).Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbBlack;
                        rng(lastDrawnNumber, j, lastDrawnNumber + DataArr.Length - 1, j).Activate();
                        //this.xs.Application.ActiveWindow.Sp = true;

                    } 
             */
            private void DrawHeaders(){
                if (this.headerGroups!=null)
                {
                    var max_lvl = this.headerGroups.Max(x => x.level);
                    //var SplitMarker = new CrawlingObject();
                    for (int i = max_lvl; i > 0; i--)
                    {
                        var groups = this.headerGroups.Where(x => x.level == i).ToArray();
                        if (groups != null)
                        {
                            foreach (var gr in groups)
                            {
                                drawHeaderGroup(gr, i-1);
                            }
                            lastDrawnNumber++;
                        }
                    }
                }

                bool atLeastOneColumnThatUsesLastRow = false;
                if (this.Settings.BubbleColumnCaptions) {
                    atLeastOneColumnThatUsesLastRow = (this.headerGroups.SelectMany(x => x.colNumbers, (x, y) => new { col = y, level = x.level })
                        .GroupBy(x => x.col, (x, y) => y.Count()).Where(x => x >= this.headerGroups.Select(z => z.level).Max()).Count()>0)?true:false;//.FirstOrDefault();
                }

                for (int i = 0; i < this.headers.Length; i++)
                {
                    drawHeader(this.headers[i], (atLeastOneColumnThatUsesLastRow?lastDrawnNumber:lastDrawnNumber-1), i);
                }

                if (this.Settings.FixHeaderRows)
                    //rng(lastDrawnNumber, j, lastDrawnNumber + DataArr.Length - 1, j).Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbBlack;
                    this.xs.Application.ActiveWindow.SplitRow = this.pivot.lcRow+lastDrawnNumber;
                if (this.Settings.FixFirstNColumns>0)
                    this.xs.Application.ActiveWindow.SplitColumn = this.pivot.lcCol + this.Settings.FixFirstNColumns-1;

                var By_the_way = this.IsMyBirthDay;
                if (atLeastOneColumnThatUsesLastRow||!this.Settings.BubbleColumnCaptions) lastDrawnNumber++;
            }


            //переписать с кроулингом, рассмотреть вопрос: ***xx----
            //                                             xxxxx--** 
            private void drawHeaderGroup(HeaderGroup group, int lineNum)
            {
                var colIndexes = group.colNumbers.OrderBy(x => x).ToArray();
                int left = colIndexes[0];
                bool is_margin_left = true;
                for (int i = 0; i < colIndexes.Length; i++)
                {
                    if (i>0 && colIndexes[i] - colIndexes[i-1] != 1)
                    {
                        rng(lineNum, left, lineNum, colIndexes[i - 1]).Merge();
                            
                        group.Format(rng(lineNum, left, lineNum, colIndexes[i - 1]));
                        if (is_margin_left)
                        {
                            rng(lineNum, left).Value = group.Caption;
                            is_margin_left = false;
                        }
                        left = colIndexes[i];
                    }
                  //  }
                }

                rng(lineNum, left, lineNum, colIndexes[colIndexes.Length-1]).Merge();
                group.Format(rng(lineNum, left, lineNum, colIndexes[colIndexes.Length - 1]));
                if (is_margin_left) rng(lineNum, left).Value = group.Caption;
            }


            private void drawHeader(Column column, int lastLineNum, int colNum)
            {
                //if (colNum != 1) return;
                var Filler = new CrawlingObject();
                for (int i = 0; i <= lastDrawnNumber; i++)
                {
                    var lvl = this.headerGroups.Where(x => (x.level == i+1) && (x.colNumbers.Contains(colNum))).FirstOrDefault();

                    if (lvl == null){
                        if (Filler.lastValue != null)
                        {
                            Filler.Break(i, null);
                        }
                        else//if (Filler.lastValue == null)
                        {
                            if (Filler.startPoint == -1) Filler.Break(i, null);
                            else
                            {
                                Filler.endPoint++;
                            }
                        }
                    }
                    if (lvl != null)
                    {
                        if (Filler.lastValue == null && Filler.startPoint!=-1)
                        {
                            rng(Filler.startPoint, colNum, Filler.endPoint, colNum).Merge();
                            column.Format(rng(Filler.startPoint, colNum, Filler.endPoint, colNum));
                            Filler.Break(i, lvl);
                        } else
                        {
                            Filler.Break(i, lvl);
                        }
                    }
                }


                if (this.Settings.BubbleColumnCaptions)
                {
                    if (Filler.startPoint <= lastLineNum)
                    {
                        rng(Filler.startPoint, colNum, lastLineNum, colNum).Merge();
                        column.Format(rng(Filler.startPoint, colNum, lastLineNum, colNum));
                    }

                    int top_cell = lastDrawnNumber;

                    var tmp2 = this.headerGroups.Select(x => new { lvl = x.level, cols = x.colNumbers })
                        .GroupBy(x => x.lvl, (x, y) => new { lvl = x, contains = y.SelectMany(z => z.cols) })
                        .ToDictionary(x=>x.lvl, (y)=>y.contains.ToArray()).OrderBy(x=>x.Key);//Where(x => !(x.colNumbers.Contains(colNum))).Select(x => x).ToArray();
                    
                    foreach(var tmp3 in tmp2)
                    {
                        if (!tmp3.Value.Contains(colNum)) { top_cell = tmp3.Key-1; break; }
                    }

                    rng(top_cell, colNum).Value = column.Caption;
                }
                else
                {
                    rng(Filler.startPoint, colNum, lastDrawnNumber, colNum).Merge();
                    column.Format(rng(Filler.startPoint, colNum, lastDrawnNumber, colNum));

                    var tmp2 = this.headerGroups.Select(x => new { lvl = x.level, cols = x.colNumbers })
                        .GroupBy(x => x.lvl, (x, y) => new { lvl = x, contains = y.SelectMany(z => z.cols) })
                        .ToDictionary(x => x.lvl, (y) => y.contains.ToArray()).OrderByDescending(x => x.Key);

                    var top_cell = 0;

                    foreach (var tmp3 in tmp2)
                    {
                        if (tmp3.Value.Contains(colNum)) {top_cell = tmp3.Key; break; } //else ;
                    }

                    rng(top_cell, colNum).Value = column.Caption;
                }
            }


            private void DrawData(ref object[][] DataArr)
            {
                //подгатавливаем общий формат
                for (int j = 0; j < this.headers.Length; j++)
                {
                    this.headers[j].ColumnRowFormats.Format(rng(lastDrawnNumber, j, lastDrawnNumber+DataArr.Length-1, j));
                }


                CrawlingObject[] delimetersTable = new CrawlingObject[this.headers.Length];
                var isDelimeterAmongUs = false;

                for (int i = 0; i < DataArr.Length; i++)
                {
                    var tmp = new object[this.headers.Length];
                    for (int j = 0; j < tmp.Length; j++)
                    {
                        tmp[j] = DataArr[i][this.headers[j].rowDataIndex];
                        //if (/*j != 2 && j != 3 &&*/ j != 4) continue;

                        var tmp_2 = tmp[j] as string;
                        if (tmp_2 !=null && tmp_2 == "") tmp[j] = null;

                        if (i == 0)
                        {/*
                            if (!this.headers[j].Settings.AutoMergeNextEmptyValues)
                                if (this.headers[j].Settings.AutoMergeSameValues) delimetersTable[j] = new CrawlingObject(lastDrawnNumber, tmp[j]);
                            else
                                if (tmp[j] != null) delimetersTable[j] = new CrawlingObject(lastDrawnNumber, tmp[j]);*/
                            if ((this.headers[j].Settings.AutoMergeNextEmptyValues)
                                && (tmp[j] != null)) delimetersTable[j] = new CrawlingObject(lastDrawnNumber, tmp[j]);
                            else if (this.headers[j].Settings.AutoMergeSameValues) delimetersTable[j] = new CrawlingObject(lastDrawnNumber, tmp[j]);
                                    
                        }
                        else
                        {
                            if (delimetersTable[j] != null)
                            {
                                if (!this.headers[j].Settings.AutoMergeNextEmptyValues)
                                {
                                    if ((delimetersTable[j].lastValue != null && !delimetersTable[j].lastValue.Equals(tmp[j])) || (delimetersTable[j].lastValue == null && tmp[j]!=null))
                                    {
                                        if (this.headers[j].Settings.AutoMergeSameValues) rng(delimetersTable[j].startPoint, j, delimetersTable[j].endPoint, j).Merge();
                                        //rng(delimetersTable[j].startPoint, j, delimetersTable[j].endPoint, j).Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbYellow;
                                        delimetersTable[j].Break(lastDrawnNumber, tmp[j]);
                                        if (this.headers[j].Settings.IsBlockDelimeter) isDelimeterAmongUs = true;
                                    }
                                    else delimetersTable[j].endPoint++;
                                }
                                else
                                {
                                    if ((delimetersTable[j].lastValue != null && !delimetersTable[j].lastValue.Equals(tmp[j]) && tmp[j] != null && !this.headers[j].Settings.AutoMergeSameValues)
                                            || ((delimetersTable[j].lastValue != null && !delimetersTable[j].lastValue.Equals(tmp[j]) && tmp[j] != null || (delimetersTable[j].lastValue == null && tmp[j] != null)) && this.headers[j].Settings.AutoMergeSameValues))
                                    {
                                        rng(delimetersTable[j].startPoint, j, delimetersTable[j].endPoint, j).Merge();
                                        //rng(delimetersTable[j].startPoint, j, delimetersTable[j].endPoint, j).Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbYellow;
                                        delimetersTable[j].Break(lastDrawnNumber, tmp[j]);
                                        if (this.headers[j].Settings.IsBlockDelimeter) isDelimeterAmongUs = true;
                                    }
                                    else delimetersTable[j].endPoint++;
                                }
                            }
                            else
                            {
                                //mb should move this -------\/ to if (i == 0), #373
                                if (tmp[j] != null/* || this.headers[j].Settings.AutoMergeSameValues*/) delimetersTable[j] = new CrawlingObject(lastDrawnNumber, tmp[j]);
                            }
                        }
                    }

                    rng(lastDrawnNumber, 0, lastDrawnNumber, tmp.Length - 1).Value = tmp;
                    this.RowFormatRules.Format(rng(lastDrawnNumber, 0, lastDrawnNumber, tmp.Length - 1), tmp);

                    if (isDelimeterAmongUs)
                    {
                        //rng(lastDrawnNumber, 8, lastDrawnNumber, 8).Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbFireBrick;
                        
                        for (int z = 0; z < this.headers.Length; z++)
                        {
                            //if (z != 2 && z != 3 && z != 4) continue;
                            if (this.headers[z].Settings.AutoMergeNextEmptyValues || this.headers[z].Settings.AutoMergeSameValues)
                                if ((delimetersTable[z] != null) && (delimetersTable[z].startPoint != -1) && (delimetersTable[z].startPoint != delimetersTable[z].endPoint))
                                {
                                    rng(delimetersTable[z].startPoint, z, (delimetersTable[z].endPoint < lastDrawnNumber) ? delimetersTable[z].endPoint : (lastDrawnNumber - 1), z).Merge();//.Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbYellowGreen;
                                    delimetersTable[z].Break(lastDrawnNumber, tmp[z]);
                                }
                        }
                    }

                    isDelimeterAmongUs = false;
                    lastDrawnNumber++;
                }

                //закрываем сет
                for (int z = 0; z < this.headers.Length; z++)
                {
                    if (((this.headers[z].Settings.AutoMergeNextEmptyValues || this.headers[z].Settings.AutoMergeSameValues))&&(delimetersTable[z] != null) && (delimetersTable[z].startPoint != -1) && (delimetersTable[z].startPoint != delimetersTable[z].endPoint))
                    {
                        rng(delimetersTable[z].startPoint, z, delimetersTable[z].endPoint, z).Merge();//Interior.Color = NetOffice.ExcelApi.Enums.XlRgbColor.rgbYellowGreen;
                    }
                }
            }
         
            private Range rng(int rx, int cy)
            {
                return this.xs.Cells[this.pivot.lcRow + rx, this.pivot.lcCol + cy];
            }

            private Range rng(int rx1, int cy1, int rx2, int cy2)
            {
                return this.xs.Range(rng(rx1, cy1), rng(rx2, cy2));
            }

        }

        public class CrawlingObject
        {
            public int startPoint;
            public int endPoint;
            public object lastValue;
            //public object endValue;

            //public abstract bool Proceed();

            public CrawlingObject(int startPoint = -1, object startValue = null)
            {
                Break(startPoint, startValue);
            }

            public void Break(int i, object newValue)
            {
                this.endPoint = i;
                this.startPoint = i;
                this.lastValue = newValue;
            }
        }

    }
}
