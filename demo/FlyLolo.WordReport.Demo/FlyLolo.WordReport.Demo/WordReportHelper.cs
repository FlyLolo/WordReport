using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace FlyLolo.WordReport.Demo
{
    public class WordReportHelper
    {
        private Word.Application wordApp = null;
        private Word.Document wordDoc = null;
        private DataSet dataSource = null;
        private object line = Word.WdUnits.wdLine;
        private string errorMsg = "";

        /// <summary>
        /// 根据模板文件,创建数据报告
        /// </summary>
        /// <param name="templateFile">模板文件名(含路径)</param>
        /// <param name="newFilePath">新文件路径)</param>
        /// <param name="dataSource">数据源,包含多个datatable</param>
        /// <param name="saveFormat">新文件格式:</param>
        public bool CreateReport(string templateFile, DataSet dataSource, out string errorMsg, string newFilePath, ref string newFileName, int saveFormat = 16)
        {
            this.dataSource = dataSource;
            errorMsg = this.errorMsg;
            bool rtn = OpenTemplate(templateFile)
                && SetContent(new WordElement(wordDoc.Range(), dataRow: dataSource.Tables[dataSource.Tables.Count - 1].Rows[0]))
                && UpdateTablesOfContents()
                && SaveFile(newFilePath, ref newFileName, saveFormat);

            CloseAndClear();
            return rtn;
        }

        /// <summary>
        /// 打开模板文件
        /// </summary>
        /// <param name="templateFile"></param>
        /// <returns></returns>
        private bool OpenTemplate(string templateFile)
        {
            if (!File.Exists(templateFile))
            {
                return false;
            }

            wordApp = new Word.Application();
            wordApp.Visible = true;//使文档可见,调试用
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
            object file = templateFile;
            wordDoc = wordApp.Documents.Open(ref file, ReadOnly: false);
            return true;
        }

        /// <summary>
        /// 为指定区域写入数据
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private bool SetContent(WordElement element)
        {
            string currBookMarkName = string.Empty;
            string startWith = "loop_" + (element.Level + 1).ToString() + "_";
            foreach (Word.Bookmark item in element.Range.Bookmarks)
            {
                currBookMarkName = item.Name;

                if (currBookMarkName.StartsWith(startWith) && (!currBookMarkName.Equals(element.ElementName)))
                {
                    SetLoop(new WordElement(item.Range, currBookMarkName, element.DataRow, element.GroupBy));
                }

            }

            SetLabel(element);

            SetTable(element);

            SetChart(element);

            return true;
        }

        /// <summary>
        /// 处理循环
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private bool SetLoop(WordElement element)
        {
            DataRow[] dataRows = dataSource.Tables[element.TableIndex].Select(element.GroupByString);
            int count = dataRows.Count();
            element.Range.Select();

            //第0行作为模板  先从1开始  循环后处理0行;
            for (int i = 0; i < count; i++)
            {

                element.Range.Copy();  //模板loop复制
                wordApp.Selection.InsertParagraphAfter();//换行 不会清除选中的内容,TypeParagraph 等同于回车,若当前有选中内容会被清除. TypeParagraph 会跳到下一行,InsertParagraphAfter不会, 所以movedown一下.
                wordApp.Selection.MoveDown(ref line, Missing.Value, Missing.Value);
                wordApp.Selection.Paste(); //换行后粘贴复制内容
                int offset = wordApp.Selection.Range.End - element.Range.End; //计算偏移量

                //复制书签,书签名 = 模板书签名 + 复制次数
                foreach (Word.Bookmark subBook in element.Range.Bookmarks)
                {
                    if (subBook.Name.Equals(element.ElementName))
                    {
                        continue;
                    }

                    wordApp.Selection.Bookmarks.Add(subBook.Name + "_" + i.ToString(), wordDoc.Range(subBook.Start + offset, subBook.End + offset));
                }

                SetContent(new WordElement(wordDoc.Range(wordApp.Selection.Range.End - (element.Range.End - element.Range.Start), wordApp.Selection.Range.End), element.ElementName + "_" + i.ToString(), dataRows[i], element.GroupBy));
            }

            element.Range.Delete();

            return true;
        }

        /// <summary>
        /// 处理简单Label
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private bool SetLabel(WordElement element)
        {
            if (element.Range.Bookmarks != null && element.Range.Bookmarks.Count > 0)
            {
                string startWith = "label_" + element.Level.ToString() + "_";
                string bookMarkName = string.Empty;
                foreach (Word.Bookmark item in element.Range.Bookmarks)
                {
                    bookMarkName = item.Name;

                    if (bookMarkName.StartsWith(startWith))
                    {
                        bookMarkName = WordElement.GetName(bookMarkName);

                        item.Range.Text = element.DataRow[bookMarkName].ToString();
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// 填充Table
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private bool SetTable(WordElement element)
        {
            if (element.Range.Tables != null && element.Range.Tables.Count > 0)
            {
                string startWith = "table_" + element.Level.ToString() + "_";
                foreach (Word.Table table in element.Range.Tables)
                {
                    if (!string.IsNullOrEmpty(table.Title) && table.Title.StartsWith(startWith))
                    {
                        WordElement tableElement = new WordElement(null, table.Title, element.DataRow);

                        TableConfig config = new TableConfig(table.Descr);

                        object dataRowTemplate = table.Rows[config.DataRow];
                        Word.Row SummaryRow = null;
                        DataRow SummaryDataRow = null;
                        DataTable dt = dataSource.Tables[tableElement.TableIndex];
                        DataRow[] dataRows = dataSource.Tables[tableElement.TableIndex].Select(tableElement.GroupByString); ;

                        if (config.SummaryRow > 0)
                        {
                            SummaryRow = table.Rows[config.SummaryRow];
                            SummaryDataRow = dt.Select(string.IsNullOrEmpty(tableElement.GroupByString) ? config.SummaryFilter : tableElement.GroupByString + " and  " + config.SummaryFilter).FirstOrDefault();
                        }

                        foreach (DataRow row in dataRows)
                        {
                            if (row == SummaryDataRow)
                            {
                                continue;
                            }

                            Word.Row newRow = table.Rows.Add(ref dataRowTemplate);
                            for (int j = 0; j < table.Columns.Count; j++)
                            {
                                newRow.Cells[j + 1].Range.Text = row[j].ToString(); ;
                            }

                        }

                        ((Word.Row)dataRowTemplate).Delete();

                        if (config.SummaryRow > 0 && SummaryDataRow != null)
                        {
                            for (int j = 0; j < SummaryRow.Cells.Count; j++)
                            {
                                string temp = SummaryRow.Cells[j + 1].Range.Text.Trim().Replace("\r\a", "");

                                if (!string.IsNullOrEmpty(temp) && temp.Length > 2 && dt.Columns.Contains(temp.Substring(1, temp.Length - 2)))
                                {
                                    SummaryRow.Cells[j + 1].Range.Text = SummaryDataRow[temp.Substring(1, temp.Length - 2)].ToString();
                                }
                            }
                        }

                        table.Title = tableElement.Name;
                    }


                }
            }

            return true;
        }

        /// <summary>
        /// 处理图表
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        private bool SetChart(WordElement element)
        {
            if (element.Range.InlineShapes != null && element.Range.InlineShapes.Count > 0)
            {
                List<Word.InlineShape> chartList = element.Range.InlineShapes.Cast<Word.InlineShape>().Where(m => m.Type == Word.WdInlineShapeType.wdInlineShapeChart).ToList();
                string startWith = "chart_" + element.Level.ToString() + "_";
                foreach (Word.InlineShape item in chartList)
                {
                    Word.Chart chart = item.Chart;
                    if (!string.IsNullOrEmpty(chart.ChartTitle.Text) && chart.ChartTitle.Text.StartsWith(startWith))
                    {
                        WordElement chartElement = new WordElement(null, chart.ChartTitle.Text, element.DataRow);

                        DataTable dataTable = dataSource.Tables[chartElement.TableIndex];
                        DataRow[] dataRows = dataTable.Select(chartElement.GroupByString);

                        int columnCount = dataTable.Columns.Count;
                        List<int> columns = new List<int>();

                        foreach (var dr in dataRows)
                        {
                            for (int i = chartElement.ColumnStart == -1 ? 0 : chartElement.ColumnStart - 1; i < (chartElement.ColumnEnd == -1 ? columnCount : chartElement.ColumnEnd); i++)
                            {
                                if (columns.Contains(i) || dr[i] == null || string.IsNullOrEmpty(dr[i].ToString()))
                                {

                                }
                                else
                                {
                                    columns.Add(i);
                                }
                            }
                        }
                        columns.Sort();
                        columnCount = columns.Count;
                        int rowsCount = dataRows.Length;

                        Word.ChartData chartData = chart.ChartData;

                        //chartData.Activate();
                        //此处有个比较疑惑的问题, 不执行此条,生成的报告中的图表无法再次右键编辑数据. 执行后可以, 但有两个问题就是第一会弹出Excel框, 处理完后会自动关闭. 第二部分chart的数据range设置总不对
                        //不知道是不是版本的问题, 谁解决了分享一下,谢谢

                        Excel.Workbook dataWorkbook = (Excel.Workbook)chartData.Workbook;
                        dataWorkbook.Application.Visible = false;

                        Excel.Worksheet dataSheet = (Excel.Worksheet)dataWorkbook.Worksheets[1];
                        //设定范围  
                        string a = (chartElement.ColumnNameForHead ? rowsCount + 1 : rowsCount) + "|" + columnCount;
                        Console.WriteLine(a);

                        Excel.Range tRange = dataSheet.Range["A1", dataSheet.Cells[(chartElement.ColumnNameForHead ? rowsCount + 1 : rowsCount), columnCount]];
                        Excel.ListObject tbl1 = dataSheet.ListObjects[1];
                        //dataSheet.ListObjects[1].Delete(); //想过重新删除再添加  这样 原有数据清掉了, 但觉得性能应该会有所下降
                        //Excel.ListObject tbl1 = dataSheet.ListObjects.AddEx();
                        tbl1.Resize(tRange);
                        for (int j = 0; j < rowsCount; j++)
                        {
                            DataRow row = dataRows[j];
                            for (int k = 0; k < columnCount; k++)
                            {
                                dataSheet.Cells[j + 2, k + 1].FormulaR1C1 = row[columns[k]];
                            }
                        }

                        if (chartElement.ColumnNameForHead)
                        {
                            for (int k = 0; k < columns.Count; k++)
                            {
                                dataSheet.Cells[1, k + 1].FormulaR1C1 = dataTable.Columns[columns[k]].ColumnName;
                            }
                        }
                        chart.ChartTitle.Text = chartElement.Name;
                        //dataSheet.Application.Quit();
                    }
                }
            }

            return true;
        }

        /// <summary>
        /// 更新目录
        /// </summary>
        /// <returns></returns>
        private bool UpdateTablesOfContents()
        {
            foreach (Word.TableOfContents item in wordDoc.TablesOfContents)
            {
                item.Update();
            }

            return true;
        }

        /// <summary>
        /// 保存文件
        /// </summary>
        /// <param name="newFilePath"></param>
        /// <param name="newFileName"></param>
        /// <param name="saveFormat"></param>
        /// <returns></returns>
        private bool SaveFile(string newFilePath, ref string newFileName, int saveFormat = 16)
        {
            if (string.IsNullOrEmpty(newFileName))
            {
                newFileName = DateTime.Now.ToString("yyyyMMddHHmmss");

                switch (saveFormat)
                {
                    case 0:// Word.WdSaveFormat.wdFormatDocument
                        newFileName += ".doc";
                        break;
                    case 16:// Word.WdSaveFormat.wdFormatDocumentDefault
                        newFileName += ".docx";
                        break;
                    case 17:// Word.WdSaveFormat.wdFormatPDF
                        newFileName += ".pdf";
                        break;
                    default:
                        break;
                }
            }

            object newfile = Path.Combine(newFilePath, newFileName);
            object wdSaveFormat = saveFormat;
            wordDoc.SaveAs(ref newfile, ref wdSaveFormat);
            return true;
        }

        /// <summary>
        /// 清理
        /// </summary>
        private void CloseAndClear()
        {
            if (wordApp == null)
            {
                return;
            }
            wordDoc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
            wordApp.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDoc);
            System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            wordDoc = null;
            wordApp = null;
            GC.Collect();
            KillProcess("Excel", "WINWORD");
        }

        /// <summary>
        /// 杀进程..
        /// </summary>
        /// <param name="processNames"></param>
        private void KillProcess(params string[] processNames)
        {
            //Process myproc = new Process();
            //得到所有打开的进程  
            try
            {
                foreach (string name in processNames)
                {
                    foreach (Process thisproc in Process.GetProcessesByName(name))
                    {
                        if (!thisproc.CloseMainWindow())
                        {
                            if (thisproc != null)
                                thisproc.Kill();
                        }
                    }
                }
            }
            catch (Exception)
            {
                //throw Exc;
                // msg.Text+=  "杀死"  +  processName  +  "失败！";  
            }
        }
    }

    /// <summary>
    /// 封装的Word元素
    /// </summary>
    public class WordElement
    {
        public WordElement(Word.Range range, string elementName = "", DataRow dataRow = null, Dictionary<string, string> groupBy = null, int tableIndex = 0)
        {
            this.Range = range;
            this.ElementName = elementName;
            this.GroupBy = groupBy;
            this.DataRow = dataRow;
            if (string.IsNullOrEmpty(elementName))
            {
                this.Level = 0;
                this.TableIndex = tableIndex;
                this.Name = string.Empty;
                this.ColumnNameForHead = false;
            }
            else
            {
                string[] element = elementName.Split('_');
                this.Level = int.Parse(element[1]);
                this.ColumnNameForHead = false;
                this.ColumnStart = -1;
                this.ColumnEnd = -1;

                if (element[0].Equals("label"))
                {
                    this.Name = element[2];
                    this.TableIndex = 0;
                }
                else
                {
                    this.Name = element[4];
                    this.TableIndex = int.Parse(element[2]) - 1;

                    if (!string.IsNullOrEmpty(element[3]))
                    {
                        string[] filters = element[3].Split(new string[] { "XX" }, StringSplitOptions.RemoveEmptyEntries);
                        if (this.GroupBy == null)
                        {
                            this.GroupBy = new Dictionary<string, string>();
                        }
                        foreach (string item in filters)
                        {
                            if (!this.GroupBy.Keys.Contains(item))
                            {
                                this.GroupBy.Add(item, dataRow[item].ToString());
                            }

                        }
                    }

                    if (element[0].Equals("chart") && element.Count() > 5)
                    {
                        this.ColumnNameForHead = element[5].Equals("1");
                        this.ColumnStart = string.IsNullOrEmpty(element[6]) ? -1 : int.Parse(element[6]);
                        this.ColumnEnd = string.IsNullOrEmpty(element[7]) ? -1 : int.Parse(element[7]);
                    }
                }
            }
        }

        public Word.Range Range { get; set; }
        public int Level { get; set; }
        public int TableIndex { get; set; }
        public string ElementName { get; set; }

        public DataRow DataRow { get; set; }
        public Dictionary<string, string> GroupBy { get; set; }

        public string Name { get; set; }

        public bool ColumnNameForHead { get; set; }
        public int ColumnStart { get; set; }
        public int ColumnEnd { get; set; }

        public string GroupByString
        {
            get
            {
                if (GroupBy == null || GroupBy.Count == 0)
                {
                    return string.Empty;
                }

                string rtn = string.Empty;
                foreach (string key in this.GroupBy.Keys)
                {
                    rtn += "and " + key + " = '" + GroupBy[key] + "' ";
                }
                return rtn.Substring(3);
            }
        }

        public static string GetName(string elementName)
        {
            string[] element = elementName.Split('_');


            if (element[0].Equals("label"))
            {
                return element[2];
            }
            else
            {
                return element[4];
            }
        }
    }

    /// <summary>
    /// Table配置项
    /// </summary>
    public class TableConfig
    {
        public TableConfig(string tableDescr = "")
        {
            this.DataRow = 2;
            this.SummaryRow = -1;

            if (!string.IsNullOrEmpty(tableDescr))
            {
                string[] element = tableDescr.Split(',');
                foreach (string item in element)
                {
                    if (!string.IsNullOrEmpty(item))
                    {
                        string[] configs = item.Split(':');
                        if (configs.Length == 2)
                        {
                            switch (configs[0].ToLower())
                            {
                                case "data":
                                case "d":
                                    this.DataRow = int.Parse(configs[1]);
                                    break;
                                case "summary":
                                case "s":
                                    this.SummaryRow = int.Parse(configs[1]);
                                    break;
                                case "summaryfilter":
                                case "sf":
                                    this.SummaryFilter = configs[1];
                                    break;
                                default:
                                    break;
                            }
                        }
                    }
                }
            }

        }
        public int DataRow { get; set; }
        public int SummaryRow { get; set; }
        public string SummaryFilter { get; set; }
    }
}
