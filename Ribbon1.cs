using Microsoft.Office.Tools.Ribbon;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Messaging;
using System.Threading;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using Word  = Microsoft.Office.Interop.Word;
namespace word插件
{
    public partial class Ribbon1
    {
        // 定义全局变量
        public static string        selectedExcelPath = string.Empty,
                                     selectedWordPath = string.Empty,
                                  GenerateCatalogPath = string.Empty,
                             selecteedExcelColumnName = string.Empty;
        // 定义Excel和Word的起始行
        public static int ExcelDateFirstRaw = 2,
                                     WordDateFirstRaw = 3;
        //读取的Excel的表头
        public static List<string> ExcelcolumnNames = new List<string>();
        //读取Excel表头
        private List<string> GetExcelColumnNames(string path, int dataRow)
        {
            var result = new List<string>();//存储表头名称
            Excel.Application app = null;//Excel应用程序对象
            Excel.Workbook wb = null;//工作簿对象
            Excel.Worksheet ws = null;//工作表对象
            Excel.Range used = null, headerRow = null;//用于存储使用范围和表头行的范围

            try
            {
                app = new Excel.Application { Visible = false, ScreenUpdating = false, DisplayAlerts = false };//初始化Excel应用程序
                wb = app.Workbooks.Open(path, ReadOnly: true);//打开指定路径的工作簿
                ws = (Excel.Worksheet)wb.Sheets[1];//获取第一个工作表

                int headerRowIndex = Math.Max(1, dataRow - 1);
                used = ws.UsedRange;//获取工作表的使用范围
                int totalCols = used.Columns.Count;//获取总列数

                headerRow = ws.Range[ws.Cells[headerRowIndex, 1], ws.Cells[headerRowIndex, totalCols]];//获取表头行的范围
                object[,] values = headerRow.Value2 as object[,];//将表头行的值存储为二维数组

                for (int c = 1; c <= totalCols; c++)
                {
                    var v = values[1, c]; // 单行区域的第一维固定为1
                    string text = v?.ToString().Trim();
                    // 跳过空白，但不提前停止
                    result.Add(string.IsNullOrEmpty(text) ? "" : text);
                }

                // 去掉尾部全空列（可选）
                for (int i = result.Count - 1; i >= 0; i--)
                {
                    if (!string.IsNullOrEmpty(result[i])) break;
                    result.RemoveAt(i);
                }

                return result;
            }
            //释放缓存
            finally
            {
                if (headerRow != null) Marshal.ReleaseComObject(headerRow);
                if (used != null) Marshal.ReleaseComObject(used);
                if (ws != null) Marshal.ReleaseComObject(ws);
                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    Marshal.ReleaseComObject(wb);
                }
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    Marshal.ReleaseComObject(app);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        
        //读取Word表头
        private List<string> GetWordColumnNames(string wordPath, int dataRow)
        {
            var result = new List<string>();
            Word.Application wordApp = new Word.Application();
            Word.Document doc = null;
            Word.Table  table = null;
            try
            {
                doc   = wordApp.Documents.Open(wordPath, ReadOnly: true);
                if(doc.Tables.Count<1) return result;//没有表格直接返回空

                table = doc.Tables[1]; // 只取第一个表
                int headerRow = Math.Max(1, dataRow - 1);
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    string cellValue = table.Cell(headerRow, col).Range.Text;
                    cellValue = cellValue?.Replace("\r\a", "").Trim(); // 去除Word单元格特殊符号
                    result.Add(cellValue ?? "");//添加到结果列表
                }
            }
            finally//释放缓存
            {
                if(table !=null)
                {
                    SafeRelease(ref table);
                }
                if (doc != null)
                {
                    try { doc.Close(false); } catch { }
                    SafeRelease(ref doc);
                }
                if (wordApp != null)
                {
                    try { wordApp.Quit(); } catch { }
                    SafeRelease(ref wordApp);
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
            return result;
        }
        //刷新Excel表头
        private void RefreshExcelColumnComboBox()
        {
            ComboBox1.Items.Clear();

            if (string.IsNullOrEmpty(selectedExcelPath))
            {
                // 可选：MessageBox.Show("请先选择Excel文件！");
                return;
            }

            ExcelcolumnNames = GetExcelColumnNames(selectedExcelPath, ExcelDateFirstRaw);

            if (ExcelcolumnNames.Count == 0)
            {
                MessageBox.Show("未读取到表头，请检查Excel和起始行设置！");
                return;
            }

            foreach (var name in ExcelcolumnNames)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = name;
                ComboBox1.Items.Add(item);
            }
        }
        public object Private { get; private set; }
        public List<TableMap> CurrentMapping { get; private set; }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        //选择Excel文件
        private void SelectExcelButton_Click(object sender, RibbonControlEventArgs e)
        {
            // 使用 WinForms 的文件选择对话框
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "选择一个 Excel 文件",
                Filter = "Excel 文件 (*.xlsx;*.xls)|*.xlsx;*.xls",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };

            // 如果用户选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedExcelPath = openFileDialog.FileName;

                // 弹出确认窗口
                MessageBox.Show("你选择的 Excel 文件是：\n" + selectedExcelPath, "文件选择成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            // 用户取消了选择
            else
            {

                MessageBox.Show("没有选择任何文件。", "文件选择取消",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }

            // Excel文件的数据处理
            string input = Microsoft.VisualBasic.Interaction.InputBox("请输入Excel数据起始行", "Excel数据起始行", "2", 2);
            if (int.TryParse(input, out int row))
            {
                ExcelDateFirstRaw = row;
                MessageBox.Show("更新起始位置" + ExcelDateFirstRaw + "行", "更新成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("输入无效，请输入一个数字。", "使用默认第二行", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            RefreshExcelColumnComboBox();
        }
        //选择Word文件
        private void SelectWordButton_Click(object sender, RibbonControlEventArgs e)
        {
            // 使用 WinForms 的文件选择对话框
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = "选择一个 Word 文件",
                Filter = "Word 文件 (*.docx;*doc)|*.docx;*.doc",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            };
            // 如果用户选择了文件
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedWordPath = openFileDialog.FileName;
                GenerateCatalogPath = Path.GetDirectoryName(openFileDialog.FileName);

                // 弹出确认窗口
                MessageBox.Show("你选择的 Excel 文件是：\n" + selectedWordPath, "文件选择成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            // 用户取消了选择
            else
            {

                MessageBox.Show("没有选择任何文件。", "文件选择取消",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);

            }
            // Word文件的数据处理
            string input = Microsoft.VisualBasic.Interaction.InputBox("请输入Excel数据起始行", "Excel数据起始行", "3", 3);
            if (int.TryParse(input, out int row))
            {
                WordDateFirstRaw = row;
                MessageBox.Show("更新起始位置" + WordDateFirstRaw + "行", "更新成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("输入无效，请输入一个数字。", "使用默认第二行", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }
        //选择生成文件夹
        private void GenerateCatalog_Click(object sender, RibbonControlEventArgs e)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog
            {
                Description = "选择生成文件的目录",
                ShowNewFolderButton = true,
                SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)

            };
            if (folderBrowserDialog.ShowDialog() == DialogResult.OK)
            {
                GenerateCatalogPath = folderBrowserDialog.SelectedPath;
                MessageBox.Show("生成目录已选择：" + GenerateCatalogPath, "目录选择成功",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("没有选择任何目录。", "目录使用Word模板所在目录",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
        //遍历Excel文件表头
        private void ComboBox1_ItemsLoading(object sender, RibbonControlEventArgs e)
        {

        }
        //选择Excel参考列
        private void ComboBox1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            string selectedItem = ComboBox1.Text; // 获取选中的项
            selecteedExcelColumnName = selectedItem; // 将选中的项赋值给全局变量
        }
        //数据处理
        private List<string[]> ReadExcelAllRows(string excelPath)
        {
            var allRows = new List<string[]>();
            var excelApp = new Excel.Application();
            Excel.Workbook wb = null;
            try
            {
                wb = excelApp.Workbooks.Open(excelPath);
                Excel.Worksheet ws = wb.Sheets[1];
                int totalRows = ws.UsedRange.Rows.Count;
                int totalCols = ws.UsedRange.Columns.Count;

                for (int i = 1; i <= totalRows; i++)
                {
                    var row = new List<string>();
                    for (int j = 1; j <= totalCols; j++)
                    {
                        var val = ws.Cells[i, j].Value;
                        row.Add(val?.ToString() ?? "");
                    }
                    allRows.Add(row.ToArray());
                }
            }
            finally
            {
                if (wb != null) wb.Close(false);
                excelApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            }
            return allRows;
        }
        //读取选择的Excel文件，以comboBox1_SelectionChanged获取的列为依据，生成.json文件，生成4个变量。TableData列中的内容，Count不同内容的个数，ValueCount同一个内容的数量，ValueRows，内容所在的行数。
        private void ExcelDataProcessing_Click(object sender, RibbonControlEventArgs e)
        {
            // 1. 基础校验
            if (string.IsNullOrWhiteSpace(selectedExcelPath))
            {
                MessageBox.Show("请先选择 Excel 文件！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (string.IsNullOrWhiteSpace(selecteedExcelColumnName))
            {
                MessageBox.Show("请先选择要统计的列！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // 2. 读取Excel所有行
            var allRows = ReadExcelAllRows(selectedExcelPath);
            if (allRows == null || allRows.Count == 0)
            {
                MessageBox.Show("Excel数据读取失败。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 3. 查找目标列索引
            int colIndex = ExcelcolumnNames.IndexOf(selecteedExcelColumnName);
            if (colIndex < 0)
            {
                MessageBox.Show("未找到所选列，请刷新表头下拉框。", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // 4. 统计
            var tableData = new List<ExcelValueDate>();
            Dictionary<string, ExcelValueDate> valueDict = new Dictionary<string, ExcelValueDate>();
            for (int i = ExcelDateFirstRaw - 1; i < allRows.Count; i++)
            {
                if (allRows[i].Length <= colIndex) continue;
                string cellValue = allRows[i][colIndex]?.Trim() ?? "";
                if (string.IsNullOrEmpty(cellValue)) continue;

                if (!valueDict.TryGetValue(cellValue, out ExcelValueDate entry))
                {
                    entry = new ExcelValueDate
                    {
                        Value = cellValue,
                        Count = 1,
                        ValueRows = new List<int> { i + 1 } // Excel 行号1-based
                    };
                    valueDict[cellValue] = entry;
                    tableData.Add(entry);
                }
                else
                {
                    entry.Count++;
                    entry.ValueRows.Add(i + 1);
                }
            }

            // 5. 生成变量
            int Count = tableData.Count; // 不同内容的个数
                                         // ValueCount: 各内容的数量列表
            List<int> ValueCount = tableData.Select(x => x.Count).ToList();
            // ValueRows: 各内容的行号列表（可选也可合并为字典/数组）
            List<List<int>> ValueRows = tableData.Select(x => x.ValueRows).ToList();

            // 6. 生成JSON
            if (!Directory.Exists(GenerateCatalogPath))
                Directory.CreateDirectory(GenerateCatalogPath);
            string fileName = $"Excel统计_{selecteedExcelColumnName}_{DateTime.Now:yyyyMMdd_HHmmss}.json";
            string filePath = Path.Combine(GenerateCatalogPath, fileName);

            // json数据结构：tableData即为所有信息
            File.WriteAllText(filePath, JsonConvert.SerializeObject(tableData, Formatting.Indented));

            // 7. 提示
            MessageBox.Show(
                $"已完成统计并导出JSON。\n\n" +
                $"列：{selecteedExcelColumnName}\n" +
                $"不同内容数量：{Count}\n" +
                $"JSON文件：\n{filePath}",
                "统计成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        //打开映射窗口
        private void BtnSetMapping_Click(object sender, RibbonControlEventArgs e)
        {
            // 1. 检查Excel、Word路径
            if (string.IsNullOrWhiteSpace(selectedExcelPath) || string.IsNullOrWhiteSpace(selectedWordPath))
            {
                MessageBox.Show("请先选择Excel和Word文件！", "提示");
                return;
            }
            // 2. 读取Word表头
            var wordHeaders = GetWordColumnNames(selectedWordPath, WordDateFirstRaw);
            // 3. 读取Excel表头
            var excelHeaders = GetExcelColumnNames(selectedExcelPath, ExcelDateFirstRaw);

            if (excelHeaders.Count == 0 || wordHeaders.Count == 0)
            {
                MessageBox.Show("表头读取失败，请确认Excel和Word文件。", "错误");
                return;
            }

            // 4. 映射弹窗
            var mapForm = new MappingForm(wordHeaders, excelHeaders);
            if (mapForm.ShowDialog() == DialogResult.OK)
            {
                List<TableMap> mapping = mapForm.MappingResult;
                this.CurrentMapping = mapping;
                MessageBox.Show("映射关系设置完成！");
            }
        }
        //释放对象
        private void SafeRelease<T>(ref T obj)
        {
            if (obj != null && Marshal.IsComObject(obj))
            {
                Marshal.ReleaseComObject(obj);
                obj = default;
            }
        }
        //查找最后一行数据
        private int FindLastValidDataRow(Word.Table table, int startRow)
        {
            for (int i = table.Rows.Count; i >= startRow; i--)
            {
                if (table.Rows[i].Cells.Count == table.Columns.Count)
                    return i;
            }
            return startRow;
        }
        //填充 Word 表格行
        private void FillTableRow(
                                   Word.Table table,
                                   int row,
                                   int excelRowIdx,
                                   List<string[]> allRows,// Excel所有行数据
                                   List<string> excelHeaders,
                                   List<string> wordHeaders,
                                   Dictionary<string, string> mapDict,
                                   Dictionary<string, string> customDict) // 新增自定义内容表
        {
            var excelRow = allRows[excelRowIdx];

            for (int c = 1; c <= table.Columns.Count; c++)
            {
                string wordHeader = wordHeaders[c - 1];

                // 优先处理自定义内容
                if (customDict.TryGetValue(wordHeader, out string staticValue))
                {
                    try { table.Cell(row, c).Range.Text = staticValue; }
                    catch { }
                    continue;
                }

                // 其次处理 Excel 映射列
                if (!mapDict.TryGetValue(wordHeader, out string excelHeader)) continue;

                int idx = excelHeaders.IndexOf(excelHeader);
                if (idx >= 0 && idx < excelRow.Length)
                {
                    try { table.Cell(row, c).Range.Text = excelRow[idx]; }
                    catch { }
                }
            }
        }

        //清空 Word 表格行
        private void ClearTableRow(Word.Table table, int row)
        {
            for (int c = 1; c <= table.Columns.Count; c++)
            {
                try { table.Cell(row, c).Range.Text = ""; }
                catch { }
            }
        }
        //粘贴模板页的表格
        private bool TryCopyTemplateTable(Word.Application wordApp, string templatePath)
        {
            Word.Document tempDoc = null;
            try
            {
                tempDoc = wordApp.Documents.Open(templatePath, ReadOnly: true, Visible: false);
                if (tempDoc.Tables.Count < 2)
                {
                    tempDoc.Close(false);
                    SafeRelease(ref tempDoc);
                    return false;
                }

                Word.Table tempTable = tempDoc.Tables[2];
                tempTable.Range.Copy();

                tempDoc.Close(false);
                SafeRelease(ref tempTable);
                SafeRelease(ref tempDoc);
                return true;
            }
            catch
            {
                if (tempDoc != null)
                {
                    tempDoc.Close(false);
                    SafeRelease(ref tempDoc);
                }
                return false;
            }
        }
        // 批量生成Word文件
        private void Button1_Click(object sender, RibbonControlEventArgs e)
        {
            string[] files = Directory.GetFiles(GenerateCatalogPath, "Excel统计_*.json");
            if (files.Length == 0)
            {
                MessageBox.Show("未找到统计文件");
                return;
            }
            string statFile = files.OrderByDescending(f => File.GetLastWriteTime(f)).First();

            var groupList = JsonConvert.DeserializeObject<List<ExcelValueDate>>(File.ReadAllText(statFile));
            var allRows = ReadExcelAllRows(selectedExcelPath);
            var excelHeaders = GetExcelColumnNames(selectedExcelPath, ExcelDateFirstRaw);
            var wordHeaders = GetWordColumnNames(selectedWordPath, WordDateFirstRaw);
            var customDict = CurrentMapping.Where(x => x.IsCustom && !string.IsNullOrWhiteSpace(x.InputContent)).ToDictionary(x => x.WordHeader, x => x.InputContent);

            if (CurrentMapping == null || CurrentMapping.Count == 0)
            {
                MessageBox.Show("未设置映射关系！");
                return;
            }
            var mapDict = CurrentMapping
            .Where(x => !string.IsNullOrWhiteSpace(x.ExcelHeader))
            .ToDictionary(x => x.WordHeader, x => x.ExcelHeader);


            string input1 = Microsoft.VisualBasic.Interaction.InputBox("第一页填充数据行数（如11）：", "设置", "11");
            string input2 = Microsoft.VisualBasic.Interaction.InputBox("后续页填充数据行数（如18）：", "设置", "18");
            int rowsPerPage1 = int.TryParse(input1, out int r1) && r1 > 0 ? r1 : 11;//第一页填入行数
            int rowsPerPage2 = int.TryParse(input2, out int r2) && r2 > 0 ? r2 : 18;//第二页填入行数
            //读取映射信息
            foreach (var map in CurrentMapping)
            {
                Debug.WriteLine($"Word列: {map.WordHeader} => Excel列: {map.ExcelHeader}, IsCustom: {map.IsCustom}, Input: {map.InputContent}");
            }

            foreach (var group in groupList)

            {   //初始化
                Word.Application wordApp = null;
                Word.Document doc = null;
                Word.Table tableFirst = null;

                try
                {
                    string outputPath = Path.Combine(GenerateCatalogPath, $"{group.Value}.docx");//获取模板信息
                    //复制模板文件
                    File.Copy(selectedWordPath, outputPath, true);
                    Thread.Sleep(100);//等待100毫秒

                    wordApp = new Word.Application();
                    doc = wordApp.Documents.Open(outputPath, ReadOnly: false, Visible: false);//

                    if (doc.Tables.Count < 1)
                    {
                        MessageBox.Show($"模板中找不到表格，跳过 {group.Value}");
                        continue;
                    }

                    tableFirst = doc.Tables[1];
                    int dataIdx = 0;
                    int totalRows = group.ValueRows?.Count ?? 0;

                    int dataStartRow = WordDateFirstRaw;
                    int dataEndRow = FindLastValidDataRow(tableFirst, dataStartRow);

                    for (int row = dataStartRow; row <= dataEndRow && dataIdx < totalRows; row++)
                    {
                        int excelRowIdx = group.ValueRows[dataIdx] - 1;
                        if (excelRowIdx >= allRows.Count) { dataIdx++; continue; }
                        FillTableRow(tableFirst, row, excelRowIdx, allRows, excelHeaders, wordHeaders, mapDict, customDict);
                        dataIdx++;
                    }
                    //日志
                    Debug.WriteLine($"正在处理文档：{group.Value}");
                    Debug.WriteLine($"对应行数：{string.Join(",", group.ValueRows)}");
                    for (int row = dataStartRow + dataIdx; row <= dataEndRow; row++)
                        ClearTableRow(tableFirst, row);
                    //第二页不填入，删除第二页的表格。
                    /*while (dataIdx>=totalRows)
                     {
                         doc.Tables[2].Delete(); 
                     }*/
                    // 后续页
                    while (dataIdx < totalRows)
                    {
                        if (!TryCopyTemplateTable(wordApp, selectedWordPath))
                        {
                            MessageBox.Show("复制模板页失败！");
                            break;
                        }

                        Debug.WriteLine($"触发分页逻辑：dataIdx = {dataIdx}, total = {totalRows}");

                        // 插入分页符并粘贴表格
                        Word.Range endRange = doc.Content;
                        object collapseEnd = Word.WdCollapseDirection.wdCollapseEnd;
                        object breakType = Word.WdBreakType.wdPageBreak;
                        endRange.Collapse(ref collapseEnd);
                        endRange.InsertBreak(ref breakType);
                        endRange.Paste();
                        // 等待新表格被插入
                        int oldTableCount = doc.Tables.Count - 1;
                        int tryCount = 0;
                        while (doc.Tables.Count <= oldTableCount && tryCount++ < 20)
                            Thread.Sleep(100);

                        if (doc.Tables.Count <= oldTableCount)
                        {
                            MessageBox.Show("粘贴表失败！");
                            break;
                        }

                        // 正确获取最新插入的分页表格（最后一个表）
                        Word.Table newTable = doc.Tables[doc.Tables.Count];

                        int newStart = WordDateFirstRaw;
                        int newEnd = FindLastValidDataRow(newTable, newStart);

                        for (int row = newStart; row <= newEnd && dataIdx < totalRows; row++)
                        {
                            int excelRowIdx = group.ValueRows[dataIdx] - 1;
                            if (excelRowIdx >= allRows.Count) { dataIdx++; continue; }

                            // 正确使用 newTable
                            FillTableRow(newTable, row, excelRowIdx, allRows, excelHeaders, wordHeaders, mapDict, customDict);
                            dataIdx++;
                        }

                        for (int row = newStart + dataIdx; row <= newEnd; row++)
                            ClearTableRow(newTable, row);

                        SafeRelease(ref newTable);
                    }

                    doc.Tables[2].Delete();//删除第二页表格
                    doc.Save();
                    doc.Close(false);
                    wordApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"生成文档 {group.Value} 失败：{ex.Message}");
                }
            }
            MessageBox.Show("批量生成Word文件已完成！");
        }

    }
}