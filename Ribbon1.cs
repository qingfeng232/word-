using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using word插件;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Excel = Microsoft.Office.Interop.Excel;


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
        public static int           ExcelDateFirstRaw = 2,
                                     WordDateFirstRaw = 2;
        //读取的Excel的表头
        public static List<string> ExcelcolumnNames = new List<string>();
        //读取Excel表头
        private List<string> GetExcelColumnNames(string Path, int Datarow)
        {
            var columnNames = new List<string>();
            Excel.Application excelApp = new Excel.Application();
            Excel.Workbook workbook = excelApp.Workbooks.Open(Path);
            Excel.Worksheet worksheet = workbook.Sheets[1];

            int HeaderRow = Datarow - 1;
            int col = 1;
            // 从第一列开始
            while (true)
            {
                var cellValue = worksheet.Cells[HeaderRow, col].Value;
                if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
                    break; // 如果单元格为空，停止读取}
                columnNames.Add(cellValue.ToString());
                col++;
            }
            workbook.Close(false);
            excelApp.Quit();
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
            return columnNames;
        }
        //读取Word表头
        private List<string> GetWordColumnNames(string wordPath, int dataRow)
        {
            var result = new List<string>();
            var wordApp = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = null;
            try
            {
                doc = wordApp.Documents.Open(wordPath, ReadOnly: true);
                var table = doc.Tables[1]; // 假设只取第一个表
                int headerRow = dataRow - 1;
                for (int col = 1; col <= table.Columns.Count; col++)
                {
                    var cellValue = table.Cell(headerRow, col).Range.Text;
                    cellValue = cellValue?.Replace("\r\a", "").Trim(); // 去除Word单元格特殊符号
                    result.Add(cellValue);
                }
            }
            finally
            {
                if (doc != null) doc.Close(false);
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }
            return result;
        }
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
            string input = Microsoft.VisualBasic.Interaction.InputBox("请输入Excel数据起始行", "Excel数据起始行", "2", 2);
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
            ComboBox1.Items.Clear();
            if (string.IsNullOrEmpty(selectedExcelPath))
            {
                MessageBox.Show("未选择Excel文件");
                return;
             }
            //刷新Excel表头
            ExcelcolumnNames = GetExcelColumnNames(selectedExcelPath, ExcelDateFirstRaw); // 你自定义的函数
            if(ExcelcolumnNames.Count == 0)
            {
                MessageBox.Show("未读取表头");
                return;
            }
            foreach (var name in ExcelcolumnNames)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = name;
                ComboBox1.Items.Add(item);
            }
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
            var mapForm = new MappingForm(wordHeaders,excelHeaders);
            if (mapForm.ShowDialog() == DialogResult.OK)
            {
                List<TableMap> mapping = mapForm.MappingResult;
                // 你可以在这里将mapping存到全局变量或导出为文件，供后续批量填充用
                // 例如:
                 this.CurrentMapping = mapping;
                MessageBox.Show("映射关系设置完成！");
            }
        }
        //生成文件
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {   
            string[] files = Directory.GetFiles(GenerateCatalogPath, "Excel统计_*.json");
            if (files.Length == 0)
            {
                MessageBox.Show("未找到任何统计JSON文件，请先执行数据统计！");
                return;
            }
            string statFile = files.OrderByDescending(f => File.GetLastWriteTime(f)).First();

            // 2. 读取Excel原数据
            var groupList = JsonConvert.DeserializeObject<List<ExcelValueDate>>(File.ReadAllText(statFile));
            var excelHeaders = GetExcelColumnNames(selectedExcelPath, ExcelDateFirstRaw);
            var wordHeaders = GetWordColumnNames(selectedWordPath, WordDateFirstRaw);
            // 3. 映射关系
            var allRows = ReadExcelAllRows(selectedExcelPath);

            if (CurrentMapping == null || CurrentMapping.Count == 0)
            {
                MessageBox.Show("未设置映射关系！"); return;
            }
            var mapDict = CurrentMapping.ToDictionary(x => x.WordHeader, x => x.ExcelHeader);
            // 4. 让用户输入每页填多少行
            string input = Microsoft.VisualBasic.Interaction.InputBox("每页填充数据行数（如13）：", "设置", "13");
            int rowsPerPage = 13;
            if (!int.TryParse(input, out rowsPerPage) || rowsPerPage <= 0) rowsPerPage = 13;

            // 5. 每个唯一Value生成Word文件
            foreach (var group in groupList)
            {
                string value = group.Value;
                string outputPath = Path.Combine(GenerateCatalogPath, $"{value}.docx");
                File.Copy(selectedWordPath, outputPath, true);

                var wordApp = new Microsoft.Office.Interop.Word.Application();
                var doc = wordApp.Documents.Open(outputPath, ReadOnly: false, Visible: false);
                var table = doc.Tables[1];
                int wordDataStart = WordDateFirstRaw;
                int dataIdx = 0;
                int totalRows = group.ValueRows.Count;
                int curTableRow = wordDataStart;

                // 先统计第一页、第二页数据区行数
                int secondPageStart = wordDataStart + rowsPerPage; // 第二页数据区起始行
                int secondPageEnd = secondPageStart + rowsPerPage - 1;


                input = Microsoft.VisualBasic.Interaction.InputBox("每页填充数据行数（如13）：", "设置", rowsPerPage.ToString());
                if (!int.TryParse(input, out rowsPerPage) || rowsPerPage <= 0) rowsPerPage = 13;

                // 主循环：每次填满一页，不足补空行
                while (dataIdx < totalRows)
                {
                    // 若当前页的数据区不存在，复制第二页的数据区到表尾
                    if (curTableRow + rowsPerPage - 1 > table.Rows.Count)
                    {
                        // 复制第二页数据区
                        for (int i = 0; i < rowsPerPage; i++)
                        {
                            int srcRow = secondPageStart + i;
                            if (srcRow > table.Rows.Count) break;
                            table.Rows.Add();
                            int newRowIdx = table.Rows.Count;
                            for (int c = 1; c <= table.Columns.Count; c++)
                            {
                                table.Cell(newRowIdx, c).Range.FormattedText =
                                    table.Cell(srcRow, c).Range.FormattedText.Duplicate;
                                table.Cell(newRowIdx, c).Range.Text = ""; // 清空内容
                            }
                        }
                    }

                    // 填一页数据
                    for (int wordRow = curTableRow; wordRow < curTableRow + rowsPerPage; wordRow++)
                    {
                        // 填充数据或补空
                        if (dataIdx < totalRows)
                        {
                            int excelRowIdx = group.ValueRows[dataIdx] - 1;
                            if (excelRowIdx >= allRows.Count) { dataIdx++; continue; }
                            var excelRow = allRows[excelRowIdx];

                            for (int c = 1; c <= wordHeaders.Count; c++)
                            {
                                string wordHeader = wordHeaders[c - 1];
                                if (!mapDict.TryGetValue(wordHeader, out string excelHeader)) continue;
                                int excelColIdx = excelHeaders.IndexOf(excelHeader);
                                if (excelColIdx < 0 || excelColIdx >= excelRow.Length) continue;
                                table.Cell(wordRow, c).Range.Text = excelRow[excelColIdx];
                            }
                            dataIdx++;
                        }
                        else
                        {
                            // 补空行
                            for (int c = 1; c <= table.Columns.Count; c++)
                                table.Cell(wordRow, c).Range.Text = "";
                        }
                    }
                    curTableRow += rowsPerPage;
                }

                doc.Save();
                doc.Close(false);
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }

            MessageBox.Show("批量生成Word文件已完成！");
        }

    }

}
