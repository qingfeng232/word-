using Microsoft.Office.Tools.Ribbon;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using word插件;


namespace word插件
{
    public partial class Ribbon1
    {
        // 定义全局变量
        public static string selectedExcelPath = string.Empty,
                                     selectedWordPath = string.Empty,
                                  GenerateCatalogPath = string.Empty,
                             selecteedExcelColumnName = string.Empty;
        // 定义Excel和Word的起始行
        public static int ExcelDateFirstRaw = 2,
                                    WordDateFirstRaw = 2;
        //读取的Excel的表头
        public static List<string> ExcelcolumnNames = new List<string>();

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
        public object Private { get; private set; }
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
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
        }
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
        private void comboBox1_ItemsLoading(object sender, RibbonControlEventArgs e)
        {
            comboBox1.Items.Clear();


            ExcelcolumnNames = GetExcelColumnNames(selectedExcelPath, ExcelDateFirstRaw); // 你自定义的函数

            foreach (var name in ExcelcolumnNames)
            {
                RibbonDropDownItem item = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem();
                item.Label = name;
                comboBox1.Items.Add(item);
            }
        }
        private void comboBox1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            string selectedItem = comboBox1.Text; // 获取选中的项
            selecteedExcelColumnName = selectedItem; // 将选中的项赋值给全局变量
        }
        private void ExcelDataProcessing_Click(object sender, RibbonControlEventArgs e)
        {      
            List<string[]> allRows = new List<string[]>(); 
            int count;
            List<int> valueRows;
            var processor = new DataProcessor();
            List<ExcelValueDate> processedData = processor.ProcessColum(allRows, out count, out valueRows);
            processor.SaveToJson(processedData);
        }
    }
}