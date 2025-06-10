using Microsoft.Office.Tools.Ribbon;
using Microsoft.VisualBasic;
using System;
using System.Collections.Generic;
using System.Drawing.Text;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;


namespace word插件
{
    public partial class Ribbon1
    {
        // 定义全局变量
        private string selectedExcelPath = string.Empty,
                       selectedWordPath = string.Empty,
                       GenerateCatalogPath = string.Empty;


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
            int ExcelDateFirstRaw = 2; // 默认起始行
            string input = Microsoft.VisualBasic.Interaction.InputBox("请输入Excel数据起始行", "Excel数据起始行", "2", 2);
            if (int.TryParse(input, out int row))
            {
                ExcelDateFirstRaw = row;
                MessageBox.Show("更新起始位置"+ ExcelDateFirstRaw + "行", "更新成功" ,MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            int WordDateFirstRaw = 2; // 默认起始行
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
    }
          
 }   


        
    

