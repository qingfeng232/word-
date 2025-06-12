using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace word插件

{
    
    public class MappingForm : Form
    {
        public List<TableMap> MappingResult { get; set; } = new List<TableMap>();

        public MappingForm(List<string> wordHeaders, List<string> excelHeaders)
        {
            this.Text = "设置Word-Excel列映射";
            var panel = new TableLayoutPanel
            {
                RowCount = wordHeaders.Count + 1,
                ColumnCount = 3,
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(10)
            };

            panel.Controls.Add(new Label() { Text = "Word列", AutoSize = true }, 0, 0);
            panel.Controls.Add(new Label() { Text = "Excel列", AutoSize = true }, 1, 0);
            panel.Controls.Add(new Label() { Text = "自定义", AutoSize = true }, 2, 0);

            var comboList = new List<ComboBox>();
            var customBoxList = new List<TextBox>();
            for (int i = 0; i < wordHeaders.Count; i++)
            {
                panel.Controls.Add(new Label() { Text = wordHeaders[i], AutoSize = true }, 0, i + 1);

                var combo = new ComboBox
                {
                    DropDownStyle = ComboBoxStyle.DropDownList,
                    Width = 160
                };
                var options = excelHeaders.ToList();
                options.Insert(0, "无对应（自定义）");
                combo.DataSource = options;
                panel.Controls.Add(combo, 1, i + 1);
                comboList.Add(combo);

                var textBox = new TextBox
                {
                    Width = 120,
                    Enabled = false // 初始不可编辑
                };
                panel.Controls.Add(textBox, 2, i + 1);
                customBoxList.Add(textBox);

                // 事件：选中“无对应”时，文本框可编辑
                combo.SelectedIndexChanged += (s, e) =>
                {
                    if (combo.SelectedIndex == 0)
                    {
                        textBox.Enabled = true;
                        textBox.Focus();
                    }
                    else
                    {
                        textBox.Enabled = false;
                    }
                };
            }

            var btnOK = new Button { Text = "确定", Dock = DockStyle.Bottom };
            btnOK.Click += (s, e) =>
            {
                MappingResult.Clear();
                for (int i = 0; i < wordHeaders.Count; i++)
                {
                    var map = new TableMap();
                    map.WordHeader = wordHeaders[i];
                    if (comboList[i].SelectedIndex == 0)
                    {
                        map.ExcelHeader = customBoxList[i].Text?.Trim();
                        map.IsCustom = true;
                    }
                    else
                    {
                        map.ExcelHeader = comboList[i].SelectedItem?.ToString();
                        map.IsCustom = false;
                    }
                    MappingResult.Add(map);
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            };

            this.Controls.Add(panel);
            this.Controls.Add(btnOK);
            this.AutoSize = true;
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
        }
    }

}

