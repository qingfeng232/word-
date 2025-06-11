using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.IO;
using System.Runtime.InteropServices;
using word插件;

namespace word插件
{
    //自定义一个类
    
    public class ExcelValueDate
    {
        public string Value { get; set; } // excel选中列的内容
        public int Count { get; set; }    // 该内容出现的次数
        public List<int> ValueRows { get; set; } = new List<int>(); //内容所在行数
    }
    internal class DataProcessor
    {
        // 对Excel的数据进行预处理
        /// <summary>
        /// 处理选中的列，统计每个唯一值、次数和所在行数
        /// </summary>
        /// <param name="allRows">Excel所有行，每行是一个字符串数组</param>
        /// <param name="Count">内容出现的次数</param>
        /// <param name="ValueRows">内容所在行</param>
        /// <returns>返回每个不同值的统计结果</returns>
        // 生成一个表，处理选中的列（selecteedExcelColumnName）遍历该列的每一行数据，提取出所有不同的值（value），记录值所在的行数，出现的次数
        public List<ExcelValueDate> ProcessColum(List<string[]> allRows, out int Count, out List<int> ValueRows)
        {
            // 初始化返回结果
            List<ExcelValueDate> result = new List<ExcelValueDate>();
            Count = 0;
            ValueRows = new List<int>();
            // 获取选中的列索引
            int columnIndex = Ribbon1.ExcelcolumnNames.IndexOf(Ribbon1.selecteedExcelColumnName);
            if (columnIndex < 0) return result; // 如果列不存在，返回空列表
            // 遍历所有行
            for (int i = Ribbon1.ExcelDateFirstRaw - 1; i < allRows.Count; i++)
            {
                string cellValue = allRows[i][columnIndex].Trim(); // 获取单元格值并去除空格
                if (string.IsNullOrEmpty(cellValue)) continue; // 如果单元格为空，跳过
                // 查找是否已经存在该值
                var existingEntry = result.FirstOrDefault(x => x.Value == cellValue);
                if (existingEntry != null)
                {
                    existingEntry.Count++; // 增加计数
                    existingEntry.ValueRows.Add(i + 1); // 添加行号（从1开始）
                }
                else
                {
                    // 创建新的统计项
                    ExcelValueDate newEntry = new ExcelValueDate
                    {
                        Value = cellValue,
                        Count = 1,
                        ValueRows = new List<int> { i + 1 } // 添加行号（从1开始）
                    };
                    result.Add(newEntry);
                }
            }
            Count = result.Count; // 设置总计数
            return result;
        }
        // 将处理结果保存为JSON文件
        public List<ExcelValueDate> SaveToJson(List<ExcelValueDate> data)
        {
            if (!Directory.Exists(Ribbon1.GenerateCatalogPath)) Directory.CreateDirectory(Ribbon1.GenerateCatalogPath);
            try
            {
                string fileName = $"Excel文件处理_{DateTime.Now:G}.json";
                string filePath = Path.Combine(Ribbon1.GenerateCatalogPath,fileName);
                // 将数据序列化为JSON格式
                string json = JsonConvert.SerializeObject(data, Formatting.Indented);
                // 写入文件
                File.WriteAllText(filePath, json);
                return data; // 返回处理结果
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Excel处理文件保存错误: {ex.Message}");
                return null;
            }
        }
     }
    //Excel与Word列映射的功能代码
    //获取Excel文件表头
    public class TableMap
    {
        public string ExcelHeader { get; set; }
        public string WordHeader { get; set; }
    }
}
