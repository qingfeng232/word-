using System;
using System.Collections.Generic;

namespace word插件
{
    public class ExcelValueDate
    {
        public string Value { get; set; }
        public int Count { get; set; }
        public List<int> ValueRows { get; set; } = new List<int>();
    }
    public class TableMap
    {
        public string WordHeader { get; set; }
        public string ExcelHeader { get; set; } // 允许为自定义内容
        public bool IsCustom { get; set; }      // 标记是否为自定义
    }

}
