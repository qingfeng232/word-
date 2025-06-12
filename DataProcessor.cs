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
        public string ExcelHeader { get; set; }
        public string WordHeader { get; set; }
    }
}
