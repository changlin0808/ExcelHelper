using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelHelper
{
    public class Person
    {
        public string Name { get; set; }
        public string Age { get; set; }
        public string Email { get; set; }
    }
    class Program
    {
        static void Main(string[] args)
        {
            // 创建一个excel表格操作对象
            ExcelHelper helper = new ExcelHelper();
            // 读表初始化
            helper.InitRead("Person.xls", 0);
            // 读取指定单元格数据
            string value = helper.ReadCell(0, 0);
            // 将数据打印在控制台
            Console.Write(value + "\n");
            // 将对象资源销毁
            helper.Close();
        }
    }
}
