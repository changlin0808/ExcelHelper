using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;

namespace ExcelHelper
{
    class ExcelHelper
    {
        // 读excel的初始化
        public void InitRead(string path, int sheet_num)
        {
            //1、创建一个工作簿workBook对象
            fStream = new FileStream(path, FileMode.Open);
            //将人员表.xls中的内容读取到fStream中
            wkBook = new HSSFWorkbook(fStream);
            sheet = wkBook.GetSheetAt(sheet_num);
        }

        // 写excel操作的初始化
        public void InitWrite(string path)
        {
            wkBook = new HSSFWorkbook();
            //2、在该工作簿中创建工作表对象
            sheet = wkBook.CreateSheet(); 
            fStream = File.OpenWrite(path);   
        }
        
        // 从指定单元格读数据
        public string ReadCell(int row, int column)
        {
            //获取工作表中的每一行
            IRow currentRow = sheet.GetRow(row);
            // 获得该行的某个单元格
            ICell cell = currentRow.GetCell(column);
            string str = cell.StringCellValue;
            return str;
        }

        // 向指定单元格写数据
        public void WriteCell(int row, int column, string value)
        {
            IRow row_obj = sheet.CreateRow(row);
            row_obj.CreateCell(column).SetCellValue(value);
            wkBook.Write(fStream);
        }

        // 关闭工作流
        public void Close()
        {
            fStream.Close(); //关闭文件流
            wkBook.Close();  //关闭工作簿
            fStream.Dispose(); //释放文件流
        }

        private FileStream fStream = null;
        private IWorkbook wkBook = null;
        private ISheet sheet = null;
    }
}
