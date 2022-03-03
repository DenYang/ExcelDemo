using System;
using System.Data;
using System.IO;
using System.Text;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.HPSF;
using NPOI.HSSF.Util;
using System.Reflection;
using System.Collections.Generic;
using NPOI.POIFS.FileSystem;
using System.Collections;

namespace ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("代码开始执行...");
            InsertData("C:/Users/cn-yangzheng/Desktop/测试.xls",12);
            //copyRange();
            string filePath = @"C:\Users\cn-yangzheng\Desktop\测试.xls";
            
            //ModelToExcel(filePath,12,3,data);
            //System.Data.DataTable dt = ReadExcel();
            //WriteExcel(dt);
            Console.WriteLine("代码执行完毕！");
            //System.Data.DataTable dt= ReadExcel();
            //SaveCsv(dt, "C:/Users/cn-yangzheng/Desktop/");
        }
        /*
         rowCount代表传入的多少行数据
         */
        public static void InsertData(string filePath,int rowCount)
        {
            int rowHead = 3;
            IWorkbook book;
            using(FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                book = new HSSFWorkbook(fs);
                ISheet sheet = book.GetSheetAt(0);//获取sheet
                IRow row ;
                for (int x = rowHead + 1; x < rowHead + rowCount; x++)
                {
                    row = sheet.CreateRow(x);
                    sheet.CopyRow(rowHead, x);
                }
                FileStream out2 = File.OpenWrite(filePath);
                book.Write(out2);
                out2.Close();
                fs.Close();
            }
            /*app.Visible = false;
            int selectNum,insertNum;
            Excel.Workbooks books = app.Workbooks;
            Excel.Workbook book = books.Open(@"C:\Users\cn-yangzheng\Desktop\多表数据 - 副本.xls");
            Excel.Sheets sts = book.Worksheets;
            Excel.Worksheet st = sts.Item[3];//sts.Item[1 - SheetCount]
            Console.WriteLine("请输入选定行数：...");
            selectNum = ReadInt();
            Excel.Range rng = st.Application.Rows[selectNum];//返回的是一个range类型，选定第5行数据
            rng.Select();
            Console.WriteLine("请输入插入行数：...");
            insertNum = ReadInt();
            rng.Select();
            rng.Resize[insertNum].Insert();//进行1行插入
            app.DisplayAlerts = false;
            book.Save();
            book.Close();
           */
        }

        public static void CopyRange(string filepath)
        {
            IWorkbook workbook;
            ISheet sheet;
        
            //string filepath = "C:/Users/cn-yangzheng/Desktop/多表数据 - 副本.xls";
            FileStream fs = new FileStream(filepath, FileMode.Open);
            workbook = new HSSFWorkbook(fs);
            sheet = workbook.GetSheetAt(0);
            Console.WriteLine("请输入开始复制的行数：");
            int startRow = ReadInt();
            Console.WriteLine("请输入结束复制的行数：");
            int endRow = ReadInt();
            Console.WriteLine("请输入复制到的行数：");
            int copyToRow = ReadInt();
            int rowNum = endRow - startRow;
            for(int i = startRow-1; i <= endRow-1; i++)
            {
                sheet.CopyRow(i, copyToRow-1);
                sheet.RemoveRow(sheet.GetRow(i));
            }
            FileStream out2 = new FileStream(filepath, FileMode.Create);
            workbook.Write(out2);
            out2.Close();
            fs.Close();
        }
        public static int ReadInt()
        {
            int number = 0;
            do
            {
                try
                {
                    //将根据提示输入的数字字符串转换成int型   
                    //Console.ReadLine(),这个函数，是以回车判断字符串结束的  
                    //number = Convert.ToInt32(Console.ReadLine());//与下面的效果一样  
                    number = System.Int32.Parse(Console.ReadLine());
                    if(number == 0&&number<0) { return -1; }
                    else {return number; }
                }
                catch
                {
                    Console.WriteLine("输入有误，重新输入！");
                }
            }
            while (true);
        }
        public static void WriteExcel(System.Data.DataTable dt)
        {
            bool flag = ExcelUtility.DataTableToExcel(dt);
            Console.WriteLine(flag);
        }
        public static System.Data.DataTable ReadExcel()
    {
            System.Data.DataTable dt = null;
            string path  = "C:/Users/cn-yangzheng/Desktop/多表数据 - 副本.xls";
            dt = ExcelUtility.ExcelToDataTable(path, true);
            return dt;
    }
        public class ExcelUtility
        {
            /// <summary>  
            /// 将excel导入到datatable  
            /// </summary>  
            /// <param name="filePath">excel路径</param>  
            /// <param name="isColumnName">第一行是否是列名</param>  
            /// <returns>返回datatable</returns>  
            public static System.Data.DataTable ExcelToDataTable(string filePath, bool isColumnName)
            {
                System.Data.DataTable dataTable = null;
                FileStream fs = null;
                DataColumn column = null;
                DataRow dataRow = null;
                IWorkbook workbook = null;
                ISheet sheet = null;
                IRow row = null;
                ICell cell = null;
                int startRow = 3;
                try
                {
                    using (fs = File.OpenRead(filePath))
                    {
                        // 2003版本  
                        if (filePath.IndexOf(".xls") > 0)
                            workbook = new HSSFWorkbook(fs);
                        // 2007版本
                        else if (filePath.IndexOf(".xlsx") > 0)
                            workbook = new XSSFWorkbook(fs);

                        if (workbook != null)
                        {
                            sheet = workbook.GetSheetAt(2);//读取第一个sheet，当然也可以循环读取每个sheet  
                            dataTable = new System.Data.DataTable();
                            if (sheet != null)
                            {
                                int rowCount = sheet.LastRowNum;//总行数  
                                if (rowCount > 0)
                                {
                                    IRow firstRow = sheet.GetRow(1);//第二行  
                                    int cellCount = firstRow.LastCellNum;//列数  
                                    

                                    //构建datatable的列  
                                    if (isColumnName)
                                    {
                                        //Console.WriteLine(firstRow.FirstCellNum);
                                        startRow =2 ;//如果第一行是列名，则从第二行开始读取  

                                        for (int i = firstRow.FirstCellNum; i < cellCount; i++)
                                        {
                                            
                                            cell = firstRow.GetCell(i);
                                            if (cell != null)
                                            {
                                                if (cell.StringCellValue != null)
                                                {
                                                    column = new DataColumn(cell.StringCellValue);
                                                    dataTable.Columns.Add(column);
                                                }
                                            }
                                        }
                                    }
                                    else
                                    {
                                        for (int i = firstRow.FirstCellNum; i < cellCount; i++)
                                        {
                                            column = new DataColumn("columns" + (i+1));
                                            dataTable.Columns.Add(column);
                                        }
                                    }

                                    //填充行  
                                    for (int i = startRow; i <= rowCount; i++)
                                    {
                                        row = sheet.GetRow(i);
                                        if (row == null) continue;
                                        dataRow = dataTable.NewRow();
                                        for (int j = row.FirstCellNum; j < cellCount; j++)
                                        {
                                            cell = row.GetCell(j);
                                            
                                            if (cell == null)
                                            {
                                                dataRow[j] = "";
                                            }
                                            else
                                            {
                                                switch (cell.CellType)
                                                {
                                                    case (NPOI.SS.UserModel.CellType)CellType.Blank:
                                                        dataRow[j] = "";
                                                        break;
                                                    case (NPOI.SS.UserModel.CellType)CellType.Numeric:
                                                        short format = cell.CellStyle.DataFormat;
                                                        //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                                        if (format == 14 || format == 31 || format == 57 || format == 58)
                                                            dataRow[j] = cell.DateCellValue;
                                                        else
                                                            dataRow[j] = (Double)cell.NumericCellValue;
                                                        break;
                                                    case (NPOI.SS.UserModel.CellType)CellType.String:
                                                        dataRow[j] = cell.StringCellValue;
                                                        break;
                                                }
                                            }
                                        }
                                      
                                        dataTable.Rows.Add(dataRow);
                                
                                    }
                                }
                            }
                        }
                    }
                    return dataTable;
                }
                catch (Exception ex)
                {
                    if (fs != null)
                    {
                        fs.Close();
                        Console.WriteLine(ex.Message);
                    }
                    return null;
                }
            }


            /// <summary>
            /// DataTable转List<T>
            /// </summary>
            /// <typeparam name="T">数据项类型</typeparam>
            /// <param name="dt">DataTable</param>
            /// <returns>List数据集</returns>
            public static List<T> DataTableToList<T>(System.Data.DataTable dt) where T : new()
            {
                List<T> list = new List<T>();
                if (dt != null && dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        T t = DataRowToModel<T>(dr);
                        list.Add(t);
                    }
                }
                return list;
            }

            /// <summary>
            /// DataRow转实体
            /// </summary>
            /// <typeparam name="IetmDmodule">数据型类</typeparam>
            /// <param name="dr">DataRow</param>
            /// <returns>模式</returns>
            public static T DataRowToModel<T>(DataRow dr) where T : new()
            {
                //T t = (T)Activator.CreateInstance(typeof(T));
                T t = new T();
                if (dr == null) return default(T);
                // 获得此模型的公共属性
                PropertyInfo[] propertys = t.GetType().GetProperties();
                DataColumnCollection Columns = dr.Table.Columns;
                foreach (PropertyInfo p in propertys)
                {
                    string columnName = p.Name;
                    if (Columns.Contains(columnName))
                    {
                        object value = dr[columnName];
                        if (value is DBNull || value == DBNull.Value)
                            continue;
                        try
                        {
                            switch (p.PropertyType.ToString())
                            {
                                case "System.String":
                                    p.SetValue(t, Convert.ToString(value), null);
                                    break;
                                case "System.Int32":
                                    p.SetValue(t, Convert.ToInt32(value), null);
                                    break;
                                case "System.Int64":
                                    p.SetValue(t, Convert.ToInt64(value), null);
                                    break;
                                case "System.DateTime":
                                    p.SetValue(t, Convert.ToDateTime(value), null);
                                    break;
                                case "System.Boolean":
                                    p.SetValue(t, Convert.ToBoolean(value), null);
                                    break;
                                case "System.Double":
                                    p.SetValue(t, Convert.ToDouble(value), null);
                                    break;
                                case "System.Decimal":
                                    p.SetValue(t, Convert.ToDecimal(value), null);
                                    break;
                                default:
                                    p.SetValue(t, value, null);
                                    break;
                            }
                        }
                        catch (Exception ex)
                        {
                            continue;
                            /*使用 default 关键字，此关键字对于引用类型会返回空，对于数值类型会返回零。对于结构，
                             * 此关键字将返回初始化为零或空的每个结构成员，具体取决于这些结构是值类型还是引用类型*/
                        }
                    }
                }
                return t;
            }

            public static bool DataTableToExcel(System.Data.DataTable dt)
            {
                bool result = false;
                IWorkbook workbook = null;
                FileStream fs = null;
                IRow row = null;
                ISheet sheet = null;
                ICell cell = null;
                try
                {
                    if (dt != null && dt.Rows.Count > 0)
                    {
                        string fileName = "C:/Users/cn-yangzheng/Desktop/1234.xls";
                        if (File.Exists(fileName))
                        {
                            fs = new FileStream(fileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                            POIFSFileSystem ps = new POIFSFileSystem(fs);
                            workbook = new HSSFWorkbook(ps);
                            sheet = workbook.GetSheetAt(0);
                        }
                        else
                        {
                            workbook = new HSSFWorkbook();
                            //sheet = workbook.CreateSheet("Sheet1");//创建一个名称为Sheet0的表  
                            sheet = workbook.CreateSheet("Sheet0");
                        }
                        row = sheet.GetRow(0);
                        int rowCount = dt.Rows.Count;//行数  
                        int columnCount = dt.Columns.Count;//列数  

                        //设置列头  
                        row = sheet.CreateRow(0);//excel第一行设为列头  
                        for (int c = 0; c < columnCount; c++)
                        {
                            cell = row.CreateCell(c);
                            cell.SetCellValue(""/*dt.Columns[c].ColumnName*/);
                        }

                        //设置每行每列的单元格,  
                        for (int i = 0; i < rowCount; i++)
                        {
                            row = sheet.CreateRow(i);
                            for (int j = 0; j < columnCount; j++)
                            {
                                cell = row.CreateCell(j);//excel第二行开始写入数据 
                                Console.WriteLine(dt.Columns[j].DataType.ToString()+ dt.Rows[i][j]);
                                if (dt.Columns[j].DataType == typeof(int))
                                {
                                    cell.SetCellValue((int)dt.Rows[i][j]);
                                }
                                else if (dt.Columns[j].DataType == typeof(float))
                                {
                                    cell.SetCellValue((float)dt.Rows[i][j]);
                                }
                                else if (dt.Columns[j].DataType == typeof(double))
                                {
                                    cell.SetCellValue((double)dt.Rows[i][j]);
                                }
                                else if (dt.Columns[j].DataType == typeof(string))
                                {
                                    cell.SetCellValue(dt.Rows[i][j].ToString());
                                }
                            }
                        }
                        using (fs = File.OpenWrite(@"C:/Users/cn-yangzheng/Desktop/1234.xls"))
                        {
                            workbook.Write(fs);//向打开的这个xls文件中写入数据  
                            result = true;
                        }
                    }
                    return result;
                }
                catch (Exception ex)
                {
                    if (fs != null)
                    {
                        
                        fs.Close();
                        Console.WriteLine(ex.Message);
                    }
                    return false;
                }

            }
        enum CellType
        {
            Unknown = -1, Numeric = 0, String = 1, Formula = 2, Blank = 3, Boolean = 4, Error = 5
        }
        }

        //rowCount 总共要写入的行数  columnCount 总共要写入的列数
        public static void ModelToExcel(string filepath,int rowCount,int columnCount,object[,] data)
        {
                  
            
            FileStream fs = null;
            fs = File.OpenRead(filepath);
            
            int rowHead = 3;//定义文档行头为3行
            int colHead = 1;//定义文档列头为1行
                
            
            
            ArrayList al = new ArrayList();
            
            IWorkbook workbook = new HSSFWorkbook(fs);
            fs.Close();
            ISheet sheet = workbook.GetSheetAt(0);
            
            /*
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(3, 6, 1, 1));//将单元格进行合并
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(3, 6, 2, 2));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(7, 11, 1, 1));
            sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(7, 11, 2, 2));
            */
            IRow row ;
            ICell cell ;
                //ICellStyle style = workbook.CreateCellStyle();
                //style.Alignment = HorizontalAlignment.Left;
            for(int i = rowHead; i < rowHead + rowCount; i++)//遍历表格的行
                {
                    row = sheet.GetRow(i);//获取第i行
                    for(int j =colHead; j <colHead + columnCount; j++)//遍历表格每一行的每一列
                    {
                    cell = row.GetCell(j);//获取第j个元素的数值
                    cell.SetCellValue(data[i-rowHead,j-colHead].ToString());//将data二维数组中数据存入相对应的单元格当中
                    }
                }
            int firstCount = 0;

            for (int k = 0; k < data.GetUpperBound(0) + 1; k++)
            {

                if (data[k, 0].ToString() != "")
                {
                    Console.WriteLine(data[k, 0]);
                    //Console.WriteLine(k);
                    firstCount = k + 1;
                    al.Add(firstCount);
                }
            }
            foreach(int i in al)
            {
                Console.WriteLine(i);
            }
            for (int i = 0; i < al.Count; i++)
            {
                int region1;
                int region2;
                try
                {
                    if (i < al.Count-1)//list集合的Count属性是多少个元素，由于i是从0开始的，所以i永远小于Count，必须先将Count减一操作之后再和i进行比较。
                    {
                        region1 = (int)al[i] + rowHead - 1; //代表数据中的第一个为Account Name
                        region2 = (int)al[i + 1] + rowHead - 2;//代表数据中的下一个Account Name
                                                               //所以要合并的区域为 region1+rowHead - 1 到 region2 -1 + rowHead -1
                    }
                    else
                    {
                        region1 = (int)al[i] + rowHead - 1; //代表数据中的第一个为Account Name
                        region2 = data.GetLength(0)+rowHead - 1;//数组的GetLength方法，当数值为0时，得到的是数组的行值，当数值为1时，得到的是数组的列值
                        //这个方法首先得到所有数据的行值，然后再加上行头，-1之后就得到表格的最后一行
                    }
                    Console.WriteLine(region1 + " " + region2);
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(region1, region2, 1, 1));//将单元格进行合并
                    sheet.AddMergedRegion(new NPOI.SS.Util.CellRangeAddress(region1, region2, 2, 2));
                }
                catch(Exception E)
                {
                    break;
                }
            }

            FileStream fileStream = File.OpenWrite(filepath);
            workbook.Write(fileStream);
            fileStream.Close();
           
            }

        public static void CreateNewWorkBook(string inputPath,string outputPath)
        {
            
            
        }
        }
    }

