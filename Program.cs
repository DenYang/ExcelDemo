using System;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
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

namespace ConsoleApp
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("代码开始执行...");
            //Excel.Application app = new Excel.Application();
            //app.Visible = true;//决定是隐藏打开，还是跳出Excel界面
            //Excel.Workbooks books = exl.Workbooks;
            //Excel.Workbook book = books.Open(@"C:\Users\cn-yangzheng\Desktop\template_Staff Exps in Key GnA Acct.xls");
            //Excel.Sheets sts = book.Worksheets;
            // int SheetCount = sts.Count;
            // Excel.Worksheet st = sts.Item[1];//sts.Item[1 - SheetCount]
            // Excel.Range rng = st.Application.Rows[5];//返回的是一个range类型，具体的方法可以用对象浏览器查看
            //dynamic value = rng.Text;
            //MessageBox.Show(value);//字符串 数字 日期 都可以正常使用
            //rng.Select();
            //rng.Resize[1].Insert();
            //book.Close();
            //insertData("C:/Users/cn-yangzheng/Desktop/多表数据 - 副本.xls",11);
            copyRange();
            //insertData(app);
            //System.Data.DataTable dt = ReadExcel();
            //WriteExcel(dt);
            

            Console.WriteLine("代码执行完毕！");

            //System.Data.DataTable dt= ReadExcel();
            //SaveCsv(dt, "C:/Users/cn-yangzheng/Desktop/");

        }
        /*
         x代表要插入的行
         */
        public static void insertData(string filePath,int x)
        {
            IWorkbook book;
            
            using(FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
               
                book = new HSSFWorkbook(fs);
                ISheet sheet = book.GetSheetAt(1);//获取sheet
                var row = sheet.GetRow(x - 1);//获取第x行
                sheet.ShiftRows(x-1, sheet.LastRowNum, 1);//从x行开始往下移动1行
                var newRow = sheet.CreateRow(x-1);//创建新的一行
                /*if (rowStyle != null)
                   newRow.RowStyle = rowStyle;
                   newRow.Height = rowSelf.Height;
                for (int col = 0; col < newRow.LastCellNum; col++)
                {
                    var cellsource = rowSelf.GetCell(col);
                    var cellInsert = newRow.CreateCell(col);
                    var cellStyle = cellsource.CellStyle;
                    //设置单元格样式　　　　
                    if (cellStyle != null)
                        cellInsert.CellStyle = cellsource.CellStyle;
                }*/
                FileStream out2 = new FileStream(filePath, FileMode.Create);
                book.Write(out2);
                out2.Close();
                fs.Close();
            }
            /*app.Visible = false;
            
            int selectNum,insertNum;
          
            Excel.Workbooks books = app.Workbooks;
            Excel.Workbook book = books.Open(@"C:\Users\cn-yangzheng\Desktop\多表数据 - 副本.xls");
            Excel.Sheets sts = book.Worksheets;
            //int SheetCount = sts.Count;
            Excel.Worksheet st = sts.Item[3];//sts.Item[1 - SheetCount]
            Console.WriteLine("请输入选定行数：...");
            selectNum = ReadInt();
            Excel.Range rng = st.Application.Rows[selectNum];//返回的是一个range类型，选定第5行数据
            //dynamic value = rng.Text;
            //MessageBox.Show(value);//字符串 数字 日期 都可以正常使用
            rng.Select();
            Console.WriteLine("请输入插入行数：...");
            insertNum = ReadInt();
            rng.Select();
            rng.Resize[insertNum].Insert();//进行1行插入
            app.DisplayAlerts = false;
            book.Save();
            book.Close();
            //app.Quit();*/
        }

        public static void copyRange()
        {
            int inputNum1, inputNum2, outputNum1, outputNum2;
          /*  Excel.Workbooks books = app.Workbooks;
            Excel.Workbook book = books.Open(@"C:\Users\cn-yangzheng\Desktop\多表数据 - 副本.xls");
            Excel.Sheets sts = book.Worksheets;
            //int SheetCount = sts.Count;
            Excel.Worksheet st = sts.Item[3]*/;//sts.Item[1 - SheetCount]
            //Excel.Range rng = st.Application.Cells[4];
           /* Console.WriteLine("请输入开始复制的行数：...");
            inputNum1 = ReadInt();
            Console.WriteLine("请输入结束复制的行数：...");
            inputNum2 = ReadInt();
            count = inputNum2 - inputNum1;
            string str1 = inputNum1.ToString();
            string str2 = inputNum2.ToString();
            Excel.Range rng = st.Range[str1+":"+str2];*///获取4到5行的表格数据
           // rng.Select();//选择表格数据
           /* Console.WriteLine("请输入要复制到的行数：...");
            outputNum1 = ReadInt();
            outputNum2 = outputNum1 + count;
            string str3 = outputNum1.ToString();
            string str4 = outputNum2.ToString();
            rng.Copy(st.Range[str3+":"+str4]);//复制到9到10行
            app.DisplayAlerts = false;
            book.Save();
            book.Close();*/
            //app.Quit();

            /**
             * 1、首先获取用户想要复制的行数
             * 2、使用遍历整行的方式获取区域内的表格数据
             * 3、将表格数据存入到DataTable中
             * 4、用户复制x行，整个文档从x-1处整体下移
             * 5、插入表格行数，并将数据重新写入到x-1处
             * 6。设置单元格格式
             */
            IWorkbook workbook = null;
            ISheet sheet = null;
            ICell cell = null;
            System.Data.DataTable dt = null;
            DataColumn column = null;
            DataRow datarow = null;

            string filepath = "C:/Users/cn-yangzheng/Desktop/多表数据 - 副本.xls";
            FileStream fs = new FileStream(filepath, FileMode.Open);
            workbook = new HSSFWorkbook(fs);
            sheet = workbook.GetSheetAt(2);
            Console.WriteLine("请输入开始复制的行数：...");
            inputNum1 = ReadInt();
            Console.WriteLine("请输入结束复制的行数：...");
            inputNum2 = ReadInt();
            int copyNum = inputNum2 - inputNum1;//用户复制的行数
            IRow row = sheet.GetRow(inputNum1-1);//获取区域行数
            int cellCount = row.LastCellNum;//获取区域列数
            dt = new System.Data.DataTable("t_copy");
            for (int i = row.FirstCellNum; i < cellCount; i++)
            {
                cell = row.GetCell(i);
                cell.SetCellType(CellType.String);
                if (cell != null)
                {
                    if (cell.StringCellValue != null)
                    {

                        column = new DataColumn(cell.StringCellValue);
                        dt.Columns.Add(column);
                    }
                }
            }

            for(int i = inputNum1 - 1; i < copyNum; i++)
            {
                IRow rowSelf = sheet.GetRow(i);
                if (rowSelf == null) continue;

                datarow = dt.NewRow();
                for(int j = rowSelf.FirstCellNum;j< cellCount; j++)
                {
                    cell = rowSelf.GetCell(j);
                    if (cell == null)
                    {
                        datarow[j] = "";
                    }
                    else
                    {
                        switch (cell.CellType)
                        {
                            case (NPOI.SS.UserModel.CellType)CellType.Blank:
                                datarow[j] = "";
                                break;
                            case (NPOI.SS.UserModel.CellType)CellType.Numeric:
                                short format = cell.CellStyle.DataFormat;
                                //对时间格式（2015.12.5、2015/12/5、2015-12-5等）的处理  
                                if (format == 14 || format == 31 || format == 57 || format == 58)
                                    datarow[j] = cell.DateCellValue;
                                else
                                    datarow[j] = cell.NumericCellValue;
                                break;
                            case (NPOI.SS.UserModel.CellType)CellType.String:
                                datarow[j] = cell.StringCellValue;
                                break;
                        }
                    }
                    dt.Rows.Add(datarow);
                }
            }
            Console.WriteLine("请输入要复制到的行数：...");
            outputNum1 = ReadInt();
            outputNum2 = outputNum1 + copyNum;
            sheet.ShiftRows(outputNum1 - 1, sheet.LastRowNum, copyNum);
            for (int i = row.FirstCellNum; i < cellCount; i++) {
                var newRow = sheet.CreateRow(outputNum1 - 1); }
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

        public static void modelToExcel()
        {
            FileStream file = new FileStream(@"templatebook1.xls",FileMode.Open,FileAccess.Read);
            HSSFWorkbook book = new HSSFWorkbook(file);
            HSSFSheet sheet = (HSSFSheet)book.GetSheet("Sheet1");
            HSSFCellStyle cellStyle = (HSSFCellStyle)book.CreateCellStyle();

            int rowCount = sheet.LastRowNum;//行数
            IRow firstRow = sheet.GetRow(0);
            int cellCount = firstRow.LastCellNum;//列数
            for (int i = 0;i<2;i++)
            {
                HSSFCell cell = (HSSFCell)sheet.GetRow(1).CreateCell(2);
            }
            
        }
    }
}

