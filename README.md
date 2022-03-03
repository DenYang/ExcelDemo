# ExcelDemo
这是一个使用VSTO和NPOI联合操作Excel的小程序  
实现了表格插入和复制功能，读取和录入数据功能目前还不完美，仅支持最基本的数据录入  
插入方法 InsertData(string path,int rowCount),path 表示文件路径，rowCount 表示传入的数据行数，该方法默认行头为3  
复制方法 CopyRange(string),path 表示文件路径，该方法会自动复制表格样式及数据  
数据导入方法 ModelToExcel(string filepath,int rowCount,int columnCount,object[,] data),filepath 表示文件路径,rowCount表示总共要写入数据的行数，columnCount表示总共要写入数据的列数，data是一个存放数据的二维数组，该方法默认行头为3，列头为1，会自动将数据写入Excel空模板中    
待续...  
