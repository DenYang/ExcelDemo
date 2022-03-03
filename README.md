# ExcelDemo
这是一个使用VSTO和NPOI联合操作Excel的小程序
实现了表格插入和复制功能，读取和录入数据功能目前还不完美，仅支持最基本的数据录入
插入方法 InsertData(string path,int rowCount),path 表示文件路径，rowCount 表示传入的数据行数，该方法默认行头为3
复制方法 CopyRange(string),path 表示文件路径，该方法会自动复制表格样式及数据
数据导入方法 ModelToExcel() 待续.....
