# ExcelStream

像操作二维数组一样操作EXCEL表格。

三个要点：
1. 命名空间:`using ExcelHelper;`
2. 类的构造 :  `public ExcelStream(string location, int sheet = 1)`
3. ExcelStream的Data数组，记录着表格数据
4. ExcelStream的方法 : void Read()  , void Write()

   注意：如果需要创建文件请Write一次，再Read一次
   注意：如果需要读取文件，请调用Read
   注意：Read作用是把Excel读到Data，Write作用是把Data写到Excel
   注意：Data数组的大小是有效数据行列数，需要需要扩展新的行列，请使用Data.SetSizeAndCopy(...);
