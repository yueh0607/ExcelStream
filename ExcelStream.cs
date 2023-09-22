
using OfficeOpenXml;
using System;
using System.IO;

namespace ExcelHelper
{
    public interface ITable
    {
        public int RowCount { get; }
        public int ColumnCount { get; }
    }
    public interface ITable<T> : ITable
    {
        T this[int row, int column] { get; set; }
    }
    public class ExcelStream : IDisposable, ITable<string>
    {

        public string SourcePath { get; set; } = string.Empty;

        public int Sheet { get; set; } = 1;
        private ConfigTable<string> data = null;

        public ConfigTable<string> Data
        {
            get
            {
                if (data == null) throw new NullReferenceException("Please read the data before accessing it");
                return data;
            }
        }


        public int RowCount => data.RowCount;


        public int ColumnCount => data.ColumnCount;

        public string this[int row, int column]
        {
            get => Data[row, column];
            set => Data[row, column] = value;

        }
//#if NETCOREAPP
//        public ExcelStream(string location, int sheet = 0, bool isCommerical = false)
//        {

//#else

 public ExcelStream(string location, int sheet = 1,bool isCommerical= false)
    {
//#endif

           


            //OfficeOpenXml.ExcelPackage.LicenseContext = isCommerical ? OfficeOpenXml.LicenseContext.Commercial : OfficeOpenXml.LicenseContext.NonCommercial;

            SourcePath = location;
            Sheet = sheet;
        }

        private void Create()
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add("Sheet1");
                package.SaveAs(new FileInfo(SourcePath));
            }
        }

        /// <summary>
        /// Read执行的操作是与Excel表进行IO交换，读取到Data
        /// </summary>
        public void Read()
        {
            if (!File.Exists(SourcePath)) throw new Exception("Please call write to create an empty file before reading a non-existent file");
            using (var package = new ExcelPackage(new FileInfo(SourcePath)))
            {

                // 获取工作表
                ExcelWorksheet worksheet = package.Workbook.Worksheets[Sheet]; // 默认读取第一个工作表.

                int rowCount;
                int columnCount;
                //初始化数据表
                if (worksheet.Dimension == null)
                {
                    rowCount = 0;
                    columnCount = 0;
                }
                else
                {
                    rowCount = worksheet.Dimension.Rows;
                    columnCount = worksheet.Dimension.Columns;
                }


                data = new ConfigTable<string>(rowCount, columnCount);

                //初始化单元格值
                for (int i = 0; i < rowCount; i++)
                {
                    for (int j = 0; j < columnCount; j++)
                    {
                        data[i, j] = worksheet.Cells[i + 1, j + 1].Text;
                    }
                }
            }
        }

        /// <summary>
        /// Write执行的操作是不存在则创建，存在则覆盖写入
        /// </summary>
        public void Write()
        {
            if (!File.Exists(SourcePath))
                Create();

            using (var package = new ExcelPackage(new FileInfo(SourcePath)))
            {

                ExcelWorksheet worksheet = package.Workbook.Worksheets[Sheet];

                for (int i = 0; i < data.RowCount; i++)
                {
                    for (int j = 0; j < data.ColumnCount; j++)
                    {
                        worksheet.Cells[i + 1, j + 1].Value = data[i, j];
                    }
                }

                package.Save();
            }
        }

        /// <summary>
        /// 添加一个页
        /// </summary>
        /// <param name="name"></param>
        public void AddSheet(string name)
        {
            using (var package = new ExcelPackage())
            {
                package.Workbook.Worksheets.Add(name);
                package.Save(SourcePath);
            }
        }

        public void Dispose()
        {
            data = null;
        }

    }
    public class ConfigTable<T> : ITable<T>
    {
        /// <summary>
        /// 行数
        /// </summary>
        public int RowCount { get; private set; } = 0;

        /// <summary>
        /// 列数
        /// </summary>
        public int ColumnCount { get; private set; } = 0;

#pragma warning disable CS8625 // 无法将 null 字面量转换为非 null 的引用类型。
        private T[,] data = null;
#pragma warning restore CS8625 // 无法将 null 字面量转换为非 null 的引用类型。

        public ConfigTable(int rowCount, int columnCount)
        {
            SetSizeAndCopy(rowCount, columnCount);
        }

        public ConfigTable(ConfigTable<T> table)
        {
            SetSizeAndCopy(table.RowCount, table.ColumnCount, false);
#pragma warning disable CS8604 // 引用类型参数可能为 null。
            table.data.CopyTo(data, 0);
#pragma warning restore CS8604 // 引用类型参数可能为 null。
        }


        /// <summary>
        /// 设置大小并拷贝
        /// </summary>
        /// <param name="newRowCount"></param>
        /// <param name="newColumnCount"></param>
        /// <param name="copy"></param>
        /// <exception cref="InvalidOperationException"></exception>
        public void SetSizeAndCopy(int newRowCount, int newColumnCount, bool copy = true)
        {
            if (newRowCount < 0 || newColumnCount < 0) throw new InvalidOperationException($"Invalid New Size:[{newRowCount},{newColumnCount}]");
            int minRowCount = Math.Min(RowCount, newRowCount);
            int minColumnCount = Math.Min(ColumnCount, newColumnCount);

            RowCount = newRowCount;
            ColumnCount = newColumnCount;

            var temp = data;
            data = new T[RowCount, ColumnCount];

            if (temp != null && copy)
            {
                for (int i = 0; i < minRowCount; i++)
                {
                    for (int j = 0; j < minColumnCount; j++)
                    {
                        data[i, j] = temp[i, j];
                    }
                }
            }
            temp = null;
        }


        private void IndexOutOfRangeCheck(int row, int column)
        {
            if (row >= RowCount || row < 0 || column >= ColumnCount || column < 0) throw new IndexOutOfRangeException($"Row:{row}({RowCount - 1})  Column:{column}({ColumnCount - 1})");
        }
        public T this[int row, int column]
        {

            get
            {
                IndexOutOfRangeCheck(row, column);
                return data[row, column];
            }
            set
            {
                IndexOutOfRangeCheck(row, column);
                data[row, column] = value;
            }
        }
    }
}