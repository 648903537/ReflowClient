using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace com.amtec.action
{
    /// <summary>
    /// NPOI帮助类
    /// </summary>
    public class ExcelHelper
    {
        #region 变量

        /// <summary>
        /// 工作薄
        /// </summary>
        private IWorkbook _IWorkbook;

        #endregion

        #region 输出信息
        public string Message { get; set; }
        #endregion

        #region 构造函数

        /// <summary>
        /// 构造函数
        /// </summary>
        public ExcelHelper() { }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="stream">文件流</param>
        public ExcelHelper(Stream stream)
        {
            _IWorkbook = CreateWorkbook(stream);
        }

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="fileName"></param>
        public ExcelHelper(string fileName)
        {
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                _IWorkbook = CreateWorkbook(fileStream);
            }
        }

        #endregion

        #region 方法

        /// <summary>
        /// 首个Sheet导入Table
        /// </summary>
        /// <returns></returns>
        public DataTable ExportToDataTable(bool needTrim = false)
        {
            return ExportToDataTable(_IWorkbook.GetSheetAt(0), needTrim);
        }

        /// <summary>
        /// 根据Sheet索引号导入Table
        /// </summary>
        /// <param name="sheetIndex">Sheet编号，从1开始</param>
        /// <returns></returns>
        public DataTable ExportExcelToDataTable(int sheetIndex)
        {
            return ExportToDataTable(_IWorkbook.GetSheetAt(sheetIndex - 1));
        }

        /// <summary>
        /// 首个Sheet导入集合
        /// </summary>
        /// <param name="fields">Excel各列，依次要转换成为的对象字段名称</param>
        /// <returns></returns>
        public IList<T> ExcelToList<T>(string[] fields) where T : class,new()
        {
            return ExportToList<T>(_IWorkbook.GetSheetAt(0), fields);
        }

        /// <summary>
        /// 指定Sheet导入集合
        /// </summary>
        /// <param name="sheetIndex">Sheet编号，从1开始</param>
        /// <param name="fields">Excel各列，依次要转换成为的对象字段名称</param>
        /// <returns></returns>
        public IList<T> ExcelToList<T>(int sheetIndex, string[] fields) where T : class,new()
        {
            return ExportToList<T>(_IWorkbook.GetSheetAt(sheetIndex - 1), fields);
        }

        /// <summary>
        /// Sheet中的数据转换为List集合
        /// </summary>
        /// <typeparam name="T">泛型</typeparam>
        /// <param name="sheetIndex">Sheet编号，从1开始</param>
        /// <param name="rowIndex">起始行索引</param>
        /// <param name="colIndex">起始列索引</param>
        /// <param name="fields">字段数组</param>
        /// <returns></returns>
        public IList<T> ExcelToList<T>(int sheetIndex, int rowIndex, int colIndex, string[] fields, int emptyIndex) where T : class,new()
        {
            return ExportToList<T>(_IWorkbook.GetSheetAt(sheetIndex - 1), rowIndex, colIndex, fields, emptyIndex);
        }

        #endregion

        #region 函数

        /// <summary>
        /// 创建工作簿对象
        /// </summary>
        /// <param name="stream"></param>
        /// <returns></returns>
        private IWorkbook CreateWorkbook(Stream stream)
        {
            try
            {
                return new XSSFWorkbook(stream);
            }
            catch(Exception ex)
            {
                return new HSSFWorkbook(stream);
            }
        }

        /// <summary>
        /// Sheet中的数据转换为DataTable
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="needTrim"></param>
        /// <returns></returns>
        private DataTable ExportToDataTable(ISheet sheet, bool needTrim = false)
        {
            // 
            DataTable dt = new DataTable();
            // 默认，第一行是字段
            IRow headRow = sheet.GetRow(0);
            // 设置datatable字段
            // 设置最大列数
            int minCellNum = headRow.FirstCellNum;
            int maxCellNum = headRow.LastCellNum;
            for (int i = headRow.FirstCellNum, len = headRow.LastCellNum; i < len; i++)
            {
                dt.Columns.Add(headRow.Cells[i].StringCellValue);
            }
            // 遍历数据行
            for (int i = (sheet.FirstRowNum + 1), len = sheet.LastRowNum + 1; i < len; i++)
            {
                IRow tempRow = sheet.GetRow(i);
                DataRow dataRow = dt.NewRow();
                // 遍历一行的每一个单元格
                bool isEmpty = true;
                for (int r = 0, j = minCellNum, len2 = maxCellNum; j < len2; j++, r++)
                {
                    if (tempRow == null)
                        continue;
                    ICell cell = tempRow.GetCell(j);
                    if (cell != null)
                    {
                        switch (cell.CellType)
                        {
                            case CellType.STRING:
                                if (needTrim)
                                {
                                    dataRow[r] = cell.StringCellValue.Trim();
                                }
                                else
                                {
                                    dataRow[r] = cell.StringCellValue;
                                }
                                break;
                            case CellType.NUMERIC:
                                if (HSSFDateUtil.IsCellDateFormatted(cell)) // 日期类型
                                {
                                    dataRow[r] = cell.DateCellValue.ToString("yyyy-MM-dd");
                                }
                                else//其他数字类型
                                {
                                    dataRow[r] = cell.NumericCellValue.ToString();
                                }
                                break;
                            case CellType.BOOLEAN:
                                dataRow[r] = cell.BooleanCellValue;
                                break;
                            default:
                                dataRow[r] = "";
                                break;
                        }
                        if (!string.IsNullOrEmpty(dataRow[r].ToString().Trim()))
                        {
                            isEmpty = false;
                        }
                    }
                    else
                    {
                        dataRow[r] = "";
                    }
                }
                if (!isEmpty)
                {
                    dt.Rows.Add(dataRow);
                }
            }
            // 返回
            return dt;
        }

        /// <summary>
        /// Sheet中的数据转换为List集合
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fields"></param>
        /// <returns></returns>
        private IList<T> ExportToList<T>(ISheet sheet, string[] fields)
            where T : class, new()
        {
            IList<T> list = new List<T>();
            try
            {
                // 遍历每一行数据
                for (int i = sheet.FirstRowNum + 1, len = sheet.LastRowNum + 1; i < len; i++)
                {
                    T t = new T();
                    IRow row = sheet.GetRow(i);
                    for (int j = 0, len2 = fields.Length; j < len2; j++)
                    {
                        if (string.IsNullOrEmpty(fields[j])) continue;
                        ICell cell = row.GetCell(j);
                        object cellValue = null;
                        if (cell != null)
                        {
                            switch (cell.CellType)
                            {
                                case CellType.STRING:    // 文本
                                    cellValue = cell.StringCellValue;
                                    break;
                                case CellType.NUMERIC:   // 数值
                                    cellValue = Convert.ToInt32(cell.NumericCellValue);//Double转换为int
                                    break;
                                case CellType.BOOLEAN:   // bool
                                    cellValue = cell.BooleanCellValue;
                                    break;
                                case CellType.BLANK:     // 空白
                                    cellValue = "";
                                    break;
                                case CellType.FORMULA:
                                    switch (cell.CachedFormulaResultType)
                                    {
                                        case CellType.NUMERIC:
                                            cellValue = Convert.ToInt32(cell.NumericCellValue);
                                            break;
                                        case CellType.STRING:
                                            cellValue = cell.StringCellValue;
                                            break;
                                        default:
                                            cellValue = "";
                                            break;
                                    }
                                    break;
                                default:
                                    cellValue = "";
                                    break;
                            }
                        }
                        else
                        {
                            cellValue = string.Empty;
                        }
                        try
                        {
                            typeof(T).GetProperty(fields[j]).SetValue(t, cellValue, null);
                        }
                        catch
                        {
                            typeof(T).GetProperty(fields[j]).SetValue(t, cellValue.ToString(), null);
                        }

                    }
                    list.Add(t);
                }
            }
            catch (Exception ex)
            {
                list = new List<T>();
            }
            return list;
        }

        /// <summary>
        /// Sheet中的数据转换为List集合
        /// </summary>
        /// <typeparam name="T">泛型</typeparam>
        /// <param name="sheet">Sheet对象</param>
        /// <param name="rowIndex">起始行索引</param>
        /// <param name="colIndex">起始列索引</param>
        /// <param name="fields">字段数组</param>
        /// <param name="emptyIndex">为空判断</param>
        /// <returns></returns>
        private IList<T> ExportToList<T>(ISheet sheet, int rowIndex, int colIndex, string[] fields, int emptyIndex)
            where T : class, new()
        {
            // 实例化泛型列表
            IList<T> list = new List<T>();
            try
            {
                // 从索引行开始进行遍历
                for (int i = rowIndex, len = sheet.LastRowNum + 1; i < len; i++)
                {
                    T t = new T();
                    IRow row = sheet.GetRow(i);
                    ICell keyCell = row.GetCell(emptyIndex);
                    if (keyCell == null)
                    {
                        break;
                    }
                    if(keyCell != null && string.IsNullOrEmpty(keyCell.StringCellValue.Trim()))
                    {
                        break;
                    }
                    // 遍历列
                    for (int j = 0, len2 = fields.Length; j < len2; j++)
                    {
                        if (string.IsNullOrEmpty(fields[j])) continue;
                        ICell cell = row.GetCell(j + colIndex);
                        object cellValue = null;
                        if (cell != null)
                        {
                            switch (cell.CellType)
                            {
                                case CellType.STRING:    // 文本
                                    cellValue = cell.StringCellValue;
                                    break;
                                case CellType.NUMERIC:   // 数值
                                    if (HSSFDateUtil.IsCellDateFormatted(cell))
                                    {
                                        cellValue = cell.DateCellValue.ToString("yyyy/MM/dd");
                                    }
                                    else
                                    {
                                        cellValue = Convert.ToInt32(cell.NumericCellValue);//Double转换为int
                                    }
                                    break;
                                case CellType.BOOLEAN:   // bool
                                    cellValue = cell.BooleanCellValue;
                                    break;
                                case CellType.BLANK:     // 空白
                                    cellValue = "";
                                    break;
                                case CellType.FORMULA:
                                    switch(cell.CachedFormulaResultType)
                                    {
                                        case CellType.NUMERIC:
                                            cellValue = Convert.ToInt32(cell.NumericCellValue);
                                            break;
                                        case CellType.STRING:
                                            cellValue = cell.StringCellValue;
                                            break;
                                        default:
                                            cellValue = "";
                                            break;
                                    }
                                    break;
                                default:
                                    cellValue = "";
                                    break;
                            }
                        }
                        else
                        {
                            cellValue = string.Empty;
                        }
                        try
                        {
                            typeof(T).GetProperty(fields[j]).SetValue(t, cellValue, null);
                        }
                        catch
                        {
                            typeof(T).GetProperty(fields[j]).SetValue(t, cellValue.ToString(), null);
                        }

                    }
                    list.Add(t);
                }
            }
            catch
            {
                list = new List<T>();
            }
            return list;
        }

        #endregion
    }
}