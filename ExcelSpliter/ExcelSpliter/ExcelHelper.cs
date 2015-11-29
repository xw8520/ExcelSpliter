using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;
using NPOI.OpenXmlFormats.Dml;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelSpliter
{
    public class ExcelHelper
    {
        private readonly string _fileName; //文件名
        private IWorkbook _workbook;
        private FileStream _fs;
        private bool _disposed;

        public ExcelHelper(string fileName)
        {
            _fileName = fileName;
            _disposed = false;
        }

        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="firstIsTitle">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcel(DataTable data, string sheetName, bool firstIsTitle)
        {
            if (data.Rows.Count == 0) return 0;
            _fs = new FileStream(_fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (_fileName.IndexOf(".xlsx") > 0) // 2007版本
                _workbook = new XSSFWorkbook();
            else if (_fileName.IndexOf(".xls") > 0) // 2003版本
                _workbook = new HSSFWorkbook();

            try
            {
                ISheet sheet = null;
                if (_workbook != null)
                {
                    sheet = _workbook.CreateSheet(sheetName);
                }
                else
                {
                    return -1;
                }
                InitSheetWidth(200 * 20, sheet, data.Rows[0].ItemArray.Length);
                var headStyle = InitTitleStyle();
                var bodyStyle = InitBodyStyle();
                int count = 0;
                int j = 0;
                if (firstIsTitle) //写入DataTable的列名
                {
                    IRow row = sheet.CreateRow(0);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        var cell = row.CreateCell(j);
                        cell.CellStyle = headStyle;
                        cell.SetCellValue(data.Columns[j].ColumnName);
                    }
                    count = 1;
                }
                else
                {
                    count = 0;
                }

                int i = 0;
                for (i = 0; i < data.Rows.Count; ++i)
                {
                    IRow row = sheet.CreateRow(count);
                    for (j = 0; j < data.Columns.Count; ++j)
                    {
                        var cell = row.CreateCell(j);
                        cell.CellStyle = bodyStyle;
                        cell.SetCellValue(data.Rows[i][j].ToString());
                    }
                    ++count;
                }
                _workbook.Write(_fs); //写入到excel
                return count;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
            finally
            {
                data.Dispose();
            }
        }


        /// <summary>
        /// 将DataTable数据导入到excel中
        /// </summary>
        /// <param name="tables">table列表</param>
        /// <param name="firstIsTitle">DataTable的列名是否要导入</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public int DataTableToExcelWithMultSheet(List<DataTable> tables, bool firstIsTitle)
        {
            if (tables.Count == 0) return 0;
            _fs = new FileStream(_fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            if (_fileName.IndexOf(".xlsx") > 0) // 2007版本
                _workbook = new XSSFWorkbook();
            else if (_fileName.IndexOf(".xls") > 0) // 2003版本
                _workbook = new HSSFWorkbook();
            try
            {
                for (int k = 0, len = tables.Count; k < len; k++)
                {
                    ISheet sheet = null;
                    if (_workbook != null)
                    {
                        sheet = _workbook.CreateSheet("Sheet" + (k + 1));
                    }
                    else
                    {
                        return -1;
                    }
                    InitSheetWidth(240 * 20, sheet, tables[k].Rows[0].ItemArray.Length);
                    var headStyle = InitTitleStyle();
                    var bodyStyle = InitBodyStyle();
                    int count = 0;
                    int j = 0;
                    if (firstIsTitle) //写入DataTable的列名
                    {
                        IRow row = sheet.CreateRow(0);
                        for (j = 0; j < tables[k].Columns.Count; ++j)
                        {
                            var cell = row.CreateCell(j);
                            cell.CellStyle = headStyle;
                            cell.SetCellValue(tables[k].Columns[j].ColumnName);
                        }
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }

                    int i = 0;
                    for (i = 0; i < tables[k].Rows.Count; ++i)
                    {
                        IRow row = sheet.CreateRow(count);
                        for (j = 0; j < tables[k].Columns.Count; ++j)
                        {
                            var cell = row.CreateCell(j);
                            cell.CellStyle = bodyStyle;
                            cell.SetCellValue(tables[k].Rows[i][j].ToString());
                        }
                        ++count;
                    }
                }

                _workbook.Write(_fs); //写入到excel
                return 1;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return -1;
            }
            finally
            {
                foreach (var table in tables)
                {
                    table.Dispose();
                }
            }
        }

        /// <summary>
        /// 初始化样式
        /// </summary>
        private void InitSheetWidth(int width, ISheet sheet, int colCount)
        {
            for (int i = 0; i < colCount; i++)
            {
                sheet.SetColumnWidth(i, width);
            }
        }

        /// <summary>
        /// 初始化标题样式
        /// </summary>
        /// <returns></returns>
        private ICellStyle InitTitleStyle()
        {
            var fHead = _workbook.CreateFont();
            fHead.Color = HSSFColor.Black.Index;
            fHead.Boldweight = (short)FontBoldWeight.Bold; //设置粗体
            fHead.FontHeightInPoints = 12;
            fHead.FontName = "宋体";
            fHead.IsStrikeout = false;
            ICellStyle styleHead = _workbook.CreateCellStyle();
            styleHead.SetFont(fHead);
            //边框
            styleHead.BorderBottom = BorderStyle.Thin;
            styleHead.BorderLeft = BorderStyle.Thin;
            styleHead.BorderRight = BorderStyle.Thin;
            styleHead.BorderTop = BorderStyle.Thin;
            //居中
            styleHead.Alignment = HorizontalAlignment.Center;
            return styleHead;
        }

        /// <summary>
        /// 初始化内容样式
        /// </summary>
        /// <returns></returns>
        private ICellStyle InitBodyStyle()
        {
            IFont fBody = _workbook.CreateFont();
            fBody.Color = HSSFColor.Black.Index;
            fBody.Boldweight = (short)FontBoldWeight.Normal; //设置粗体
            fBody.FontHeightInPoints = 12;
            fBody.FontName = "宋体";
            fBody.IsStrikeout = false;
            ICellStyle bodyStyle = _workbook.CreateCellStyle();
            bodyStyle.SetFont(fBody);
            //边框
            bodyStyle.BorderBottom = BorderStyle.Thin;
            bodyStyle.BorderLeft = BorderStyle.Thin;
            bodyStyle.BorderRight = BorderStyle.Thin;
            bodyStyle.BorderTop = BorderStyle.Thin;
            //居中
            bodyStyle.Alignment = HorizontalAlignment.Left;
            return bodyStyle;
        }

        /// <summary>
        /// 将excel中的数据导入到DataTable中
        /// </summary>
        /// <param name="sheetName">excel工作薄sheet的名称</param>
        /// <param name="firstIsTitle">第一行是否是DataTable的列名</param>
        /// <returns>返回的DataTable</returns>
        public DataTable ExcelToDataTable(string sheetName, bool firstIsTitle)
        {
            ISheet sheet = null;
            DataTable data = new DataTable();
            int startRow = 0;
            try
            {
                _fs = new FileStream(_fileName, FileMode.Open, FileAccess.Read);
                using (_fs)
                {
                    if (_fileName.IndexOf(".xlsx") > 0) // 2007版本
                        _workbook = new XSSFWorkbook(_fs);
                    else if (_fileName.IndexOf(".xls") > 0) // 2003版本
                        _workbook = new HSSFWorkbook(_fs);

                    if (sheetName != null)
                    {
                        sheet = _workbook.GetSheet(sheetName);
                        if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet
                        {
                            sheet = _workbook.GetSheetAt(0);
                        }
                    }
                    else
                    {
                        sheet = _workbook.GetSheetAt(0);
                    }
                    if (sheet != null)
                    {
                        IRow firstRow = sheet.GetRow(0);
                        int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数

                        if (firstIsTitle)
                        {
                            for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                            {
                                ICell cell = firstRow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellValue = cell.StringCellValue;
                                    if (cellValue != null)
                                    {
                                        var column = new DataColumn(cellValue);
                                        if (data.Columns.Contains(cellValue))
                                        {
                                            column = new DataColumn(cellValue + i);
                                        }
                                        data.Columns.Add(column);
                                    }
                                }
                            }
                            startRow = sheet.FirstRowNum + 1;
                        }
                        else
                        {
                            startRow = sheet.FirstRowNum;
                        }

                        //最后一列的标号
                        int rowCount = sheet.LastRowNum;
                        for (int i = startRow; i <= rowCount; ++i)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null) continue; //没有数据的行默认是null　　　　　　　

                            DataRow dataRow = data.NewRow();
                            for (int j = row.FirstCellNum; j < cellCount; ++j)
                            {
                                if (row.GetCell(j) != null) //同理，没有数据的单元格都默认是null
                                    dataRow[j] = row.GetCell(j).ToString();
                            }
                            data.Rows.Add(dataRow);
                        }
                    }
                }
                return data;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception: " + ex.Message);
                return null;
            }
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this._disposed)
            {
                if (disposing)
                {
                    if (_fs != null)
                        _fs.Close();
                }

                _fs = null;
                _disposed = true;
            }
        }
    }
}
