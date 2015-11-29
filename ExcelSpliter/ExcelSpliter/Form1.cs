using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelSpliter
{
    public partial class Form1 : Form
    {
        private readonly OpenFileDialog dialog;

        public delegate void SetProcess(string msg);
        public Form1()
        {
            InitializeComponent();
            dialog = new OpenFileDialog
            {
                Filter = "excel 2007|*.xlsx|excel 2003|*.xls"
            };
            dgvExcelList.ReadOnly = true;
        }

        //异步访问控件
        private void SetLableText(string msg)
        {
            if (lblProgress.InvokeRequired)
            {
                var set = new SetProcess(SetLableText);
                //委托的方法参数应和SetCalResult一致  
                //此方法第二参数用于传入方法,代替形参result  
                lblProgress.Invoke(set, new object[] { msg });
            }
            else
            {
                lblProgress.Text = msg;
            }
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            dialog.ShowDialog();
            var fileName = dialog.FileName;
            dgvExcelList.Rows.Add(fileName);
        }

        private void btnDone_Click(object sender, EventArgs e)
        {
            var files = GetFiles();
            if (files.Count == 0) return;
            int fileCount;
            int.TryParse(txtFileCount.Text.Trim(), out fileCount);
            SetLableText("开始处理");
            //单文件分拆
            if (rdSplitFile.Checked || rdSplitSheet.Checked)
            {
                var dic = new Dictionary<string, object>
                    {
                        {"isToFile", rdSplitFile.Checked},
                        {"file", files[0]},
                        {"fileCount", fileCount == 0 ? 2 : fileCount}
                    };
                ThreadPool.QueueUserWorkItem(FileSplit, dic);
            }
            //多文件合并
            if (rdMergeToManySheet.Checked || rdMergeToSingleSheet.Checked)
            {
                var dic = new Dictionary<string, object>
                    {
                        {"isToFile", rdMergeToSingleSheet.Checked},
                        {"file", files}
                    };
                ThreadPool.QueueUserWorkItem(FileMerge, dic); 
            }
        }

        //文件拆分 
        //isToFile 是否拆分成文件
        private void FileSplit(object obj)
        {
            var dic = (Dictionary<string, object>)obj;
            var isToFile = Convert.ToBoolean(dic["isToFile"]);
            var file = dic["file"].ToString();
            var count = Convert.ToInt32(dic["fileCount"]);
            var helper = new ExcelHelper(file);
            var table = helper.ExcelToDataTable("Sheet1", true);
            if (table == null || table.Rows.Count == 0) return;
            var i = isToFile ? SplitToMultFile(table, count)
                : SplitMultSheet(table, count);
            SetLableText("文件已经保存到D盘");
        }

        //文件合并
        private void FileMerge(object obj)
        {
            try
            {
                var dic = (Dictionary<string, object>)obj;
                var isToFile = Convert.ToBoolean(dic["isToFile"]);
                var files = dic["file"] as List<string>;
                if (files == null) return;
                var tables = new List<DataTable>();
                for (int i = 0, len = files.Count; i < len; i++)
                {
                    var helper = new ExcelHelper(files[i]);
                    var table = helper.ExcelToDataTable("Sheet1", true);
                    if (table == null || table.Rows.Count == 0) continue;
                    tables.Add(table);
                }
                if (tables.Count == 0) return;
                var k = isToFile ? MergeToSingleSheet(tables) : MergeToMultSheet(tables);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            SetLableText("文件已经保存到D盘");
        }

        private List<string> GetFiles()
        {
            var rows = dgvExcelList.Rows;
            var list = new List<string>();
            foreach (DataGridViewRow row in rows)
            {
                list.Add(row.Cells[0].Value.ToString());
            }
            return list;
        }

        //合并成单工作区间
        private int MergeToSingleSheet(List<DataTable> tables)
        {
            try
            {
                var maxCol = 0;
                var tableIndex = -1;
                foreach (DataTable table in tables)
                {
                    var len = table.Rows[0].ItemArray.Length;
                    if (len > maxCol)
                    {
                        maxCol = len;
                        tableIndex++;
                    }
                }
                var newTable = InitTable(tables[tableIndex]);
                const string path = "D:\\{0}.xlsx";
                foreach (var table in tables)
                {
                    var rows = table.Rows;
                    foreach (DataRow row in rows)
                    {
                        var newRow = newTable.NewRow();
                        newTable.Rows.Add(GetRow(row, newRow));
                    }
                }
                if (newTable.Rows.Count == 0) return -1;
                new ExcelHelper(string.Format(path, DateTime.Now.ToString("yyyyMMddHHmmss")))
                           .DataTableToExcel(newTable, "Sheet1", true);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           
            return 0;
        }

        //合并成多个工作区间
        private int MergeToMultSheet(List<DataTable> tables)
        {
            try
            {
                const string path = "D:\\{0}.xlsx";
                new ExcelHelper(string.Format(path, DateTime.Now.ToString("yyyyMMddHHmmss")))
                           .DataTableToExcelWithMultSheet(tables, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
           
            return 0;
        }

        //拆分文件
        private int SplitToMultFile(DataTable table, int fileCount)
        {
            try
            {
                var rowsCount = table.Rows.Count;
                var sheetRowCount = rowsCount / fileCount;
                var newTable = InitTable(table);
                var index = 1;
                const string path = "D:\\{0}_{1}.xlsx";
                for (int i = 0; i < rowsCount; i++)
                {
                    var newRow = newTable.NewRow();
                    newTable.Rows.Add(GetRow(table.Rows[i], newRow));
                    if (i != 0 && i % sheetRowCount == 0)
                    {
                        new ExcelHelper(string.Format(path, DateTime.Now.ToString("yyyyMMddHHmmss"), index))
                            .DataTableToExcel(newTable, "Sheet1", true);
                        newTable.Rows.Clear();
                        index++;
                    }
                }
                if (newTable.Rows.Count > 0)
                {
                    new ExcelHelper(string.Format(path, DateTime.Now.ToString("yyyyMMddHHmmss"), index))
                            .DataTableToExcel(newTable, "Sheet1", true);
                }
                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return 0;
        }

        //拆分成多个工作区间
        private int SplitMultSheet(DataTable table, int sheetCount)
        {
            try
            {
                var rowsCount = table.Rows.Count;
                var sheetRowCount = rowsCount / sheetCount;
                var tableList = new List<DataTable>();
                var newTable = InitTable(table);
                const string path = "D:\\{0}.xlsx";
                for (int i = 0; i < rowsCount; i++)
                {
                    var newRow = newTable.NewRow();
                    newTable.Rows.Add(GetRow(table.Rows[i], newRow));
                    if (i != 0 && i % sheetRowCount == 0)
                    {
                        tableList.Add(newTable);
                        newTable = InitTable(table);
                    }
                }
                if (newTable.Rows.Count > 0)
                {
                    tableList.Add(newTable);
                }
                new ExcelHelper(string.Format(path, DateTime.Now.ToString("yyyyMMddHHmmss")))
                           .DataTableToExcelWithMultSheet(tableList, true);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            
            return 0;
        }

        //初始化表
        private DataTable InitTable(DataTable oldTable)
        {
            var table = new DataTable();
            var cols = oldTable.Columns;
            foreach (DataColumn col in cols)
            {
                table.Columns.Add(new DataColumn
                {
                    ColumnName = col.ColumnName,
                    Caption = col.Caption,
                    DataType = col.DataType
                });
            }
            return table;
        }

        //初始化row
        private DataRow GetRow(DataRow oldRow, DataRow newRow)
        {
            var cells = oldRow.ItemArray.Length;
            for (int i = 0; i < cells; i++)
            {
                newRow[i] = oldRow[i];
            }
            return newRow;
        }

        private void btnRemove_Click(object sender, EventArgs e)
        {
            var current = dgvExcelList.CurrentRow;
            if (current != null)
            {
                dgvExcelList.Rows.Remove(current);
            }
        }
    }
}
