using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Excel2DualSql
{
    public partial class frmExcel2DualSql : Form
    {
        public frmExcel2DualSql()
        {
            InitializeComponent();
        }

        private void 打开OToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog())
            {
                ofd.Filter = "Excel Files|*.xls|Excel Files|*.xlsx";
                ofd.FilterIndex = 0;
                ofd.RestoreDirectory = true;
                ofd.Multiselect = true;
                ofd.Title = "读取Excel文件";
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    if (string.IsNullOrEmpty(ofd.FileName) || ofd.FileNames.Length == 0)
                    {
                        MessageBox.Show("请选择Excel文件！");
                    }
                    else
                    {
                        try
                        {
                            string[] filenames = ofd.FileNames;
                            foreach (var item in filenames)
                            {
                                string sql = Excel2DualSql(item, "SQL Results", true);
                                if (!string.IsNullOrEmpty(sql))
                                {
                                    UpdateText(sql);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message);
                        }
                    }
                }
            }
        }

        private string Excel2DualSql(string fileName, string sheetName, bool isFirstRowColumn)
        {
            ISheet sheet = null;
            IWorkbook workbook = null;
            int startRow = 0;
            StringBuilder sbSql = new StringBuilder();
            sbSql.AppendLine(fileName);

            FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read);
            if (fileName.IndexOf(".xlsx") > 0) // 2007版本  
                workbook = new XSSFWorkbook(fs);
            else if (fileName.IndexOf(".xls") > 0) // 2003版本  
                workbook = new HSSFWorkbook(fs);
            if (sheetName != null)
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null) //如果没有找到指定的sheetName对应的sheet，则尝试获取第一个sheet  
                    sheet = workbook.GetSheetAt(0);
            }
            else
            {
                sheet = workbook.GetSheetAt(0);
            }

            if (sheet != null)
            {

                IRow firstRow = sheet.GetRow(0);
                int cellCount = firstRow.LastCellNum; //一行最后一个cell的编号 即总的列数  

                List<string> colNameList = new List<string>();
                if (isFirstRowColumn)
                {
                    for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                    {
                        ICell cell = firstRow.GetCell(i);
                        if (cell != null)
                        {
                            string cellValue = cell.StringCellValue;
                            colNameList.Add(cellValue);
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
                    sbSql.Append("select ");

                    IRow row = sheet.GetRow(i);
                    if (row == null) continue; //没有数据的行默认是null　　　　　　　  
                    for (int j = row.FirstCellNum + 1; j < cellCount; j++)
                    {
                        ICell cell = row.GetCell(j);
                        if (cell != null) //同理，没有数据的单元格都默认是null  
                        {
                            switch (cell.CellType)
                            {
                                case CellType.Numeric:
                                    sbSql.AppendFormat("{0} {1}", cell.ToString(), colNameList[j]);
                                    break;
                                default:
                                    sbSql.AppendFormat("'{0}' {1}", cell.ToString(), colNameList[j]);
                                    break;
                            }
                            sbSql.Append(j != cellCount - 1 ? ", " : "");
                        }
                        else
                        {
                            sbSql.AppendFormat("NULL {0}", colNameList[j]);
                            sbSql.Append(j != cellCount - 1 ? ", " : "");
                        }
                    }
                    sbSql.Append(i != rowCount ? " from dual union all " : " from dual");
                    sbSql.AppendLine("");
                }
            }
            return sbSql.ToString();
        }

        private void UpdateText(string text)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(o =>
                {
                    this.textBox1.AppendText(o);
                }));
            }
            else
            {
                this.textBox1.AppendText(text);
            }
        }
    }
}
