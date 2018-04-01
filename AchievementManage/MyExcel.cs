using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;//DataColumn和DataRow需要
using System.Windows.Forms;//DataGridView需要
using Excel;//Excel相关操作需要

namespace AchievementManage
{
    class MyExcel
    {
        public static System.Data.DataTable GetDgvToTable(DataGridView dgv)//将DataGridView控件中的内容转化成System.Data.DataTable
        {
            System.Data.DataTable dt = new System.Data.DataTable();
            for (int count = 0; count < dgv.Columns.Count; count++)//列强制转换
            {
                DataColumn dc = new DataColumn(dgv.Columns[count].Name.ToString());
                dt.Columns.Add(dc);
            }
            for (int count = 0; count < dgv.Rows.Count; count++)//循环行
            {
                DataRow dr = dt.NewRow();
                for (int countsub = 0; countsub < dgv.Columns.Count; countsub++)
                {
                    dr[countsub] = Convert.ToString(dgv.Rows[count].Cells[countsub].Value);
                }
                dt.Rows.Add(dr);
            }
            return dt;
        }

        public static bool DateTimeRemoveTime(System.Data.DataTable dt)//将System.Data.DataTable中时间所在列进行处理，去除时间只留日期(System.Data.DataTable传递的为引用值，类似于地址，结果会影响传入的类，不需要通过返回DataTable来实现改变)
        {
            try
            {
                string str = string.Empty;
                for (int count = 0; count < dt.Rows.Count - 1; count++)//因为最后一行为空数据所以最后一行不处理(count < dt.Rows.Count - 1)，否则最后一行在执行Convert.ToDateTime函数时会出现无法转化的异常
                {
                    str = Convert.ToDateTime(dt.Rows[count][2].ToString().Trim()).ToString("yyyy-MM-dd");//去除时间只留日期
                    dt.Rows[count][2] = str;
                }
                return true;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                //throw new Exception(ex.Message);
                return false;
            }
        }

        public static System.Data.DataTable LoadDataFromExcel(string filePath)//从指定文件路径中的Excel读取所有内容到System.Data.DataTable(xls和xlsx格式均可以打开)
        {
            Excel.Application excel = new Excel.Application();//对象实例化;
            try
            {
                Excel.Workbook workbook = excel.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                Excel.Worksheet worksheet = (Worksheet)workbook.Worksheets[1];

                excel.Visible = false;//不显示Excel内容

                int rowCount = worksheet.UsedRange.Rows.Count;//定义Excel文件中的行数
                int colCount = worksheet.UsedRange.Columns.Count;//定义Excel文件中的列数
                Excel.Range range;
                System.Data.DataTable dt = new System.Data.DataTable();

                for (int i = 0; i < colCount; i++)//循环列
                {
                    range = (Excel.Range)excel.Cells[1, i + 1];
                    dt.Columns.Add(range.Value2.ToString());
                }
                for (int j = 1; j < rowCount; j++)//循环行
                {
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < colCount; i++)
                    {
                        range = (Excel.Range)excel.Cells[j + 1, i + 1];
                        dr[i] = range.Value2.ToString();
                    }
                    dt.Rows.Add(dr);
                }
                excel.DisplayAlerts = false; //设置禁止弹出保存和覆盖的询问提示框
                excel.AlertBeforeOverwriting = false;
                excel.Workbooks.Close();//关闭工作簿
                excel.Quit();//退出Excel程序
                return dt;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        public static bool SaveDataToExcel(System.Data.DataTable dt, string filePath)//将System.Data.DataTable内容存储到指定路径的Excel中(存储默认格式xlsx)
        {
            Excel.Application excel = new Excel.Application();//对象实例化;
            try
            {
                excel.Visible = false;//不显示Excel内容
                Workbook workbook = excel.Workbooks.Add(true);
                Worksheet worksheet = workbook.Worksheets[1] as Worksheet;
                if (dt.Rows.Count > 0)
                {
                    int row = row = dt.Rows.Count;
                    int col = dt.Columns.Count;
                    for (int i = 0; i < row; i++)
                    {
                        for (int j = 0; j < col; j++)
                        {
                            string str = dt.Rows[i][j].ToString();
                            worksheet.Cells[i + 2, j + 1] = str;
                        }
                    }
                }

                int size = dt.Columns.Count;
                for (int i = 0; i < size; i++)
                {
                    worksheet.Cells[1, 1 + i] = dt.Columns[i].ColumnName;
                }

                excel.DisplayAlerts = false; //设置禁止弹出保存和覆盖的询问提示框
                excel.AlertBeforeOverwriting = false;

                workbook.SaveAs(filePath);//保存工作簿(当不指定FileFormat时，保存为默认格式，新建的文档默认格式为xlsx，即使filePath后缀为xls，内部数据格式仍为xlsx)
                workbook.Close();//关闭工作簿
                excel.Quit();//退出Excel程序
                return true;
            }
            catch (Exception ex)//导出Excel出错
            {
                //throw new Exception(ex.Message);
                //MessageBox.Show(ex.Message);
                return false;
            }
        }
    }
}
