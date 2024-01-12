using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;


namespace DataTable2Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string strSql = "select * from H3AV61MeterMainData";
            DataSet ds = SQLiteHellper.ExecuteQuery(strSql);//导出到Excel的DataTable
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                FolderBrowserDialog folder = new FolderBrowserDialog();
                if (folder.ShowDialog() == DialogResult.OK)
                {
                    string ExportDir = folder.SelectedPath + "\\数据导出" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xls";
                    int i = DataTableToExcel(ds.Tables[0], "测试", ExportDir, true);
                    MessageBox.Show("导出" + i.ToString() + "条记录");
                }
            }
        }



        /// <summary>
        /// 从Excel读取数据
        /// 只支持单表
        /// </summary>
        /// <param name="FilePath">文件路径</param>
        public static DataTable ReadFromExcel(string FilePath)
        {
            DataTable result = null;
            IWorkbook wk = null;
            string extension = System.IO.Path.GetExtension(FilePath); //获取扩展名
            try
            {
                using (FileStream fs = File.OpenRead(FilePath))
                {
                    if (extension.Equals(".xls")) //2003
                    {
                        wk = new HSSFWorkbook(fs);
                    }
                    else                         //2007以上
                    {
                        wk = new XSSFWorkbook(fs);
                    }
                }

                //读取当前表数据
                ISheet sheet = wk.GetSheetAt(0);

                //构建DataTable
                IRow row = sheet.GetRow(0);
                result = BuildDataTable(row);
                if (result != null)
                {
                    if (sheet.LastRowNum >= 1)
                    {
                        for (int i = 1; i < sheet.LastRowNum + 1; i++)
                        {
                            IRow temp_row = sheet.GetRow(i);
                            if (temp_row == null) { continue; }         //修复行为空时的报错
                            List<object> itemArray = new List<object>();
                            for (int j = 0; j < result.Columns.Count; j++)
                            {
                                var str = temp_row.GetCell(j).ToString();
                                if (str.Contains("&"))// { "&","&amp;"},  { "<","&lt;"},   { ">","&gt;"},
                                {
                                  str= str.Replace("&", "&amp;");
                                }
                                if (str.Contains("<"))
                                {
                                    str = str.Replace("<", "&lt;");
                                }
                                if (str.Contains(">"))
                                {
                                    str = str.Replace(">", "&lt;");
                                }
                                Console.WriteLine(str);
                                itemArray.Add(str);                          
                            }
                            result.Rows.Add(itemArray.ToArray());
                        }
                    }
                }

                return result;

            }
            catch (Exception ex)
            {
                return null;
            }
        }


        //&为  &amp;	<为  &lt;		c为  &gt;
        Dictionary<string, string> mDic = new Dictionary<string, string>() {
            { "&","&amp;"},
            { "<","&lt;"},
            { ">","&gt;"},
        };

/// <summary>
/// 构建DataTable框架
/// </summary>
/// <param name="Row">Excel第一行</param>
/// <returns></returns>
private static DataTable BuildDataTable(IRow Row)
        {
            DataTable result = null;
            if (Row.Cells.Count > 0)
            {
                result = new DataTable();
                for (int i = 0; i < Row.LastCellNum; i++)
                {
                    if (Row.GetCell(i) != null)
                    {
                        result.Columns.Add(Row.GetCell(i).ToString());
                    }
                }
            }
            return result;
        }


        /// <summary>
        /// 将DataTable数据导出到excel中
        /// </summary>
        /// <param name="data">要导入的数据</param>
        /// <param name="isColumnWritten">DataTable的列名是否要导入</param>
        /// <param name="sheetName">要导入的excel的sheet的名称</param>
        /// <param name="fileName">导出的途径包含了文件名</param>
        /// <returns>导入数据行数(包含列名那一行)</returns>
        public static int DataTableToExcel(DataTable data, string sheetName, string fileName, bool isColumnWritten)
        {
            int i = 0;
            int j = 0;
            int count = 0;
            ISheet sheet = null;
            IWorkbook workbook = null;

            using (FileStream fs = new FileStream(fileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                if (fileName.IndexOf(".xlsx") > 0) // 2007版本
                    workbook = new XSSFWorkbook();
                else if (fileName.IndexOf(".xls") > 0) // 2003版本
                    workbook = new HSSFWorkbook();

                try
                {
                    Console.WriteLine(data);
                    if (workbook != null)
                    {
                        sheet = workbook.CreateSheet(sheetName);
                    }
                    else
                    {
                        return -1;
                    }

                    if (isColumnWritten == true) //写入DataTable的列名
                    {
                        IRow row = sheet.CreateRow(0);
                        for (j = 0; j < data.Columns.Count; ++j)
                        {

                            row.CreateCell(j).SetCellValue(data.Columns[j].ColumnName);
                        }
                        count = 1;
                    }
                    else
                    {
                        count = 0;
                    }

                    for (i = 0; i < data.Rows.Count; ++i)
                    {
                        IRow row = sheet.CreateRow(count);
                        for (j = 0; j < data.Columns.Count; ++j)
                        {
                            Console.WriteLine(data.Rows[i][j]);
                            Console.WriteLine(data.Rows[i][j].ToString());
                            row.CreateCell(j).SetCellValue(data.Rows[i][j].ToString());

                            Console.WriteLine(System.Security.SecurityElement.Escape(data.Rows[i][j].ToString()));
                        }
                        ++count;
                    }
                    workbook.Write(fs); //写入到excel

                    return count;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                    return -1;
                }
            }
        }

        private static object GetValueType(ICell cell)
        {
            if (cell == null)
                return null;
            switch (cell.CellType)
            {
                case CellType.Blank:
                    return null;
                case CellType.Boolean:
                    return cell.BooleanCellValue;
                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue;
                    }
                    return cell.NumericCellValue;
                case CellType.String:
                    return cell.StringCellValue;
                case CellType.Error:
                    return cell.ErrorCellValue;
                case CellType.Formula:
                    cell.SetCellType(CellType.String);
                    return cell.StringCellValue;
                default:
                    return "=" + cell.CellFormula;
            }
        }

        /// <summary>
        /// 存储到指定位置下
        /// </summary>
        /// <param name="filename">指定位置下文件名</param>
        /// <param name="dt">dataTable转换成Xml</param>
        public static void SaveXml(string filename, DataTable dt)
        {
            string strXml = ConvertDataTableToXml(dt);
            #region 写入数据
            using (FileStream fsWrite = new FileStream(filename, FileMode.OpenOrCreate, FileAccess.Write))
            {
                byte[] buffer = Encoding.UTF8.GetBytes(strXml);
                fsWrite.Write(buffer, 0, buffer.Length);
            }
            #endregion
        }

        //转换xml后的格式 特殊处理的一版
        private static string ConvertDataTableToXml(DataTable dt)
        {
            StringBuilder strXml = new StringBuilder();
            strXml.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            strXml.AppendLine("<resources>");
            string value = "<string";
            //Excule 读行再读列
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                value = "<string";
                //strXml.AppendLine("<HAHAData>");
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                    if (j == 2)
                    {
                        value += ">" + dt.Rows[i][j];
                    }
                    else
                    {
                        value += " " + dt.Columns[j].ColumnName + "=\"" + dt.Rows[i][j] + "\"";
                    }
                }
                value += "</string>";
                strXml.AppendLine(value);
                //strXml.AppendLine("</HAHAData>");//下一行的内容
            }
            strXml.AppendLine("</resources>");
            return strXml.ToString();
        }

        /// <summary>
        /// 读取xml文件转换成DataTable
        /// </summary>
        /// <param name="fileName">指定位置下文件名</param>
        /// <returns></returns>
        public static DataTable ConvertXmlToDataTable(string fileName)
        {
            DataTable dt = null;
            DataSet DS = new DataSet();
            DS.ReadXml(fileName);
            dt = DS.Tables[0];
            return dt;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string strSql = "select * from H3AV61MeterMainData";//导出到Excel的DataTable
            DataSet ds = SQLiteHellper.ExecuteQuery(strSql);
            if (ds != null && ds.Tables.Count > 0 && ds.Tables[0].Rows.Count > 0)
            {
                FolderBrowserDialog folder = new FolderBrowserDialog();
                if (folder.ShowDialog() == DialogResult.OK)
                {
                    string ExportDir = folder.SelectedPath + "\\数据导出" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
                    int i = DataTableToExcel(ds.Tables[0], "测试", ExportDir, true);
                    MessageBox.Show(i.ToString());
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("暂无用");
            return;
            OpenFileDialog openFile = new OpenFileDialog();
            if (openFile.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFile.FileName;
                DataTable excelDt = ReadFromExcel(filePath);
                //MessageBox.Show("导入" + excelDt.Rows.Count + "条");
            }
        }

        public void button4_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(AppDomain.CurrentDomain.BaseDirectory);
            DataTable dataTabale = null;
            string path = AppDomain.CurrentDomain.BaseDirectory + "en_fui.xlsx";
            if (File.Exists(path) == false)
            {
                MessageBox.Show("同目录下 必须要有 en_fui.xlsx 文件");
                return;
            }
            dataTabale = ReadFromExcel(path);
       
            string path2 = AppDomain.CurrentDomain.BaseDirectory + "en_fui.xml";
            if (File.Exists(path2))
            {
                MessageBox.Show("同目录下 必须删掉 en_fui.xml 文件");
                return;
            }
            SaveXml(path2, dataTabale);
            MessageBox.Show("成功生成XML文件 可以导入fgui了");
        }


        public void btnReadXml_Click(object sender, EventArgs e)
        {
            //MessageBox.Show(AppDomain.CurrentDomain.BaseDirectory);
            string path2 = AppDomain.CurrentDomain.BaseDirectory + "CN_fui.xml";
            if (File.Exists(path2) == false)
            {
                MessageBox.Show("同目录下 必须要有 CN_fui.xml文件");
                return;
            }
            string fileName = AppDomain.CurrentDomain.BaseDirectory + "CN_fui.xlsx";
            if (File.Exists(fileName))
            {
                MessageBox.Show("同目录下 必须删掉 CN_fui.xlsx 文件");
                return;
            }

            DataTable dt = ConvertXmlToDataTable(path2);
         
            DataTableToExcel(dt, "hello", fileName, true);
            MessageBox.Show("成功生成Excel ");
        }
    }
}
