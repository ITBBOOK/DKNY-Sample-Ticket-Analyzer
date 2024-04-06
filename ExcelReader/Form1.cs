using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.Dynamic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Window;
using Application = Microsoft.Office.Interop.Excel.Application;
using DataTable = System.Data.DataTable;

namespace ExcelReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
             
            InitializeComponent(); 
            System.Windows.Forms.Control.CheckForIllegalCrossThreadCalls = false;
        }

         
        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Multiselect = false;//不允许打开多个文件

            openFileDialog.Filter = "Excel文件|*.xls;*.xlsx;*.xlsb;*.xlsm";
            openFileDialog.Title = "请选择一个Excel文件";

            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (dataGridView1.Rows.Count > 0)
                {

                    dataGridView1.Rows.Clear();
                }
               
                string filePath = openFileDialog.FileName;
                textBox1.Text = filePath;
                Application excelApp = new Application();

                // 打开工作簿
                Workbook workbook = excelApp.Workbooks.Open(filePath);

                try
                {
                      

                    // 遍历每个工作表
                    foreach (Worksheet worksheet in workbook.Sheets)
                    {
                        if (worksheet.Visible == XlSheetVisibility.xlSheetVisible)
                        {
                            // 将工作表表名添加到 DataTable
                            dataGridView1.Rows.Add(false, worksheet.Name);

                        }

                    }

                    // 将 DataTable 绑定到 dataGridView1
                    //dataGridView1.DataSource = table;
                    //dataGridView1.Columns[1].Width = 300;//设置列宽度
                }
                finally
                {
                    // 关闭工作簿和Excel应用程序
                    workbook.Close(false);
                    excelApp.Quit();
                }

                 

               
                // 释放资源
                System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                // 将引用置为空，确保对象被垃圾回收
                excelApp = null;

                // 强制垃圾回收
                System.GC.Collect();
                System.GC.WaitForPendingFinalizers();
            }
        }

        bool readStart = false;
        private void button2_Click(object sender, EventArgs e)
        {
            if (this.button2.Text == "开始")
            {
                readStart = true;
                this.button2.Text = "停止";

                Thread t = new Thread(new ThreadStart(ReadSheet));
                t.Start();


            }
            else if (this.button2.Text == "停止")
            {
                readStart = false;

            }
        }


        private void ReadSheet()
        {
            string excelPath = textBox1.Text;
            List<string> selectedSheets = new List<string>();
            DataTable allInfo = new DataTable();
             ;
            if (dataGridView2.Rows.Count > 0)
            {
                OpeMainFormControl(delegate ()
                {

                    dataGridView2.DataSource = null;
                    //dataGridView2.Refresh();
                });
            }

             

            // 遍历 dataGridView1 中的行
            foreach (DataGridViewRow drow in dataGridView1.Rows)
            {
                // 检查 Selected 列中的勾选状态
                DataGridViewCheckBoxCell checkBoxCell = drow.Cells["Column1"] as DataGridViewCheckBoxCell;
                if (checkBoxCell != null && Convert.ToBoolean(checkBoxCell.Value))
                {
                    // 如果被勾选，读取 ID 和 Name 列的内容

                    string sheetName = drow.Cells["Column2"].Value.ToString();
                    selectedSheets.Add(sheetName);
                }
            }

            allInfo.Columns.Add("ITEM");
            allInfo.Columns.Add("Category");
            allInfo.Columns.Add("SheetName");
            allInfo.Columns.Add("Group");
            allInfo.Columns.Add("Theme/Body Material");
            allInfo.Columns.Add("Style Name");
            // allInfo.Columns.Add("Hardware");
            allInfo.Columns.Add("Factory");
            allInfo.Columns.Add("Color");
            allInfo.Columns.Add("Style Number");
            DataColumn NY = new DataColumn($"NY", typeof(string));
            DataColumn MILAN = new DataColumn($"MILAN", typeof(string));
            DataColumn NY_AGENT = new DataColumn($"NY AGENT", typeof(string));
            DataColumn EU_AGENT = new DataColumn($"EU AGENT", typeof(string));

            allInfo.Columns.Add(NY);
            allInfo.Columns.Add(MILAN);
            allInfo.Columns.Add(NY_AGENT);
            allInfo.Columns.Add(EU_AGENT);
             
            int colNum = 9;



            Application excelApp = new Application(); 
            Workbook workbook = excelApp.Workbooks.Open(excelPath);
            if (selectedSheets.Count > 0)
            {
                int stylesNum = 1;
                 
                foreach (string sheetname in selectedSheets)
                {
                    
                    Worksheet sheet = (Worksheet)workbook.Sheets[sheetname];
                    
                    try
                    {
                        //  Range StyleDes = SearchCellAddress(sheet, "STYLE DESCRIPTION", 0);
                        Range StyleDes = SearchCellAddressBetweenColumns(sheet, "STYLE DESCRIPTION", 0, 1, 3);
                        // Range StyleDes = SearchCellAddressBetweenCells(sheet, "STYLE DESCRIPTION", 0, "A1", "C65");

                        string[] styleNames = new string[10];


                        // Range StyleNumber = SearchCellAddress(sheet, "STYLE  NUMBER", 0);
                        Range StyleNumber = SearchCellAddressBetweenColumns(sheet, "STYLE  NUMBER", 0, 1, 3);
                        string[] styleNums = new string[10];
                        for (int i = 0; i < styleNums.Length; i++)
                        {
                            styleNums[i] = GetStringByCell(StyleNumber.Offset[0, i * 4 + 1]);

                        }

                        //Range groupCell = SearchCellAddress(sheet, "GROUP NAME", 0);
                        Range groupCell = SearchCellAddressBetweenColumns(sheet, "GROUP NAME", 0, 10, 15);
                        string groupName = GetStringByCell(groupCell.Offset[0, 1]);

                        // Range themeCell = SearchCellAddress(sheet, "THEME", 0);
                        Range themeCell = SearchCellAddressBetweenColumns(sheet, "THEME", 0, 27, 33);
                        string theme = GetStringByCell(themeCell.Offset[0, 1]);

                        // Range factoryCell = SearchCellAddress(sheet, "FACTORY", 0);
                        Range factoryCell = SearchCellAddressBetweenColumns(sheet, "FACTORY", 0, 38, 41);
                        string factory = GetStringByCell(factoryCell.Offset[0, 1]);

                        //Range categoryCell = SearchCellAddress(sheet, "CATEGORY", 0);
                        Range categoryCell = SearchCellAddressBetweenColumns(sheet, "CATEGORY", 0, 27, 33);
                        string category = GetStringByCell(categoryCell.Offset[0, 1]);



                        string colorName = "";
                        //Range sampleColor = SearchCellAddress(sheet, "SAMPLE COLOR", 0);
                        Range sampleColor = SearchCellAddressBetweenColumns(sheet, "SAMPLE COLOR", 0, 1, 3);
                        List<string> colorNames = new List<string>();
                        bool haveQTY = false;
                        bool haveColor = false;





                        for (int j = 0; j < styleNames.Length; j++) //int j = 0; j < styleNames.Length; j++
                        {


                            styleNames[j] = GetStringByCell(StyleDes.Offset[0, j * 4 + 1]);

                            if (styleNames[j] != "" && styleNames[j] != null && styleNames[j] != "0")

                            {

                                for (int i = 0; i < 15; i++)
                                {

                                    colorName = GetStringByCell(sampleColor.Offset[i + 1, 0]);

                                    if (colorName != "" && colorName != null && colorName != "0")

                                    {
                                        if (colorName.Equals("SAMPLE ORDER TOTAL", StringComparison.OrdinalIgnoreCase))
                                        {
                                            break;
                                        }
                                        haveColor = true;
                                        colorNames.Add(colorName);
                                        string qty = "";
                                        DataRow dr = allInfo.NewRow();
                                        dr[0] = stylesNum;
                                        dr[1] = category;
                                        dr[2] = sheetname;
                                        dr[3] = groupName;
                                        dr[4] = theme;
                                        dr[5] = styleNames[j];
                                        dr[6] = factory;
                                        dr[7] = colorName;
                                        dr[8] = styleNums[j];

                                        for (int m = 0; m < 4; m++)

                                        {
                                            qty = GetStringByCell(sampleColor.Offset[i + 1, j * 4 + m + 1]);

                                            dr[colNum + m] = qty;

                                        }

                                        allInfo.Rows.Add(dr);


                                    }

                                }

                                stylesNum++;
                            }






                        }

                    }

                    catch (Exception ex)

                    {
                        MessageBox.Show(ex.Message + "\\" + ex.StackTrace);
                        return;
                    }

                }

                try
                {
                    OpeMainFormControl(delegate ()
                    {

                        dataGridView2.DataSource = allInfo;
                        //dataGridView2.Refresh();
                    });
                    dataGridView2.DataSource = allInfo;
                }
                catch (Exception ex)

                {
                    MessageBox.Show(ex.Message + "\\" + ex.StackTrace);
                    return;
                }

                // dataGridView2.Columns[0].AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells;



            }


            excelApp.Workbooks.Close();
            // 关闭 Excel 应用程序
            excelApp.Quit();

            // 释放 COM 对象，防止内存泄漏
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);

            // 将引用置为空，确保对象被垃圾回收
            excelApp = null;

            // 强制垃圾回收
            System.GC.Collect();
            System.GC.WaitForPendingFinalizers();


            try
            {



                Application excelApp2 = new Application();
                excelApp2.Visible = true;

                // 添加一个工作簿
                Workbook workbook2 = excelApp2.Workbooks.Add();

                // 在工作簿中添加一个工作表
                Worksheet worksheet2 = (Worksheet)workbook2.Worksheets[1];

                // 将 DataTable 数据写入 Excel 工作表
                int row = 1;
                int col = 1;

                // 写入列名
                foreach (DataColumn column in allInfo.Columns)
                {
                    worksheet2.Cells[row, col] = column.ColumnName;
                    col++;
                }

                // 写入数据
                row++;
                foreach (DataRow dataRow in allInfo.Rows)
                {
                    col = 1;
                    foreach (var item in dataRow.ItemArray)
                    {
                        worksheet2.Cells[row, col] = item;
                        col++;
                    }
                    row++;
                }

                DateTime currentTime = DateTime.Now;

                // 提取秒数并格式化为字符串
                string secondsString = currentTime.ToString("ss");


                string saveExcel = System.Windows.Forms.Application.StartupPath + "\\Sample Ticket Data_" + currentTime.ToString("yyyyMMddHHmm") + ".xlsx";

                // 保存 Excel 文件
                worksheet2.SaveAs(saveExcel);
            }
            catch (Exception ex) 
            {
                MessageBox.Show(ex.Message + "\\" + ex.StackTrace);
                return;
            }

            // 关闭 Excel 应用程序
            //excelApp2.Quit();


            this.button2.Text = "开始";







        }
        private void OpeMainFormControl(System.Action action)
        {
            if (this.InvokeRequired)
                this.Invoke(action); //返回主线程（创建控件的线程）
            else
                action();
        }

        private static Range FindCell(Worksheet sheet, string targetText, int startRow, int endRow)
        {
            foreach (Range cell in sheet.Range[sheet.Cells[startRow, 1], sheet.Cells[endRow, sheet.Columns.Count]].Cells)
            {
                if (cell.Value != null && cell.Value.ToString().IndexOf(targetText, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return cell;
                }
            }

            return null;
        }

        private string SearchCell(Workbook workbook, string sheetname, string cellName, int offset, int start, int end)
        {
            Worksheet worksheetName = (Worksheet)workbook.Sheets[sheetname];
            Range groupNameCell = FindCell(worksheetName, cellName, start, end);
            string groupName = "";
            if (groupNameCell != null)
            {
                // 读取目标单元格右边一格的内容


                Range adjacentCell = worksheetName.Cells[groupNameCell.Row, groupNameCell.Column + offset];

                if (adjacentCell.Value == null && adjacentCell.MergeCells)
                {
                    Range mergedArea = groupNameCell.Offset[0, offset].MergeArea;

                    // groupName = mergedArea.Value != null ? mergedArea.Value.ToString() : "";

                    foreach (Range cell in mergedArea)
                    {
                        if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                        {
                            groupName = cell.Value.ToString();
                        }
                    }


                }

                else
                {
                    groupName = adjacentCell.Value.ToString();
                }



            }
            else
            {

            }

            return groupName;
        }

        private string GetStringByCell(Range cell)

        {
            try
            {
                
                 

                // 处理合并单元格的情况
                if (cell.MergeCells)
                {
                    var mergedCell = cell.MergeArea;
                    object[,] values = (object[,])mergedCell.Value2;

                    // 将值转换为字符串
                    string cellValue = string.Empty;
                    for (int i = 1; i <= values.GetLength(0); i++)
                    {
                        for (int j = 1; j <= values.GetLength(1); j++)
                        {
                            if (values[i, j] != null)
                            {
                                cellValue += values[i, j].ToString() + " ";
                            }
                        }
                    }

                    return cellValue.Trim();


                }
                else
                {
                    string stringValue = cell.Value != null ? cell.Value.ToString() : "";

                    return stringValue;
                }
            }

            catch { return null; }

               


        }
        private bool IsCellInRange(Range cell, Range startCell, Range endCell)
        {
            return cell.Row >= startCell.Row && cell.Row <= endCell.Row &&
                   cell.Column >= startCell.Column && cell.Column <= endCell.Column;
        }

        private Range SearchCellAddressBetweenCells(Worksheet sheet, string searchString, int cellNum, string startCellAddress, string endCellAddress)
        {
            List<Range> cellList = new List<Range>();
            try
            {
                Range usedRange = sheet.UsedRange;

                // 获取起始单元格和结束单元格
                Range startCell = usedRange.Range[startCellAddress];
                Range endCell = usedRange.Range[endCellAddress];

                foreach (Range cell in usedRange)
                {
                    if (cell.Value2 != null && cell.Value2.ToString().IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // 检查单元格是否在指定范围内
                        if (IsCellInRange(cell, startCell, endCell))
                        {
                            cellList.Add(cell);
                        }
                    }
                }

                if (cellList.Count > cellNum)
                {
                    return cellList[cellNum];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }
        private Range SearchCellAddressBetweenColumns(Worksheet sheet, string searchString, int cellNum, int startColumn, int endColumn)
        {
            List<Range> cellList = new List<Range>();
            try
            {
                Range usedRange = sheet.UsedRange;
                int rowCount = 65;// usedRange.Rows.Count;

                for (int col = startColumn; col <= endColumn; col++)
                {
                    for (int row = 1; row <= rowCount; row++)
                    {
                        object cellValue = usedRange.Cells[row, col].Value2;

                        if (cellValue != null && cellValue.ToString().IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            cellList.Add(usedRange.Cells[row, col] as Range);
                        }
                    }
                }

                if (cellList.Count > cellNum)
                {
                    return cellList[cellNum];
                }
                else
                {
                    return null;
                }
            }
            catch (Exception ex)
            {
                return null;
            }
        }


        private Range SearchCellAddress(Worksheet sheet, string searchString, int cellNum )

        {
           
            List<Range> cellList = new List<Range>();
            try
            {
               
                // 获取第一个工作表

                Range usedRange = sheet.UsedRange;
                foreach (Range cell in usedRange)
                {

                    if (cell.Value != null && cell.Value.ToString().IndexOf(searchString, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // 如果找到包含特定字符串的单元格，返回其地址
                        cellList.Add(cell);
                    }

                }
                     
                 

                if (cellList.Count > cellNum)
                {
                    return cellList[cellNum];
                }
               
                else { return null; }
            }
            catch(Exception ex)
            {
                return null;
            }


        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                int count = dataGridView1.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[0];
                    Boolean flag = Convert.ToBoolean(checkCell.Value);
                    if (flag == false)
                    {
                        checkCell.Value = true;
                    }
                    else
                    {
                        continue;
                    }
                }
            }
            else
            {
                int count = dataGridView1.Rows.Count;
                for (int i = 0; i < count; i++)
                {
                    DataGridViewCheckBoxCell checkCell = (DataGridViewCheckBoxCell)dataGridView1.Rows[i].Cells[0];
                    Boolean flag = Convert.ToBoolean(checkCell.Value);
                    if (flag == true)
                    {
                        checkCell.Value = false;
                    }
                    else
                    {
                        continue;
                    }
                }

            }
        }
    }
}
