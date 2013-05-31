using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Excel = Microsoft.Office.Interop.Excel;

namespace Schedule_Test
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel.Application xlApp;
        private Excel.Workbook xlWb;
        private Excel.Worksheet xlWs;
        private string xlFilePath;

        private Excel.Application targetXlApp;
        private Excel.Workbook targetXlWb;
        private Excel.Worksheet targetXlWs;



        public MainWindow()
        {
            InitializeComponent();

            
        
        }

        private void btnConvertExcel_Click(object sender, RoutedEventArgs e)
        {
            
            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            //openFileDialog.FileName = "Document";
            openFileDialog.DefaultExt = ".xls";
            openFileDialog.Filter = "XLS文件（*.xls）|*.xls";

            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                xlFilePath = openFileDialog.FileName;

                xlApp = new Excel.Application();
                xlWb = xlApp.Workbooks.Open(xlFilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWs = (Excel.Worksheet)xlWb.Worksheets.get_Item(1);

                Excel.Range range;

                range = xlWs.UsedRange;
                //MessageBox.Show("Rows: " + range.Rows.Count + "\n" + "Columns: " + range.Columns.Count);

                SegmentExcelIntoLineOfRange(range, range.Rows.Count, range.Columns.Count);

                //xlWb.Close(true, null, null);
               // xlApp.Quit();
            }
        }

        private void SegmentExcelIntoLineOfRange(Excel.Range range, int rCount, int cCount)
        {
            //先对123.xls文件作一次预处理，把它转成更易于解析的形式
            //一行一行的解析
            //从一行的首列开始，遇null则跳过
            //如遇到的列内容是"教师"，则跳过该行
            //临时存储一行首列的人名，解析其它
            //如遇下一行中没有人名，则使用上边的人名


            int rowCount, colCount;
            string teacherName = "";
            string gradeNum = "";
            string classList = "";
            string[] classArray;
            string[] stringSeparators = new string[] { "、" };
            Excel.Range r;
            //遍历老师一列、年级一列、任课班级一列
            for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                //probarConvertProgress.SetValue(rowCount / rCount);

                r = range.Rows[rowCount];
                //老师一列
                if (r.Columns[1].Value2 != null)
                {
                    teacherName = r.Columns[1].Value2;
                }
                else
                {
                    r.Columns[1].Value2 = teacherName;
                }
                //年级一列
                if (r.Columns[2].Value2 != null)
                {
                    gradeNum = r.Columns[2].Value2;
                    gradeNum = gradeNum.Replace("小", "");
                    gradeNum = gradeNum.Replace("级", "");
                    r.Columns[2].Value2 = gradeNum;
                }
                else
                {
                    r.Columns[2].Value2 = gradeNum;
                }

                //任课班级一列
                if (r.Columns[5].Value2 != null)
                {
                    classList = r.Columns[5].Value2;
                    classList = classList.Replace("小(", "");
                    classList = classList.Replace(")班", "");
                    r.Columns[5].Value2 = classList;

                    
                   
                }
                else
                {//这一部分可能不用要
                    r.Columns[5].Value2 = classList;
                }
                
            }

            //把一位老师在一门课中的多个班级的分成一个班级一行
            Excel.Range tempR;
            object misValue = System.Reflection.Missing.Value;

           for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
               // probarConvertProgress.Value = rowCount / rCount * 0.5;

                //Excel.Range r = range.Rows[rowCount];
                r = range.Rows[rowCount];
                classList = r.Columns[5].Value2.ToString();
                //MessageBox.Show(classList);
               classArray = classList.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
 
                int classArrayLength = classArray.Length;
                if (classArrayLength > 1)
                {
                    for (int arrayI = 0; arrayI < classArrayLength; arrayI++)
                    {
                        //MessageBox.Show("classArray[" + arrayI + "]: " + classArray[arrayI]);

                        if (arrayI == 0)
                        {
                            r.Columns[5].Value2 = classArray[0];

                            /*MessageBox.Show("!!!!"+r.Columns[1].Value2.ToString() + " "
                              + r.Columns[2].Value2.ToString() + " "
                              + r.Columns[3].Value2.ToString() + " "
                              + r.Columns[4].Value2.ToString() + " "
                              + r.Columns[5].Value2.ToString() + " "
                               );*/
                        }
                        else
                        {
                            
                            //新插入的格在当前格之后,在当前格前边插完了就可以继续往下进行了
                            r.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                            tempR = range.Rows[rowCount];
                            tempR.Columns[1].Value2 = r.Columns[1].Value2;
                            tempR.Columns[2].Value2 = r.Columns[2].Value2;
                            tempR.Columns[3].Value2 = r.Columns[3].Value2;
                            tempR.Columns[4].Value2 = r.Columns[4].Value2;
                            tempR.Columns[5].Value2 = classArray[arrayI];

                            rowCount++;
                            //新插入行了，总行数当然也要增加
                            rCount++;
                        }
                    }
                    
                }
                    /*MessageBox.Show(classArray.Length.ToString());*/
                
            }


            /*一行一行的MessageBox.Show()
            for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                Excel.Range r = range.Rows[rowCount];
                MessageBox.Show(r.Columns[1].Value2.ToString() + " "
                              + r.Columns[2].Value2.ToString() + " "
                              + r.Columns[3].Value2.ToString() + " "
                              + r.Columns[4].Value2.ToString() + " "
                              + r.Columns[5].Value2.ToString() + " "
                               );
            }*/

            /*for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                Excel.Range r = range.Rows[rowCount];

                if (r.Columns[5].Value2 != null)
                {
                    MessageBox.Show(r.Columns[5].Value2.ToString());
                }
                else
                {
                    MessageBox.Show("[null]");
                }
            }*/

            /*for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                Excel.Range r = range.Rows[rowCount];

                //123.xls中A号列都为空所以设cCnt的初值为1跳过该列
                for (int cCnt = 1; cCnt < cCount + 1; cCnt++)
                {
                    if (r.Cells[rowCount, cCnt].Value2 == "教师")
                    {
                        break;
                    }
                }
            }*/


            
            //行必须要从1号开始，列可以从0号开始
            /*先把一张表分成一行一行，再读取每行的各个列
            for (rowCount = 1; rowCount < rCount; rowCount++)
            {
                Excel.Range r = range.Rows[1];

                for (int cCnt = 0; cCnt < cCount+1; cCnt++)
                {
                    if(r.Cells[rowCount, cCnt].Value2 == null)
                    {
                        MessageBox.Show("[null]");
                    }
                    else
                    {
                        MessageBox.Show(r.Cells[rowCount, cCnt].Value2.ToString());
                    }
                }
            }*/
            
            xlWb.SaveAs("e:\\saved01.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive);
            xlWb.Close(true);
            xlApp.Quit();
            MessageBox.Show("Done!");
        }

        private void btnCreateExcel_Click(object sender, RoutedEventArgs e)
        {
            object misValue = System.Reflection.Missing.Value;
            string fileSavePath = "";

            targetXlApp = new Excel.Application();
            targetXlWb = targetXlApp.Workbooks.Add();
            targetXlWs = (Excel.Worksheet)targetXlWb.Worksheets.get_Item(1);
            //Cells的行和列都要从1号开始
            targetXlWs.Cells[1, 1] = "年级";
            targetXlWs.Cells[1, 2] = "班级";
            targetXlWs.Cells[1, 3] = "课程";
            targetXlWs.Cells[1, 4] = "教师";
            targetXlWs.Cells[1, 5] = "场地";
            targetXlWs.Cells[1, 6] = "周课时";
            targetXlWs.Cells[1, 7] = "每周连课次数";
            targetXlWs.Cells[1, 8] = "每次连课节数";
            targetXlWs.Cells[1, 9] = "课程性质";
            targetXlWs.Cells[1, 10] = "所在校区";


            Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            saveFileDialog.Filter = "XLS文件（*.xls）|*.xls";
            saveFileDialog.FilterIndex = 2;
            saveFileDialog.RestoreDirectory = true;
            if (saveFileDialog.ShowDialog() == true)
            {
                fileSavePath = saveFileDialog.FileName.ToString();


                targetXlWb.SaveAs(fileSavePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive);
                targetXlWb.Close(true);
                targetXlApp.Quit();
                MessageBox.Show("Excel file Saved!");
            }
        }

        private void ProgressBar_ValueChanged_1(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }


    }
}
