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

using System.ComponentModel;

using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.WindowsAPICodePack.Dialogs;

namespace Schedule_Test
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private BackgroundWorker backgroundWorker;
        private BackgroundWorker allSchoolScheduleBW;

        private Excel.Application xlApp;
        private Excel.Workbook xlWb;
        private Excel.Worksheet xlWs;
        private string xlFilePath;

        private Excel.Application targetXlApp;
        private Excel.Workbook targetXlWb;
        private Excel.Worksheet targetXlWs;

        private Excel.Application allSchoolScheduleApp;
        private Excel.Workbook allSchoolScheduleWb;
        private Excel.Worksheet allSchoolScheduleWs;

        private Excel.Workbook resultScheduleWb;
        private Excel.Worksheet resultScheduleWs;
        private int resultScheduleLineNum;


       // private CommonOpenFileDialog openFolderDialog;// = new CommonOpenFileDialog();

        private object misValue;


        public MainWindow()
        {
            InitializeComponent();

            Version systemVersion = System.Environment.OSVersion.Version;
            if (systemVersion.Major < 6)
            {
                MessageBox.Show("哦~~~你的系统太老了，请使用Windows Vista及以上系统版本");
                this.Close();//关闭程序，没试过，有空试试能不能用
            }

            backgroundWorker = new BackgroundWorker();
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
            backgroundWorker.DoWork += new DoWorkEventHandler(backgroundWorker1_DoWork);
            backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(backgroundWorker1_ProgressChanged);
            backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(backgroundWorker1_RunWorkerCompleted);

            allSchoolScheduleBW = new BackgroundWorker();
            allSchoolScheduleBW.WorkerReportsProgress = true;
            allSchoolScheduleBW.WorkerSupportsCancellation = true;
            allSchoolScheduleBW.DoWork += new DoWorkEventHandler(allSchoolScheduleBW_DoWork);
            allSchoolScheduleBW.ProgressChanged += new ProgressChangedEventHandler(allSchoolScheduleBW_ProgressChanged);
            allSchoolScheduleBW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(allSchoolScheduleBW_RunWorkerCompleted);

            misValue = System.Reflection.Missing.Value;
        }

        private void btnConvertExcel_Click(object sender, RoutedEventArgs e)
        {
            if (backgroundWorker.IsBusy != true)
            {
                backgroundWorker.RunWorkerAsync();
            }
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            //下边这个变量是用来取消bakcgroundworker用的
            //BackgroundWorker worker = sender as BackgroundWorker;

            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog();
            CommonFileDialogFilter filter = new CommonFileDialogFilter("*.xls文件", ".xls");
            openFileDialog.Filters.Add(filter);
            CommonFileDialogResult commonFileDialogResult = CommonFileDialogResult.None;

            App.Current.Dispatcher.Invoke(new Action(() =>
            {
                commonFileDialogResult = openFileDialog.ShowDialog();
            }));

            //openFolderDialog = new CommonOpenFileDialog();

            //Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            //openFileDialog.FileName = "Document";
            //openFileDialog.DefaultExt = ".xls";
            //openFileDialog.Filter = "XLS文件（*.xls）|*.xls";

           // Nullable<bool> result = openFileDialog.ShowDialog();
            //if (result == true)
            if (commonFileDialogResult == CommonFileDialogResult.Ok) 
            {
                backgroundWorker.ReportProgress(0);

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



        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //resultLabel.Content = (e.ProgressPercentage.ToString() + "%");
            if (e.ProgressPercentage == 0)
            {
                probarConvertProgress.IsIndeterminate = true;
            }
            else if(e.ProgressPercentage == 100)
            {
                probarConvertProgress.IsIndeterminate = false;
            }
          //MessageBox.Show(e.ProgressPercentage.ToString());
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
           /* if (e.Cancelled == true)
            {
                //resultLabel.Content = "Canceled!";
                MessageBox.Show("Canceled!");
            }
            else if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
               // resultLabel.Content = "Error: " + e.Error.Message;
            }
            else
            {
                MessageBox.Show("Done!");
                //resultLabel.Content = "Done!";
            }*/
        }

        private void SegmentExcelIntoLineOfRange(Excel.Range range, int rCount, int cCount)
        {
            //先对123.xls文件作一次预处理，把它转成更易于解析的形式
            //一行一行的解析
            //从一行的首列开始，遇null则跳过
            //如遇到的列内容是"教师"，则跳过该行
            //临时存储一行首列的人名，解析其它
            //如遇下一行中没有人名，则使用上边的人名

            
            int rowCount;//, colCount;
            string teacherName = "";
            string gradeNum = "";
            string classList = "";
            string[] classArray;
            string[] stringSeparators = new string[] { "、" };
            Excel.Range r;
            //遍历老师一列、年级一列、任课班级一列
            for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                

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

            int timesPerWeek = 0;
           for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                //probarConvertProgress.Value =  40 + rowCount / rCount * 0.6;

                //Excel.Range r = range.Rows[rowCount];
                r = range.Rows[rowCount];
                classList = r.Columns[5].Value2.ToString();
                //MessageBox.Show(classList);
                classArray = classList.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
 
                int classArrayLength = classArray.Length;
                if (classArrayLength > 1)
                {
                    //周总课时
                    timesPerWeek = (int)((int)(r.Columns[4].Value2 / classArrayLength));

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
                            //周总课时
                            //timesPerWeek = (int)((int)(r.Columns[4].Value2 / classArrayLength));
                            tempR.Columns[4].Value2 = timesPerWeek;
                            //这句话效率低，如果有二个以上班级，下边一句会冗余多次执行
                            range.Rows[rowCount+1].Columns[4].Value2 = timesPerWeek;

                            tempR.Columns[5].Value2 = classArray[arrayI];

                            rowCount++;
                            //新插入行了，总行数当然也要增加
                            rCount++;

                            //if (rowCount == rCount - 1)
                            //{
                            //    range.Rows[rowCount + 1].Columns[4].Value2 = timesPerWeek;
                            //}
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
            
            //xlWb.SaveAs("e:\\saved01.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive);
           // xlWb.Close(true);
            //xlApp.Quit();
            range = xlWs.UsedRange;
            //MessageBox.Show("count: " + range.Rows.Count.ToString());
            CreateTargetExcel(range, range.Rows.Count);

            //MessageBox.Show("Done!");
        }

        private void CreateTargetExcel(Excel.Range range, int rCount)
        {
            string fileSavePath = "";

            targetXlApp = new Excel.Application();

            targetXlWb = targetXlApp.Workbooks.Add();
            resultScheduleWb = targetXlApp.Workbooks.Add();

            targetXlWs = (Excel.Worksheet)targetXlWb.Worksheets.get_Item(1);
            resultScheduleWs = (Excel.Worksheet)resultScheduleWb.Worksheets.get_Item(1);

            /* 生成“教学计划” */
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

            


            int rowCount;
            Excel.Range r;
            System.DateTime dateTime = System.DateTime.Now;
            //yearGrad减去201X级的序号就可以得到当前是几年级
            int yearGrad = dateTime.Year+6;
            int grade = 0;
            string gradeString = "";

            //令rowCount的初值为2跳过静养一行,"rCount - 1"是要去掉表尾的合计一行
            for (rowCount = 2; rowCount < rCount - 1; rowCount++)
            {
                r = range.Rows[rowCount];
                
                //年级
                grade = yearGrad - Int32.Parse(r.Columns[2].Value2.ToString());

                switch (grade)
                {
                    case 6:
                        gradeString = "六年级";
                        targetXlWs.Cells[rowCount, 1] = "六年级";
                        //resultScheduleWs.Cells[rowCount, 1] = "六年级";
                        break;
                    case 5:
                        gradeString = "五年级";
                        targetXlWs.Cells[rowCount, 1] = "五年级";
                        //resultScheduleWs.Cells[rowCount, 1] = "五年级";
                        break;
                    case 4:
                        gradeString = "四年级";
                        targetXlWs.Cells[rowCount, 1] = "四年级";
                        //resultScheduleWs.Cells[rowCount, 1] = "四年级";
                        break;
                    case 3:
                        gradeString = "三年级";
                        targetXlWs.Cells[rowCount, 1] = "三年级";
                        //resultScheduleWs.Cells[rowCount, 1] = "三年级";
                        break;
                    case 2:
                        gradeString = "二年级";
                        targetXlWs.Cells[rowCount, 1] = "二年级";
                        //resultScheduleWs.Cells[rowCount, 1] = "二年级";
                        break;
                    case 1:
                        gradeString = "一年级";
                        targetXlWs.Cells[rowCount, 1] = "一年级";
                        //resultScheduleWs.Cells[rowCount, 1] = "一年级";
                        break;
                }

                //班级
                targetXlWs.Cells[rowCount, 2] = gradeString + "(" + r.Columns[5].Value2.ToString() + ")";
                //resultScheduleWs.Cells[rowCount, 2] = gradeString + "(" + r.Columns[5].Value2.ToString() + ")";
                
                //课程
                targetXlWs.Cells[rowCount, 3] = r.Columns[3].Value2.ToString();
                //resultScheduleWs.Cells[rowCount, 3] = r.Columns[3].Value2.ToString();

                //教师
                targetXlWs.Cells[rowCount, 4] = r.Columns[1].Value2.ToString();
                //resultScheduleWs.Cells[rowCount, 4] = r.Columns[1].Value2.ToString();

                //周课时
                targetXlWs.Cells[rowCount, 6] = r.Columns[4].Value2.ToString();

                //所在校区
                targetXlWs.Cells[rowCount, 10] = "人和街小学";

                //resultScheduleWs.Cells[rowCount, 5] = "自动";
               // resultScheduleWs.Cells[rowCount, 6] = "星期X";
               // resultScheduleWs.Cells[rowCount, 7]


            }

            CommonOpenFileDialog openFolderDialog = new CommonOpenFileDialog();
            openFolderDialog.IsFolderPicker = true;
            openFolderDialog.Title = "选择文件保存目录";
            //openFolderDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
            CommonFileDialogResult commonFileDialogResult = CommonFileDialogResult.None;
            App.Current.Dispatcher.Invoke(new Action(() =>
            {
                commonFileDialogResult = openFolderDialog.ShowDialog();
            }));


            if (commonFileDialogResult == CommonFileDialogResult.Ok)
            {
                fileSavePath = openFolderDialog.FileName;

                targetXlWb.SaveAs(fileSavePath+"教学计划"+dateTime.Year+dateTime.Month+dateTime.Day+dateTime.Hour+dateTime.Minute+".xls",
                                  Excel.XlFileFormat.xlWorkbookNormal,
                                  misValue,
                                  misValue,
                                  misValue,
                                  misValue,
                                  Excel.XlSaveAsAccessMode.xlExclusive);

                /*resultScheduleWb.SaveAs(fileSavePath + "排课结果" + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls",
                                  Excel.XlFileFormat.xlWorkbookNormal,
                                  misValue,
                                  misValue,
                                  misValue,
                                  misValue,
                                  Excel.XlSaveAsAccessMode.xlExclusive);*/
                targetXlWb.Close(false);
                //resultScheduleWb.Close(false);
                targetXlApp.Quit();
            }


            //Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
            //Microsoft.Win32.f
            //saveFileDialog.Filter = "XLS文件（*.xls）|*.xls";
            //saveFileDialog.FilterIndex = 2;
            //saveFileDialog.RestoreDirectory = true;
            /*if (saveFileDialog.ShowDialog() == true)
            {
                fileSavePath = saveFileDialog.FileName.ToString();


                targetXlWb.SaveAs(fileSavePath, 
                                  Excel.XlFileFormat.xlWorkbookNormal, 
                                  misValue, 
                                  misValue, 
                                  misValue, 
                                  misValue, 
                                  Excel.XlSaveAsAccessMode.xlExclusive);
                targetXlWb.Close(false);
                targetXlApp.Quit();
                //MessageBox.Show("Excel file Saved!");
            }
            */
            xlWb.Close(false);
            xlApp.Quit();

            //probarConvertProgress.IsIndeterminate = false;


            backgroundWorker.ReportProgress(100);
            MessageBox.Show("All work Done!");
        }

        private void ReadAllSchoolSchedule()
        {
            //在rhjxx33.xls中每个人对应的一行单独制成一个WorkSheet(可怜的表呀，第一范式也不满足)
            //NO!!!!!!不制成WorkSheet使用List<List<string>>即二维字符串List
            //不对，用数组什么的就不能使行或列的表头有意义了，还有是用
            //嗯。。。找到的更好的解决办法Dictionary
            CommonOpenFileDialog openFileDialog = new CommonOpenFileDialog();
            CommonFileDialogFilter filter = new CommonFileDialogFilter("*.xls文件", ".xls");
            openFileDialog.Filters.Add(filter);
            CommonFileDialogResult commonFileDialogResult = CommonFileDialogResult.None;

            App.Current.Dispatcher.Invoke(new Action(() =>
            {
                commonFileDialogResult = openFileDialog.ShowDialog();
            }));

            if (commonFileDialogResult == CommonFileDialogResult.Ok)
            {
                xlFilePath = openFileDialog.FileName;

                allSchoolScheduleApp = new Excel.Application();
                allSchoolScheduleWb = allSchoolScheduleApp.Workbooks.Open(xlFilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                allSchoolScheduleWs = (Excel.Worksheet)allSchoolScheduleWb.Worksheets.get_Item(1);

                Excel.Range range;

                range = allSchoolScheduleWs.UsedRange;

                int rowCount, rCount, cCount;

                rCount = range.Rows.Count;
                cCount = range.Columns.Count;
                MessageBox.Show("rCount: " + rCount +
                                "cCount: " + cCount);

                //这里先试试一个人的
               // Dictionary<string, string> dicSchedule = new Dictionary<string, string>();
               // rowCount = 4;
                string cellString = "";
                //string tempGradeClass = "";
                //string tempSubject = "";
                string[] tempSplitArry;
                int tempWeek;
                int tempYear = System.DateTime.Now.Year;
                Excel.Range tempResultRange;

               // targetXlApp = new Excel.Application();

               // targetXlWb = targetXlApp.Workbooks.Add();
                //resultScheduleWb = targetXlApp.Workbooks.Add();

                //targetXlWs = (Excel.Worksheet)targetXlWb.Worksheets.get_Item(1);
               // resultScheduleWs = (Excel.Worksheet)resultScheduleWb.Worksheets.get_Item(1);

                resultScheduleWb = allSchoolScheduleApp.Workbooks.Add();
                resultScheduleWs = (Excel.Worksheet)resultScheduleWb.Worksheets.get_Item(1);
                /* 生成“排课结果” */
                resultScheduleWs.Cells[1, 1] = "年级";
                resultScheduleWs.Cells[1, 2] = "班级";
                resultScheduleWs.Cells[1, 3] = "课程";
                resultScheduleWs.Cells[1, 4] = "教师";
                resultScheduleWs.Cells[1, 5] = "场地";
                resultScheduleWs.Cells[1, 6] = "星期";
                resultScheduleWs.Cells[1, 7] = "节次";
                resultScheduleLineNum = 2;

                for (rowCount = 4; rowCount < rCount; rowCount++ )
                {
                    //colCount = 2是为了避过姓名一列,colCount是从1号开始的cCount要加一
                    for (int colCount = 2; colCount < cCount + 1; colCount++)
                    {

                        if (range.Rows[rowCount].Columns[colCount].Value2 != null)
                        {
                            cellString = range.Rows[rowCount].Columns[colCount].Value2.ToString();
                            cellString = cellString.Replace("\n", ".");
                            tempSplitArry = cellString.Split('.');

                            tempResultRange = resultScheduleWs.Rows[resultScheduleLineNum];

                            //写入“排课结果的年级、班级两列
                            switch (tempSplitArry[0])
                            {
                                case "小六":
                                    tempResultRange.Columns[1].Value2 =
                                        "六年级";
                                    tempResultRange.Columns[2].Value2 =
                                        "六年级(" + tempSplitArry[1] + ")";
                                    //tempGradeClass = ((tempYear + 6) - 6) + tempSplitArry[1];
                                    break;
                                case "小五":
                                    tempResultRange.Columns[1].Value2 =
                                        "五年级";
                                    tempResultRange.Columns[2].Value2 =
                                        "五年级(" + tempSplitArry[1] + ")";
                                    //tempGradeClass = ((tempYear + 6) - 5) + tempSplitArry[1];
                                    break;
                                case "小四":
                                    tempResultRange.Columns[1].Value2 =
                                        "四年级";
                                    tempResultRange.Columns[2].Value2 =
                                        "四年级(" + tempSplitArry[1] + ")";
                                    //tempGradeClass = ((tempYear + 6) - 4) + tempSplitArry[1];
                                    break;
                                case "小三":
                                    tempResultRange.Columns[1].Value2 =
                                        "三年级";
                                    tempResultRange.Columns[2].Value2 =
                                        "三年级(" + tempSplitArry[1] + ")";
                                    //tempGradeClass = ((tempYear + 6) - 3) + tempSplitArry[1];
                                    break;
                                case "小二":
                                    tempResultRange.Columns[1].Value2 =
                                        "二年级";
                                    tempResultRange.Columns[2].Value2 =
                                        "二年级(" + tempSplitArry[1] + ")";
                                    //tempGradeClass = ((tempYear + 6) - 2) + tempSplitArry[1];
                                    break;
                                case "小一":
                                    tempResultRange.Columns[1].Value2 =
                                        "一年级";
                                    tempResultRange.Columns[2].Value2 =
                                        "一年级(" + tempSplitArry[1] + ")";
                                    //tempGradeClass = ((tempYear + 6) - 1) + tempSplitArry[1];
                                    break;
                            }

                            //写入课程
                            tempResultRange.Columns[3].Value2 = tempSplitArry[2];

                            //写入教师
                            tempResultRange.Columns[4].Value2 = range.Rows[rowCount].Columns[1].Value2;

                            //写入场地
                            tempResultRange.Columns[5].Value2 = "自动";

                            //写入星期
                            tempWeek = Int32.Parse(Math.Ceiling((colCount - 1F) / 6F).ToString());
                            switch (tempWeek)
                            {
                                case 1:
                                    tempResultRange.Columns[6].Value2 = "星期一";
                                    //tempGradeClass += ".星期一";
                                    break;
                                case 2:
                                    tempResultRange.Columns[6].Value2 = "星期二";
                                    //tempGradeClass += ".星期二";
                                    break;
                                case 3:
                                    tempResultRange.Columns[6].Value2 = "星期三";
                                    //tempGradeClass += ".星期三";
                                    break;
                                case 4:
                                    tempResultRange.Columns[6].Value2 = "星期四";
                                    //tempGradeClass += ".星期四";
                                    break;
                                case 5:
                                    tempResultRange.Columns[6].Value2 = "星期五";
                                    //tempGradeClass += ".星期五";
                                    break;


                            }
                            //写入节次
                            tempResultRange.Columns[7].Value2 = range.Rows[3].Columns[colCount].Value2;

                            //tempSubject = tempSplitArry[2];
                            //这里还有一个星期几的问题没有解决，如：周长胜一行中星期一第一节和星期二第一节一样，Dictionary会报错：same Key!!!
                            //dicSchedule.Add(tempGradeClass, tempSubject);
                            /*if (tempResultRange.Columns[4].Value2.ToString() == "刘培珍")
                            {
                                MessageBox.Show("" + tempResultRange.Columns[1].Value2.ToString() +
                                                " " + tempResultRange.Columns[2].Value2.ToString() +
                                                " " + tempResultRange.Columns[3].Value2.ToString() +
                                                " " + tempResultRange.Columns[4].Value2.ToString() +
                                                " " + tempResultRange.Columns[5].Value2.ToString() +
                                                " " + tempResultRange.Columns[6].Value2.ToString() +
                                                " " + tempResultRange.Columns[7].Value2.ToString()


                                    );
                            }*/
                            resultScheduleLineNum++;
                        }
                       /* else
                        {
                           // dicSchedule.Add("[" + rowCount.ToString() + "." + colCount.ToString() + "]", "");
                        }*/
                    }
                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += range.Rows[rowCount].Columns[1].Value2.ToString() + "\n";
                        txtTextBox.ScrollToEnd();
                    }));
                    
                }

                CommonOpenFileDialog openFolderDialog = new CommonOpenFileDialog();
                openFolderDialog.IsFolderPicker = true;
                openFolderDialog.Title = "选择文件保存目录";
                //openFolderDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
                commonFileDialogResult = CommonFileDialogResult.None;
                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    commonFileDialogResult = openFolderDialog.ShowDialog();
                }));


                if (commonFileDialogResult == CommonFileDialogResult.Ok)
                {
                    string fileSavePath = openFolderDialog.FileName;
                    System.DateTime dateTime = System.DateTime.Now;
                    resultScheduleWb.SaveAs(fileSavePath + "排课结果" + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls",
                                      Excel.XlFileFormat.xlWorkbookNormal,
                                      misValue,
                                      misValue,
                                      misValue,
                                      misValue,
                                      Excel.XlSaveAsAccessMode.xlExclusive);

                    resultScheduleWb.Close(false);
                    allSchoolScheduleWb.Close(false);
                    allSchoolScheduleApp.Quit();
                }

                //for (rowCount = 4; rowCount < rCount; rowCount++)
                //{
 
                //}

                /*MessageBox.Show("rCount: " + rCount + "\n" + "cCount: " + cCount);

                for (rowCount = 1; rowCount < rCount; rowCount++)
                {
                    Excel.Range r = range.Rows[rowCount];
                }*/

                MessageBox.Show("Done!");
            }
        }

        private void allSchoolScheduleBW_DoWork(object sender, DoWorkEventArgs e)
        {
            ReadAllSchoolSchedule();
        }

        private void allSchoolScheduleBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        { 
        }

        private void allSchoolScheduleBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("allSchoolScheduleBW Done!");
        }

        //private void btnCreateExcel_Click(object sender, RoutedEventArgs e)
        //{
        //    string fileSavePath = "";

        //    targetXlApp = new Excel.Application();
        //    targetXlWb = targetXlApp.Workbooks.Add();
        //    targetXlWs = (Excel.Worksheet)targetXlWb.Worksheets.get_Item(1);
        //    //Cells的行和列都要从1号开始
        //    targetXlWs.Cells[1, 1] = "年级";
        //    targetXlWs.Cells[1, 2] = "班级";
        //    targetXlWs.Cells[1, 3] = "课程";
        //    targetXlWs.Cells[1, 4] = "教师";
        //    targetXlWs.Cells[1, 5] = "场地";
        //    targetXlWs.Cells[1, 6] = "周课时";
        //    targetXlWs.Cells[1, 7] = "每周连课次数";
        //    targetXlWs.Cells[1, 8] = "每次连课节数";
        //    targetXlWs.Cells[1, 9] = "课程性质";
        //    targetXlWs.Cells[1, 10] = "所在校区";


        //    Microsoft.Win32.SaveFileDialog saveFileDialog = new Microsoft.Win32.SaveFileDialog();
        //    saveFileDialog.Filter = "XLS文件（*.xls）|*.xls";
        //    saveFileDialog.FilterIndex = 2;
        //    saveFileDialog.RestoreDirectory = true;
        //    if (saveFileDialog.ShowDialog() == true)
        //    {
        //        fileSavePath = saveFileDialog.FileName.ToString();


        //        targetXlWb.SaveAs(fileSavePath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive);
        //        targetXlWb.Close(true);
        //        targetXlApp.Quit();
        //        MessageBox.Show("Excel file Saved!");
        //    }
        //}

        private void ProgressBar_ValueChanged_1(object sender, RoutedPropertyChangedEventArgs<double> e)
        {

        }

        private void btnOpenAllSchoolSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (allSchoolScheduleBW.IsBusy != true)
            {
                allSchoolScheduleBW.RunWorkerAsync();
            }
           /* Excel.Application allSchoolScheduleApp;
            Excel.Workbook allSchoolScheduleWb;
            Excel.Worksheet allSchoolScheduleWs;

            Microsoft.Win32.OpenFileDialog openFileDialog = new Microsoft.Win32.OpenFileDialog();
            //openFileDialog.FileName = "Document";
            openFileDialog.DefaultExt = ".xls";
            openFileDialog.Filter = "XLS文件（*.xls）|*.xls";

            Nullable<bool> result = openFileDialog.ShowDialog();
            if (result == true)
            {
                //backgroundWorker.ReportProgress(0);

                xlFilePath = openFileDialog.FileName;

                allSchoolScheduleApp = new Excel.Application();
                allSchoolScheduleWb = allSchoolScheduleApp.Workbooks.Open(xlFilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                allSchoolScheduleWs = (Excel.Worksheet)allSchoolScheduleWb.Worksheets.get_Item(1);

                Excel.Range range;

                range = allSchoolScheduleWs.UsedRange;

                int rowCount, rCount, cCount;

                rCount = range.Rows.Count;
                cCount = range.Columns.Count;

                MessageBox.Show("rCount: " + rCount + "\n" + "cCount: " + cCount);

                for(rowCount = 1; rowCount < rCount; rowCount++)
                {
                    Excel.Range r = range.Rows[rowCount];
                }


                //MessageBox.Show("Rows: " + range.Rows.Count + "\n" + "Columns: " + range.Columns.Count);

                

               // SegmentExcelIntoLineOfRange(range, range.Rows.Count, range.Columns.Count);

                //xlWb.Close(true, null, null);
                // xlApp.Quit();
            }



            */
        }


    }
}
