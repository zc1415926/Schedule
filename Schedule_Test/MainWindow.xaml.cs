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
        private BackgroundWorker btnXRhjxxConverterBW;

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

        private Excel.Application xRhjxxApp;
        private Excel.Workbook xRhjxxWb;
        private Excel.Worksheet xRhjxxWs;

        private Excel.Workbook resultScheduleWb;
        private Excel.Worksheet resultScheduleWs;
        private int resultScheduleLineNum;

        private Excel.Workbook teachingPlanWb;
        private Excel.Worksheet teachingPlanWs;


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

            btnXRhjxxConverterBW = new BackgroundWorker();
            btnXRhjxxConverterBW.WorkerReportsProgress = true;
            btnXRhjxxConverterBW.WorkerSupportsCancellation = true;
            btnXRhjxxConverterBW.DoWork += new DoWorkEventHandler(btnXRhjxxConverterBW_DoWork);
            btnXRhjxxConverterBW.ProgressChanged += new ProgressChangedEventHandler(btnXRhjxxConverterBW_ProgressChanged);
            btnXRhjxxConverterBW.RunWorkerCompleted += new RunWorkerCompletedEventHandler(btnXRhjxxConverterBW_RunWorkerCompleted);

            misValue = System.Reflection.Missing.Value;
        }

#region 生成教学计划按钮相关代码

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

            if (commonFileDialogResult == CommonFileDialogResult.Ok) 
            {
                backgroundWorker.ReportProgress(0);

                xlFilePath = openFileDialog.FileName;

                xlApp = new Excel.Application();
                xlWb = xlApp.Workbooks.Open(xlFilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWs = (Excel.Worksheet)xlWb.Worksheets.get_Item(1);

                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       "开始对文件 " + xlFilePath + "进行处理\n" + 
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));

                Excel.Range range;

                range = xlWs.UsedRange;

                SegmentExcelIntoLineOfRange(range, range.Rows.Count/*, range.Columns.Count*/);
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

        private void SegmentExcelIntoLineOfRange(Excel.Range range, int rCount/*, int cCount*/)
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

                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "正在对 " + teacherName + " 的信息进行处理\n";
                        txtTextBox.ScrollToEnd();
                    }));
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
            string tempTeacher = "";
            for (rowCount = 1; rowCount < rCount + 1; rowCount++)
            {
                r = range.Rows[rowCount];
                classList = r.Columns[5].Value2.ToString();
                classArray = classList.Split(stringSeparators, StringSplitOptions.RemoveEmptyEntries);
 
                int classArrayLength = classArray.Length;
                if (classArrayLength > 1)
                {
                    if (tempTeacher != r.Columns[1].Value2.ToString())
                    {
                        tempTeacher = r.Columns[1].Value2.ToString();

                        App.Current.Dispatcher.Invoke(new Action(() =>
                        {
                            txtTextBox.Text += "正在对 " + tempTeacher + " 的任课班级信息进行处理\n";
                            txtTextBox.ScrollToEnd();
                        }));
                    }
                    

                    //周总课时
                    timesPerWeek = (int)((int)(r.Columns[4].Value2 / classArrayLength));

                    for (int arrayI = 0; arrayI < classArrayLength; arrayI++)
                    {
                        if (arrayI == 0)
                        {
                            r.Columns[5].Value2 = classArray[0];
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
                        }
                    }        
                } 
            }

            range = xlWs.UsedRange;

            CreateTargetExcel(range, range.Rows.Count);
        }

        private void CreateTargetExcel(Excel.Range range, int rCount)
        {
            App.Current.Dispatcher.Invoke(new Action(() =>
            {
                txtTextBox.Text += "====================\n" +
                                   "正在生成 教学计划 表\n" +
                                   "====================\n";
                txtTextBox.ScrollToEnd();
            }));

            string fileSavePath = "";

            targetXlApp = new Excel.Application();
            //targetXlApp.Visible = false;
            targetXlWb = targetXlApp.Workbooks.Add();
            //resultScheduleWb = targetXlApp.Workbooks.Add();

            targetXlWs = (Excel.Worksheet)targetXlWb.Worksheets.get_Item(1);
            //resultScheduleWs = (Excel.Worksheet)resultScheduleWb.Worksheets.get_Item(1);

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
            string tempTechaer = "";

            //令rowCount的初值为2跳过静养一行,"rCount"而不是rCount+1是要去掉表尾的合计一行
            for (rowCount = 2; rowCount < rCount; rowCount++)
            {
                r = range.Rows[rowCount];
                
                //年级
                grade = yearGrad - Int32.Parse(r.Columns[2].Value2.ToString());

                switch (grade)
                {
                    case 6:
                        gradeString = "六年级";
                        targetXlWs.Cells[rowCount, 1] = "六年级";
                        break;

                    case 5:
                        gradeString = "五年级";
                        targetXlWs.Cells[rowCount, 1] = "五年级";
                        break;

                    case 4:
                        gradeString = "四年级";
                        targetXlWs.Cells[rowCount, 1] = "四年级";
                        break;

                    case 3:
                        gradeString = "三年级";
                        targetXlWs.Cells[rowCount, 1] = "三年级";
                        break;

                    case 2:
                        gradeString = "二年级";
                        targetXlWs.Cells[rowCount, 1] = "二年级";
                        break;

                    case 1:
                        gradeString = "一年级";
                        targetXlWs.Cells[rowCount, 1] = "一年级";
                        break;
                }

                //班级
                targetXlWs.Cells[rowCount, 2] = gradeString + "(" + r.Columns[5].Value2.ToString() + ")";
                
                //课程
                targetXlWs.Cells[rowCount, 3] = r.Columns[3].Value2.ToString();

                //教师
                if(tempTechaer != r.Columns[1].Value2.ToString())
                {
                    tempTechaer = r.Columns[1].Value2.ToString();
                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "正在生成 " + tempTechaer + " 的教学计划\n";
                        txtTextBox.ScrollToEnd();
                    }));
                }
                
                targetXlWs.Cells[rowCount, 4] = tempTechaer;

                //周课时
                targetXlWs.Cells[rowCount, 6] = r.Columns[4].Value2.ToString();

                //所在校区
                targetXlWs.Cells[rowCount, 10] = "人和街小学";
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

                targetXlWb.SaveAs(fileSavePath+"\\教学计划"+dateTime.Year+dateTime.Month+dateTime.Day+dateTime.Hour+dateTime.Minute+".xls",
                                  Excel.XlFileFormat.xlWorkbookNormal,
                                  misValue,
                                  misValue,
                                  misValue,
                                  misValue,
                                  Excel.XlSaveAsAccessMode.xlExclusive);

                targetXlWb.Close(false);
                //targetXlApp.UserControl = true;
                targetXlApp.Quit();

                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       "成功生成文件" + fileSavePath + "\\教学计划" + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls\n" +
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));
            }
            else if (commonFileDialogResult == CommonFileDialogResult.Cancel)
            {
                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       "教学计划 未保存\n" +
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));
            }

            //把本程序关闭以后，EXCEL进程就会自动结束
            xlWb.Close(false);
            //xlApp.UserControl = true;
            xlApp.Quit();

            backgroundWorker.ReportProgress(100);
            MessageBox.Show("Mission Acomplished!!!");
        }

#endregion

#region "打开全校总课表->生成排课结果"按钮的相关代码
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
                allSchoolScheduleBW.ReportProgress(0);

                xlFilePath = openFileDialog.FileName;

                allSchoolScheduleApp = new Excel.Application();
                //allSchoolScheduleApp.Visible = false;
                allSchoolScheduleWb = allSchoolScheduleApp.Workbooks.Open(xlFilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                allSchoolScheduleWs = (Excel.Worksheet)allSchoolScheduleWb.Worksheets.get_Item(1);

                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       "开始对文件 " + xlFilePath + "进行处理\n" +
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));

                Excel.Range range;

                range = allSchoolScheduleWs.UsedRange;

                int rowCount, rCount, cCount;

                rCount = range.Rows.Count;
                cCount = range.Columns.Count;
               /* MessageBox.Show("rCount: " + rCount +
                                "cCount: " + cCount);*/

                //这里先试试一个人的
                string cellString = "";
                string[] tempSplitArry;
                int tempWeek;
                //int tempYear = System.DateTime.Now.Year;
                Excel.Range tempResultRange;

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

                for (rowCount = 4; rowCount < rCount + 1; rowCount++ )
                {
                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "正在生成 " + range.Rows[rowCount].Columns[1].Value2.ToString() + " 的排课结果\n";
                        txtTextBox.ScrollToEnd();
                    }));

                    //colCount = 2是为了避过姓名一列,colCount是从1号开始的cCount要加一
                    for (int colCount = 2; colCount < cCount + 1; colCount++)
                    {

                        if (range.Rows[rowCount].Columns[colCount].Value2 != null)
                        {
                            cellString = range.Rows[rowCount].Columns[colCount].Value2.ToString();
                            cellString = cellString.Replace("\n", ".");
                            tempSplitArry = cellString.Split('.');

                            if (tempSplitArry.Length != 3)
                            {
                                MessageBox.Show("这个表不对头哦，只有班级没有学科！");
                                
                            }
                            //传地址了吗？忘了tempResultRange是干什么用的了
                            tempResultRange = resultScheduleWs.Rows[resultScheduleLineNum];

                            //写入“排课结果的年级、班级两列
                            switch (tempSplitArry[0])
                            {
                                case "小六":
                                    tempResultRange.Columns[1].Value2 = "六年级";
                                    tempResultRange.Columns[2].Value2 = "六年级(" + tempSplitArry[1] + ")";
                                    break;

                                case "小五":
                                    tempResultRange.Columns[1].Value2 = "五年级";
                                    tempResultRange.Columns[2].Value2 = "五年级(" + tempSplitArry[1] + ")";
                                    break;

                                case "小四":
                                    tempResultRange.Columns[1].Value2 = "四年级";
                                    tempResultRange.Columns[2].Value2 = "四年级(" + tempSplitArry[1] + ")";
                                    break;

                                case "小三":
                                    tempResultRange.Columns[1].Value2 = "三年级";
                                    tempResultRange.Columns[2].Value2 = "三年级(" + tempSplitArry[1] + ")";
                                    break;

                                case "小二":
                                    tempResultRange.Columns[1].Value2 = "二年级";
                                    tempResultRange.Columns[2].Value2 = "二年级(" + tempSplitArry[1] + ")";
                                    break;

                                case "小一":
                                    tempResultRange.Columns[1].Value2 = "一年级";
                                    tempResultRange.Columns[2].Value2 = "一年级(" + tempSplitArry[1] + ")";
                                    break;
                            }

                            //写入课程
                            tempResultRange.Columns[3].Value2 = tempSplitArry[2];

                            //写入教师
                            tempResultRange.Columns[4].Value2 = range.Rows[rowCount].Columns[1].Value2;

                            //写入场地
                            tempResultRange.Columns[5].Value2 = "自动";

                            //写入星期
                            tempWeek = int.Parse(Math.Ceiling((colCount - 1F) / 6F).ToString());
                            switch (tempWeek)
                            {
                                case 1:
                                    tempResultRange.Columns[6].Value2 = "星期一";
                                    break;

                                case 2:
                                    tempResultRange.Columns[6].Value2 = "星期二";
                                    break;

                                case 3:
                                    tempResultRange.Columns[6].Value2 = "星期三";
                                    break;

                                case 4:
                                    tempResultRange.Columns[6].Value2 = "星期四";
                                    break;

                                case 5:
                                    tempResultRange.Columns[6].Value2 = "星期五";
                                    break;
                            }
                            //写入节次
                            tempResultRange.Columns[7].Value2 = range.Rows[3].Columns[colCount].Value2;

                            resultScheduleLineNum++;
                        }
                    }

                    
                    
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
                    allSchoolScheduleBW.ReportProgress(100);

                    string fileSavePath = openFolderDialog.FileName;
                    System.DateTime dateTime = System.DateTime.Now;
                    resultScheduleWb.SaveAs(fileSavePath + "\\排课结果" + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls",
                                      Excel.XlFileFormat.xlWorkbookNormal,
                                      misValue,
                                      misValue,
                                      misValue,
                                      misValue,
                                      Excel.XlSaveAsAccessMode.xlExclusive);

                    resultScheduleWb.Close(false);
                    allSchoolScheduleWb.Close(false);
                    allSchoolScheduleApp.Quit();

                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "====================\n" +
                                           "成功生成文件" + fileSavePath + "\\排课结果" + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls" +
                                           "====================\n";
                        txtTextBox.ScrollToEnd();
                    }));
                }
                else if (commonFileDialogResult == CommonFileDialogResult.Cancel)
                {
                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "====================\n" +
                                           "排课结果 未保存\n" +
                                           "====================\n";
                        txtTextBox.ScrollToEnd();
                    }));
                }

                MessageBox.Show("Mission Accomplished!!!");
            }
        }

        private void allSchoolScheduleBW_DoWork(object sender, DoWorkEventArgs e)
        {
            ReadAllSchoolSchedule();
        }

        private void allSchoolScheduleBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                //probarConvertProgress.IsIndeterminate = true;
            }
            else if (e.ProgressPercentage == 100)
            {
              //  probarConvertProgress.IsIndeterminate = false;
            }
        }

        private void allSchoolScheduleBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("allSchoolScheduleBW Done!");
        }

        private void btnOpenAllSchoolSchedule_Click(object sender, RoutedEventArgs e)
        {
            if (allSchoolScheduleBW.IsBusy != true)
            {
                allSchoolScheduleBW.RunWorkerAsync();
            }
        }

#endregion

        #region “重庆天地课表转换”按钮相关代码

        private void btnXRhjxxConverter_Click(object sender, RoutedEventArgs e)
        {
            if (btnXRhjxxConverterBW.IsBusy != true)
            {
                btnXRhjxxConverterBW.RunWorkerAsync();
            }
        }

        private void btnXRhjxxConverterBW_DoWork(object sender, DoWorkEventArgs e)
        {
            XRhjxxConverter();
        }

        private void btnXRhjxxConverterBW_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 0)
            {
                probarConvertProgress.IsIndeterminate = true;
            }
            else if (e.ProgressPercentage == 100)
            {
                probarConvertProgress.IsIndeterminate = false;
            }
        }

        private void btnXRhjxxConverterBW_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show("重庆天地课表转换完成!");
        }

        private void XRhjxxConverter()
        {
            xlFilePath = OpenExcelFile();
            
            if("" != xlFilePath)
            {
                //btnXRhjxxConverterBW.ReportProgress(0);

                xRhjxxApp = new Excel.Application();
                //allSchoolScheduleApp.Visible = false;
                xRhjxxWb = xRhjxxApp.Workbooks.Open(xlFilePath, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xRhjxxWs = (Excel.Worksheet)xRhjxxWb.Worksheets.get_Item(1);

                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       "开始对文件 " + xlFilePath + "进行处理\n" +
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));

                Excel.Range range;

                range = xRhjxxWs.UsedRange;

                int rowCount, rCount, cCount;
                rCount = range.Rows.Count;
                cCount = range.Columns.Count;

                string strClassAndGrade = "";
                string[] strClassGradeSplitArry;

                string[] strSubjectTeacherSplitArray;

                int intClassWeek;
                /* MessageBox.Show("rCount: " + rCount + "cCount: " + cCount);*/

                //结果表，将排课结果汇总，最后输出该表即可
                resultScheduleWb = xRhjxxApp.Workbooks.Add();
                resultScheduleWs = (Excel.Worksheet)resultScheduleWb.Worksheets.get_Item(1);

                teachingPlanWb = xRhjxxApp.Workbooks.Add();
                teachingPlanWs = (Excel.Worksheet)teachingPlanWb.Worksheets.get_Item(1);

                //临时存储表格一行的内容
                Excel.Range tempResultRange;
                Excel.Range rangeClassesNum;
                rangeClassesNum = range.Rows[2];

                /* 生成“排课结果” */
                resultScheduleWs.Cells[1, 1] = "年级";
                resultScheduleWs.Cells[1, 2] = "班级";
                resultScheduleWs.Cells[1, 3] = "课程";
                resultScheduleWs.Cells[1, 4] = "教师";
                resultScheduleWs.Cells[1, 5] = "场地";
                resultScheduleWs.Cells[1, 6] = "星期";
                resultScheduleWs.Cells[1, 7] = "节次";
                resultScheduleLineNum = 2;
                Excel.Range currentResultScheduleRow;

                string strClass = "";

                //跳过两行表头
                for (rowCount = 3; rowCount < rCount + 1; rowCount++)
                {
                    //tempResultRange=全校班级总课表.xls中的第几行
                    tempResultRange = range.Rows[rowCount];

                    strClassAndGrade = tempResultRange.Columns[1].Value2.ToString();

                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "正在生成 " + strClassAndGrade + " 的排课结果\n";
                        txtTextBox.ScrollToEnd();
                    }));

                    strClassAndGrade = strClassAndGrade.Replace("级", "级.");

                    strClassGradeSplitArry = strClassAndGrade.Split('.');
                    strClassGradeSplitArry[1] = strClassGradeSplitArry[1].Replace("班", "");

                    //编号是从1开始的，所以后边要+1
                    for (int colCount = 2; colCount < cCount + 1; colCount++)
                    {
                        strSubjectTeacherSplitArray = tempResultRange.Columns[colCount].Value2.ToString().Split('\n');
                        strClass = strClassGradeSplitArry[0];
                        //Range赋值传的是引用
                        currentResultScheduleRow = resultScheduleWs.Rows[resultScheduleLineNum];
                        //年级
                        currentResultScheduleRow.Columns[1].Value2 = strClass;
                        //班级
                        currentResultScheduleRow.Columns[2].Value2 = strClass + "(" + strClassGradeSplitArry[1] + ")";
                        //课程
                        currentResultScheduleRow.Columns[3].Value2 = strSubjectTeacherSplitArray[0];
                        //教师
                        currentResultScheduleRow.Columns[4].Value2 = strSubjectTeacherSplitArray[1];
                        //场地
                        currentResultScheduleRow.Columns[5].Value2 = "自动";

                        //计算星期几
                        intClassWeek = int.Parse(Math.Ceiling((colCount - 1F) / 6F).ToString());

                        switch (intClassWeek)
                            {
                                case 1:
                                    currentResultScheduleRow.Columns[6].Value2 = "星期一";
                                    break;

                                case 2:
                                    currentResultScheduleRow.Columns[6].Value2 = "星期二";
                                    break;

                                case 3:
                                    currentResultScheduleRow.Columns[6].Value2 = "星期三";
                                    break;

                                case 4:
                                    currentResultScheduleRow.Columns[6].Value2 = "星期四";
                                    break;

                                case 5:
                                    currentResultScheduleRow.Columns[6].Value2 = "星期五";
                                    break;
                            }

                        
                        currentResultScheduleRow.Columns[7].Value2 = rangeClassesNum.Columns[colCount].Value2.ToString();

                        resultScheduleLineNum++;
                    }

                   //反正只有一个源表，干脆一块儿把“教学计划”表给生成了！！！

                    
                }

                SaveExcelFile("重庆天地校区-排课结果", resultScheduleWb);

                //教学计划
                teachingPlanWs.Cells[1, 1] = "年级";
                teachingPlanWs.Cells[1, 2] = "班级";
                teachingPlanWs.Cells[1, 3] = "课程";
                teachingPlanWs.Cells[1, 4] = "教师";
                teachingPlanWs.Cells[1, 5] = "场地";
                teachingPlanWs.Cells[1, 6] = "周课时";
                teachingPlanWs.Cells[1, 7] = "每周连课次数";
                teachingPlanWs.Cells[1, 8] = "每次连课节数";
                teachingPlanWs.Cells[1, 9] = "课程性质";
                teachingPlanWs.Cells[1, 10] = "所在校区";
                Excel.Range teachingPlanRow;

                Dictionary<string, byte> dicClassGradeSubjectTeacherToTimes = new Dictionary<string, byte>();

                range = resultScheduleWs.UsedRange;
                rCount = range.Rows.Count;
                cCount = range.Columns.Count;

                string strTeacher = "";
                string strSubject = "";

                string strCompareClassAndGrade = "";

                for (rowCount = 2; rowCount < rCount + 1; rowCount++)
                {
                    teachingPlanRow = range.Rows[rowCount];

                    strClassAndGrade = teachingPlanRow.Columns[2].Value2.ToString();
                    strSubject = teachingPlanRow.Columns[3].Value2.ToString();
                    strTeacher = teachingPlanRow.Columns[4].Value2.ToString();

                    if (strCompareClassAndGrade != strClassAndGrade)
                    {
                        strCompareClassAndGrade = strClassAndGrade;

                        App.Current.Dispatcher.Invoke(new Action(() =>
                        {
                            txtTextBox.Text += "正在处理 " + strCompareClassAndGrade + " 的教学计划\n";
                            txtTextBox.ScrollToEnd();
                        }));
                    }
                    

                    if (!dicClassGradeSubjectTeacherToTimes.ContainsKey(strClassAndGrade + "." + strSubject + "." + strTeacher))
                    {
                        dicClassGradeSubjectTeacherToTimes.Add(strClassAndGrade + "." + strSubject + "." + strTeacher, 1);
                    }
                    else
                    {
                        dicClassGradeSubjectTeacherToTimes[strClassAndGrade + "." + strSubject + "." + strTeacher]++;
                    }

                }

                //string dicKey = "";
                string[] strClassGradeSubjectTeacherArray;
               
                rowCount = 2;
                
                foreach (KeyValuePair<string, byte> dic in dicClassGradeSubjectTeacherToTimes)
                {
                   // dicKey = dic.Key;
                    strClassGradeSubjectTeacherArray = dic.Key.Split('.');
                    strClassAndGrade = strClassGradeSubjectTeacherArray[0];
                    strSubject = strClassGradeSubjectTeacherArray[1];
                    //年级
                    teachingPlanWs.Rows[rowCount].Columns[1].Value2 = strClassAndGrade.Substring(0, 3);//dicKey.Substring(0, 3);
                    //班级
                    teachingPlanWs.Rows[rowCount].Columns[2].Value2 = strClassAndGrade;
                    //课程
                    teachingPlanWs.Rows[rowCount].Columns[3].Value2 = strSubject;
                    //教师
                    teachingPlanWs.Rows[rowCount].Columns[4].Value2 = strClassGradeSubjectTeacherArray[2];
                    //周课时
                    teachingPlanWs.Rows[rowCount].Columns[6].Value2 = dic.Value.ToString();
                    //所在校区
                    teachingPlanWs.Rows[rowCount].Columns[10].Value2 = "重庆天地";

                    App.Current.Dispatcher.Invoke(new Action(() =>
                    {
                        txtTextBox.Text += "正在生成教学计划\n        " + strClassAndGrade + "班 " + strSubject + "\n";
                        txtTextBox.ScrollToEnd();
                    }));

                    rowCount++;
                }


                SaveExcelFile("重庆天地校区-教学计划", teachingPlanWb);


                resultScheduleWb.Close(false);
                teachingPlanWb.Close(false);
                xRhjxxWb.Close(false);
                xRhjxxApp.Quit();
               // 干脆做成存储成mdb格式
            }
        }

        private string OpenExcelFile()
        {
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
                return openFileDialog.FileName;
            }
            else
            {
                return "";
            }

            
        }

        private void SaveExcelFile(string fileCoreName, Excel.Workbook workbook)
        {
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
                // btnXRhjxxConverterBW.ReportProgress(100);

                string fileSavePath = openFolderDialog.FileName;
                System.DateTime dateTime = System.DateTime.Now;
                workbook.SaveAs(fileSavePath + "\\" + fileCoreName + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls",
                                  Excel.XlFileFormat.xlWorkbookNormal,
                                  misValue,
                                  misValue,
                                  misValue,
                                  misValue,
                                  Excel.XlSaveAsAccessMode.xlExclusive);

                ///                    resultScheduleWb.Close(false);
                ///                   xRhjxxWb.Close(false);
                ///                    xRhjxxApp.Quit();

                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       "成功生成文件" + "\\" + fileSavePath + fileCoreName + dateTime.Year + dateTime.Month + dateTime.Day + dateTime.Hour + dateTime.Minute + ".xls" +
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));
            }
            else if (commonFileDialogResult == CommonFileDialogResult.Cancel)
            {
                App.Current.Dispatcher.Invoke(new Action(() =>
                {
                    txtTextBox.Text += "====================\n" +
                                       fileCoreName + " 未保存\n" +
                                       "====================\n";
                    txtTextBox.ScrollToEnd();
                }));
            }
        }

#endregion
    }
}
