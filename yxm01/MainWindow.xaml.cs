using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
using System.Collections;
using System.IO;

namespace yxm01
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private Excel.Application xlAPP;
        private Excel.Workbook xlWb;
        private Excel.Worksheet xlWs;

        private Excel.Application app = new Excel.Application();
        private Excel.Workbook[] book = new Excel.Workbook[9];
        private Excel.Worksheet[] sheet = new Excel.Worksheet[9];

        public MainWindow()
        {
            InitializeComponent();

        }

        public void WriteSourceFile()
        {
            xlAPP = new Excel.Application();
            xlWb = xlAPP.Workbooks.Open("D:\\初一学籍\\钢实2013片区学生.XLS", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWs = (Excel.Worksheet)xlWb.Worksheets.get_Item(2);

            Excel.Range xRange = xlWs.UsedRange;
           // MessageBox.Show(xRange.Rows.Count.ToString());

            /*ArrayList xNameAL = new ArrayList();

            for (int i = 3; i <= xRange.Rows.Count; i++)
            {
                xNameAL.Add(xRange.Rows[i].Columns[3].Value2.ToString());
            }

            MessageBox.Show("AL[0]: " + xNameAL[0]);
            MessageBox.Show("AL[xNameAL.Count]: " + xNameAL[xNameAL.Count-1]);*/

            FileStream fs = new FileStream("D:\\初一学籍\\钢实2013片区学生.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);

            for (int i = 3; i <= xRange.Rows.Count; i++)
            {
                //xNameAL.Add(xRange.Rows[i].Columns[3].Value2.ToString());

                sw.Write(xRange.Rows[i].Columns[3].Value2.ToString() + "\n");
            }

            sw.Close();
            fs.Close();

            xlWb.Close(false);
            xlAPP.Quit();

            MessageBox.Show("Done");
        }

        public string[] ReadSourceFile()
        {
            

            FileStream fs = new FileStream("D:\\初一学籍\\钢实2013片区学生.txt", FileMode.Open);
            byte[] dataArray = new byte[fs.Length];
            int numByteToRead = (int)fs.Length;
            int numBytesRead = 0;

            int bufferN = 0;

            while (numByteToRead > 0)
            {
                bufferN = fs.Read(dataArray, numBytesRead, numByteToRead);

                if (bufferN == 0)
                    break;

                numBytesRead += bufferN;
                numByteToRead -= bufferN;
            }

            string sourceStr = Encoding.UTF8.GetString(dataArray);
            string[] namesArray;
           // MessageBox.Show(sourceStr);
            namesArray = sourceStr.Split('\n');
            MessageBox.Show(namesArray[0] + "," + namesArray[namesArray.Length-2]);
            fs.Close();

            return namesArray;
        }

        public void WriteTagetFile()
        {
            FileStream fs = new FileStream("D:\\初一学籍\\1-9.txt", FileMode.Create);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            Excel.Range range;

            app = new Excel.Application();

            for(int i = 0; i < 9; i++)
            {
                book[i] = app.Workbooks.Open("D:\\初一学籍\\" + (i + 1) + ".xls", 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                sheet[i] = (Excel.Worksheet)book[i].Worksheets.get_Item(1);

                range = sheet[i].UsedRange;
                //MessageBox.Show(range.Rows.Count.ToString());
                //MessageBox.Show(range.Rows[2].Columns[2].Value2.ToString());
                for (int j = 2; j <= range.Rows.Count; j++)
                {
                    sw.Write(range.Rows[j].Columns[2].Value2.ToString() + "\n");
                }
                //MessageBox.Show(range.Rows[2].Columns[2].Value2.ToString() + "," + range.Rows[range.Rows.Count].Columns[2].Value2.ToString());



                book[i].Close(false);
            }

            sw.Close();
            fs.Close();
   
            app.Quit();

            MessageBox.Show("Done");
        }

        public string[] ReadTargetFile()
        {
            FileStream fs = new FileStream("D:\\初一学籍\\1-9.txt", FileMode.Open);
            byte[] dataArray = new byte[fs.Length];
            int numByteToRead = (int)fs.Length;
            int numBytesRead = 0;

            int bufferN = 0;

            while (numByteToRead > 0)
            {
                bufferN = fs.Read(dataArray, numBytesRead, numByteToRead);

                if (bufferN == 0)
                    break;

                numBytesRead += bufferN;
                numByteToRead -= bufferN;
            }

            string sourceStr = Encoding.UTF8.GetString(dataArray);
            string[] namesArray;
            // MessageBox.Show(sourceStr);
            namesArray = sourceStr.Split('\n');
            MessageBox.Show(namesArray[0] + "," + namesArray[namesArray.Length - 2]);
            fs.Close();
            return namesArray;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            WriteSourceFile();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            ReadSourceFile();
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            WriteTagetFile();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            ReadTargetFile();
        }

        private void Button_Click_4(object sender, RoutedEventArgs e)
        {
            string[] source = ReadSourceFile();
            string[] target = ReadTargetFile();

            ArrayList sourceAL = new ArrayList();
            ArrayList targetAL = new ArrayList();

            for (int i = 0; i < source.Length; i++)
            {
                sourceAL.Add(source[i]);
            }
            //MessageBox.Show(sourceAL[sourceAL.Count - 2].ToString());

            for (int i = 0; i < target.Length; i++)
            {
                targetAL.Add(target[i]);
            }
            //MessageBox.Show(targetAL[targetAL.Count - 2].ToString());

            int countMatch = 0;

            for(int j = 0; j < sourceAL.Count; j++)
            {
                if (targetAL.Contains(sourceAL[j]))
                {
                    //MessageBox.Show("有");
                    countMatch++;
                }
                else
                {
                    // MessageBox.Show("没得");
                }

                //MessageBox.Show("正在搜索：" + countMatch + " " + sourceAL[j].ToString());
            }

            
            MessageBox.Show("总共有:" + countMatch + "人");


            
        }
    }
}
