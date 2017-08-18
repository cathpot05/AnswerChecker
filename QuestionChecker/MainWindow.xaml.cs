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
using Microsoft.Win32;
using System.Data;
using System.Net;
using System.IO;
using System.Web.Script.Serialization;
using Newtonsoft.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Diagnostics;
using System.ComponentModel;
using System.Windows.Threading;
using System.Threading;
using System.Runtime.InteropServices;
using System.Threading.Tasks;

namespace QuestionChecker
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        DataTable dt;
        DataTable skippedDt;
        private readonly BackgroundWorker worker = new BackgroundWorker();

        const int SWP_NOZORDER = 0x4;

        [DllImport("kernel32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool AllocConsole();

       
        int countofRows = 0;
        string pathFile = "";
        bool isAutoGen = false;
        public MainWindow()
        {
            InitializeComponent();
            #region DataViews Initializations
            dt = new DataTable();
            DataColumn dc1 = new DataColumn("Sign");
            DataColumn dc2 = new DataColumn("Questions");
            DataColumn dc3 = new DataColumn("Answer");
            DataColumn dc4 = new DataColumn("Server Answer");
            DataColumn dc5 = new DataColumn("Remarks");
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            dt.Columns.Add(dc3);
            dt.Columns.Add(dc4);
            dt.Columns.Add(dc5);
            DataRow dr11 = dt.NewRow();
            //dgExcel.DataSource = dt;
            dgExcel.ItemsSource = dt.DefaultView;

            skippedDt = new DataTable();
            DataColumn x1 = new DataColumn("Questions");
            DataColumn x2 = new DataColumn("Answer");
            DataColumn x3 = new DataColumn("");
            DataColumn x4 = new DataColumn("");
            skippedDt.Columns.Add(x1);
            skippedDt.Columns.Add(x2);
            skippedDt.Columns.Add(x3);
            skippedDt.Columns.Add(x4);
            DataRow ex1 = skippedDt.NewRow();
            dgErrorRows.ItemsSource = skippedDt.DefaultView;
            #endregion
            
            worker.DoWork += new DoWorkEventHandler(bw_DoWork);
            worker.ProgressChanged += new ProgressChangedEventHandler(bw_ProgressChanged);
            worker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);
            worker.WorkerReportsProgress = true;
        }
        
        #region BGWorker

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            this.btnPost.IsEnabled = true;
        }

        private void bw_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //Console.WriteLine(e.ProgressPercentage);
            //txtCount.Content = e.ProgressPercentage.ToString() + "% complete";
        }

         private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            
            BackgroundWorker worker = (BackgroundWorker)sender;
            //Application.Current.Dispatcher.BeginInvoke(
            //DispatcherPriority.Background,
            //new Action(() =>
            //    panel_z.Visibility = Visibility.Visible
                
            //    ));
            this.Dispatcher.Invoke(() => {
                //postData();
            });

                   
            //        for (int i = 0; i < 100; ++i)
            //        {
            //            worker.ReportProgress(i);
            //            System.Threading.Thread.Sleep(100);
            //        }
        }
                
        #endregion

        private void btnOpen_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openfile = new OpenFileDialog();
            openfile.DefaultExt = ".xlsx";
            openfile.Filter = "(.xlsx)|*.xlsx";
            //openfile.ShowDialog();

            var browsefile = openfile.ShowDialog();

            if (browsefile == true)
            {
                txtFilePath.Text = openfile.FileName;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelBook = excelApp.Workbooks.Open(txtFilePath.Text.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                cboSheetName.Items.Clear();
                cboSheetName.IsEnabled = false;
                Excel.Sheets test = excelBook.Worksheets;
                for (int x = 1; x <= test.Count; x++)
                {
                    Excel.Worksheet sample = (Excel.Worksheet)excelBook.Worksheets.get_Item(x);
                    Console.WriteLine(sample.Name);
                    cboSheetName.Items.Add(sample.Name);
                }
                
                excelBook.Close(true, null, null);
                excelApp.Quit();
            }
            cboSheetName.IsEnabled = true;
            pathFile = openfile.FileName;
            
        }

        private void loadDatathroughSheets(int x)
        {
            try
            {


                int items = x + 1;
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook excelBook = excelApp.Workbooks.Open(pathFile.ToString(), 0, true, 5, "", "", true, Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                Excel.Worksheet excelSheet = (Excel.Worksheet)excelBook.Worksheets.get_Item(items);
                Excel.Range excelRange = excelSheet.UsedRange;
                string strCellData = "";
                //string douCellData;
                double douCellData;
                int rowCnt = 0;
                int colCnt = 0;
                txtCount.Content = "Row Count: " + (excelRange.Rows.Count - 1);
                countofRows = excelRange.Rows.Count - 1;
                DataTable dt = new DataTable();
                //string header1 = (excelRange.Cells[1, 1] as Excel.Range).Value2;
                //string header2 = (excelRange.Cells[1, 2] as Excel.Range).Value2;
                //if (!header1.Contains("Question"))
                //{
                //    MessageBox.Show("Make the header of Column 1 - 'QUESTIONS' ");
                //    panel_z.Visibility = Visibility.Collapsed;
                //    return;
                //}
                //if (!header2.Contains("Answer"))
                //{
                //    MessageBox.Show("Make the header of Column 2 - 'ANSWERS' ");
                //    panel_z.Visibility = Visibility.Collapsed;
                //    return;
                //}

                LoggedDateandTimeStart("excelgen");
                for (colCnt = 1; colCnt <= 2; colCnt++)
                {
                    string strColumn = "";
                    strColumn = (string)(excelRange.Cells[1, colCnt] as Excel.Range).Value2;
                    dt.Columns.Add(strColumn, typeof(string));
                }

                for (rowCnt = 2; rowCnt <= excelRange.Rows.Count; rowCnt++)
                {
                    string strData = "";
                    for (colCnt = 1; colCnt <= 2; colCnt++)
                    {
                        try
                        {
                            strCellData = (string)(excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += strCellData + "|";
                        }
                        catch (Exception ex)
                        {
                            //MessageBox.Show(ex.Message);
                            douCellData = (excelRange.Cells[rowCnt, colCnt] as Excel.Range).Value2;
                            strData += douCellData.ToString() + "|";
                        }
                    }
                    strData = strData.Remove(strData.Length - 1, 1);
                    dt.Rows.Add(strData.Split('|'));
                }

                dtGrid.ItemsSource = dt.DefaultView;

                excelBook.Close(true, null, null);
                excelApp.Quit();

                LoggedDateandTimeEnd("excelgen");
                panel_z.Visibility = Visibility.Collapsed;
                btnPost.IsEnabled = true;
            }
            catch (Exception xxx)
            {
                MessageBox.Show("Error occured: "+xxx.Message);
            }

        }

        private void clearcboSheetList()
        {
            cboSheetName.IsEnabled = false;
        }

        private bool checkisOpen(string filename)
        {
            string[] getfilename = filename.ToString().Trim().Split(new string[] { "\\" }, StringSplitOptions.None);
            string lastItem = getfilename.Last();
            bool isOpened = true;
            Excel.Application xcelapp = new Excel.Application();

            xcelapp =  (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            try
            {
                xcelapp.Workbooks.get_Item(lastItem);
            }
            catch (Exception ex)
            {
                isOpened =  false;
            }
            return isOpened;
        }
        private void LoggedDateandTimeStart(string description)
        {
            string desc = description;
            string logDir = @"C:\ErrorLog\";
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }
            var LogFile = "DataLogged" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            switch (desc)
            {
                case "posting":
                    
                    using (StreamWriter LogWriter = new StreamWriter(logDir + LogFile, true))
                    {
                        LogWriter.WriteLine(Environment.NewLine + "-------------------- " + "Posting Data Start- " + DateTime.Now.ToString() + " --------------------" + Environment.NewLine);
                        LogWriter.WriteLine("Started Posting Data: " + DateTime.Now.ToString() + Environment.NewLine);
                    }

                    break;
                case "excelgen":

                    using (StreamWriter LogWriter = new StreamWriter(logDir + LogFile, true))
                    {
                        LogWriter.WriteLine(Environment.NewLine + "-------------------- " + "Excel Start- " + DateTime.Now.ToString() + " --------------------" + Environment.NewLine);
                        LogWriter.WriteLine("Started Importing/Exporting: " + DateTime.Now.ToString() + Environment.NewLine);
                    }
                    break;
            }
            

        }


        private void LoggedDateandTimeEnd(string description)
        {
            string desc = description;
            string logDir = @"C:\ErrorLog\";
            if (!Directory.Exists(logDir))
            {
                Directory.CreateDirectory(logDir);
            }
            var LogFile = "DataLogged" + DateTime.Now.ToString("yyyyMMdd") + ".txt";

            switch (desc)
            {
                case "posting":

                    using (StreamWriter LogWriter = new StreamWriter(logDir + LogFile, true))
                    {
                        LogWriter.WriteLine("Finish Posting Data: " + DateTime.Now.ToString() + Environment.NewLine);
                        LogWriter.WriteLine("----------------------------------- " + "End" + "-----------------------------------" + Environment.NewLine);
                    }

                    break;
                case "excelgen":

                    using (StreamWriter LogWriter = new StreamWriter(logDir + LogFile, true))
                    {
                        LogWriter.WriteLine("Finish Exporting/Importing Data: " + DateTime.Now.ToString() + Environment.NewLine);
                        LogWriter.WriteLine("----------------------------------- " + "End" + "-----------------------------------" + Environment.NewLine);
                    }
                    break;
            }
        }
        
        private async Task<bool> postData()
        {
            //AllocConsole();
            
            string rawquestion = "";
            string rawanswer = "";
            string finalanswer = "";
            string userid = txtUserId.Text.Trim();
            string langUsed = txtLang.Text.Trim();
            string sign = "";
            string remarks = "";
            string skipdata = "";
            int errcount = 0;
            
            #region Looping within the DataGrid
            LoggedDateandTimeStart("posting");
            foreach (DataRowView dr in dtGrid.ItemsSource)
                {
                //consolecontrolx.WriteOutput("sample", Colors.White);
                try
                    {
                        rawquestion = dr[0].ToString();
                        rawanswer = dr[1].ToString().Trim();

                    #region HttpRequest and Response
                    var httpWebRequest = (HttpWebRequest)WebRequest.Create("http://13.90.88.110:8080/semantic/api/v1/query/");
                        httpWebRequest.ContentType = "application/json";
                        httpWebRequest.Method = "POST";
                        //Console.WriteLine(Environment.NewLine+"The Timeout time of the request before setting is : {0} milliseconds", httpWebRequest.Timeout);
                        httpWebRequest.Proxy = null;
                        httpWebRequest.ServicePoint.ConnectionLeaseTimeout = 5000;
                        httpWebRequest.ServicePoint.MaxIdleTime = 5000;
                        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
                        ServicePointManager.ServerCertificateValidationCallback = delegate { return true; };
                        using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                        {
                            string json = new JavaScriptSerializer().Serialize(new
                            {
                                query = rawquestion,
                                language = langUsed,
                                userId = userid
                            });

                            streamWriter.Write(json);
                        }
                    var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                    await Task.Delay(1000);
                    //var httpResponse = (WebResponse)httpWebRequest.GetResponseAsync();
                    #region Synchronous response
                    using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                    {
                        var result = streamReader.ReadToEnd();
                        dynamic results = JsonConvert.DeserializeObject<dynamic>(result);
                        finalanswer = results.answer;
                        if (string.Equals(finalanswer.ToString().Trim(), rawanswer.ToString().Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            sign = "o";
                            remarks = "DATA MATCH";
                            Console.WriteLine("REMARKS: Answer is MATCH!" + Environment.NewLine + " Question: " + rawquestion + Environment.NewLine + "Answer on File: " + rawanswer + Environment.NewLine + "Answer on Server: " + finalanswer + Environment.NewLine);
                        }
                        else
                        {
                            int cnt = 0;
                            string s1 = "";
                            string s2 = "";
                            string[] firstarray = finalanswer.ToString().Trim().Split(new string[] { " " }, StringSplitOptions.None);
                            for (cnt = 0; cnt < firstarray.Length; cnt++)
                            {
                                s1 += firstarray[cnt];
                            }
                            string[] secondarray = rawanswer.Split(new string[] { " " }, StringSplitOptions.None);
                            for (cnt = 0; cnt < secondarray.Length; cnt++)
                            {
                                s2 += secondarray[cnt];
                            }
                            //Console.WriteLine("string1: " + s1 + Environment.NewLine + "string2: " + s2);

                            if (string.Equals(s1.ToString().Trim(), s2.ToString().Trim(), StringComparison.OrdinalIgnoreCase))
                            {
                                sign = "o";
                                remarks = "DATA MATCH";
                                Console.WriteLine("REMARKS: Answer is MATCH!" + Environment.NewLine + " Question: " + rawquestion + Environment.NewLine + "Answer on File: " + rawanswer + Environment.NewLine + "Answer on Server: " + finalanswer + Environment.NewLine);
                            }
                            else
                            {
                                sign = "x";
                                remarks = "DON'T MATCH";
                                Console.WriteLine("REMARKS: Answer DONT MATCH!" + Environment.NewLine + " Question: " + rawquestion + Environment.NewLine + "Answer on File: " + rawanswer + Environment.NewLine + "Answer on Server: " + finalanswer + Environment.NewLine);
                            }

                        }
                        streamReader.Close();
                        httpResponse.Close();

                    }

                    #endregion
                    #endregion

                    #region Filling up dgExcelView
                    DataRow dr1 = dt.NewRow();
                        dr1[0] = sign.ToString();
                        dr1[1] = rawquestion.ToString();
                        dr1[2] = rawanswer.ToString();
                        dr1[3] = finalanswer.ToString();
                        dr1[4] = remarks.ToString();
                        dt.Rows.Add(dr1);
                        dgExcel.ItemsSource = dt.DefaultView;
                    #endregion
                    countofRows = countofRows - 1;
                    txtCount.Content = "Row Count: "+ countofRows;
                }
                    catch (Exception ex)
                    {
                        #region Logging Exceptions
                        skipdata = @"See also skipped rows due to found errors at C:\ErrorLog\Skipped Data.xls";
                        string ErrorDirectory = @"C:\ErrorLog\";
                        if (!Directory.Exists(ErrorDirectory))
                        {
                            Directory.CreateDirectory(ErrorDirectory);
                        }
                        var ErrorFile = "ErrorLog_" + DateTime.Now.ToString("yyyyMMdd") + ".txt";
                        var ErrorTrace = new StackTrace(ex, true);
                        int ErrorLine = ErrorTrace.GetFrame(ErrorTrace.FrameCount - 1).GetFileLineNumber();
                        using (StreamWriter ErrorWriter = new StreamWriter(ErrorDirectory + ErrorFile, true))
                        {
                            ErrorWriter.WriteLine(Environment.NewLine + "-------------------- " + "Exception Log Date - " + DateTime.Now.ToString() + " --------------------" + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Date: " + DateTime.Now.ToString() + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Message: " + ex.Message + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Type: " + ex.GetType() + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Line: " + ErrorLine + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Source: " + ex.Source + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Target: " + ex.TargetSite + Environment.NewLine);
                            ErrorWriter.WriteLine("Exception Details: " + Environment.NewLine + ErrorTrace + Environment.NewLine);
                            ErrorWriter.WriteLine("----------------------------------- " + "Exception Log End" + "-----------------------------------" + Environment.NewLine);
                        }
                        #endregion

                        #region Skipped Rows
                        errcount += 1;
                        DataRow drskipped = skippedDt.NewRow();
                        drskipped[0] = rawquestion.ToString();
                        drskipped[1] = rawanswer.ToString();
                        drskipped[2] = "x";
                        drskipped[3] = "" + errcount;   
                        skippedDt.Rows.Add(drskipped);
                        dgErrorRows.ItemsSource = skippedDt.DefaultView;
                    #endregion
                    }
                }//end of foreach
                #endregion


                LoggedDateandTimeEnd("posting");
                btnExport.IsEnabled = true;
                panel_z.Visibility = Visibility.Collapsed;
                
            if(chkautoGen.IsChecked == true)
            {
                btnExport_Click(new object(), new RoutedEventArgs());
                MessageBox.Show("Done getting the result. \n " + skipdata);
            }
            else
            {
                MessageBox.Show("Done getting the result. Click export now. \n " + skipdata);
            }
                panel_z.Visibility = Visibility.Collapsed;

                #region Exporting Rows in Catch Block
                if (dgErrorRows.Items.Count != 0)
                {
                    loggedRowsException();
                }

                 #endregion
            return true;
        }

        private void loggedRowsException()
        {
            if (checkisOpen(@"C:\ErrorLog\Skipped Data.xls"))
            {
                MessageBox.Show("Please close the Skipped Data excel file");
                return;
            }

            try
            {
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "Count";
                xlWorkSheet.Cells[1, 2] = "Note";
                xlWorkSheet.Cells[1, 3] = "Question";
                xlWorkSheet.Cells[1, 4] = "Answer";
                int x = 2;
                foreach (DataRowView dr in dgErrorRows.ItemsSource)
                {
                    xlWorkSheet.Cells[x, 1] = dr[0].ToString().Trim();
                    xlWorkSheet.Cells[x, 2] = dr[1].ToString().Trim();
                    xlWorkSheet.Cells[x, 3] = dr[2].ToString().Trim();
                    xlWorkSheet.Cells[x, 4] = dr[3].ToString().Trim();
                    x = x + 1;
                }
                xlWorkBook.SaveAs(@"C:\ErrorLog\SkippedData.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);

            }
        }

        private void btnExport_Click(object sender, RoutedEventArgs e)
        {
            string[] getfilename = txtFilePath.Text.ToString().Trim().Split(new string[] { "\\" }, StringSplitOptions.None);
            string lastItem = getfilename.Last();
            if (checkisOpen(txtFilePath.Text.Trim()))
            {
                MessageBox.Show("Please close the "+ lastItem + " file");
                return;
            }

            panel_z.Visibility = Visibility.Visible;
            //Excel.Range chartRange;
            try
            {
                LoggedDateandTimeStart("excelgen");
                Excel.Application xlApp = new Excel.Application();

                if (xlApp == null)
                {
                    MessageBox.Show("Excel is not properly installed!!");
                    return;
                }

                xlApp.DisplayAlerts = false;
                string filePath = @"" + txtFilePath.Text;
                Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                Excel.Sheets worksheets = xlWorkBook.Worksheets;
                var xlNewSheet = (Excel.Worksheet)worksheets.Add(worksheets[1], Type.Missing, Type.Missing, Type.Missing);
                int sheetcnt = 0;
                foreach(Excel.Worksheet samp in xlWorkBook.Sheets)
                {
                    if(samp.Name == cboSheetName.SelectedValue + "_RESULT")
                    {
                        sheetcnt = sheetcnt + 1;
                    }
                }
                if (sheetcnt > 0)
                {
                    xlNewSheet.Name = cboSheetName.SelectedValue + "_RESULT" + sheetcnt;
                }
                else
                {
                    xlNewSheet.Name = cboSheetName.SelectedValue + "_RESULT";
                }
                xlNewSheet.Cells[1, 1] = "Question";
                xlNewSheet.Cells[1, 2] = "Answer";
                xlNewSheet.Cells[1, 3] = "Sever Answer";
                xlNewSheet.Cells[1, 4] = "Remarks";
                xlNewSheet.Cells[1, 5] = "Note";
                int x = 2;
                foreach (DataRowView dr in dgExcel.ItemsSource)
                {
                    xlNewSheet.Cells[x, 1] = dr[1].ToString().Trim();
                    xlNewSheet.Cells[x, 2] = dr[2].ToString().Trim();
                    xlNewSheet.Cells[x, 3] = dr[3].ToString().Trim();
                    xlNewSheet.Cells[x, 4] = dr[4].ToString().Trim();
                    xlNewSheet.Cells[x, 5] = dr[0].ToString().Trim();
                    x = x + 1;
                }


                xlNewSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                xlNewSheet.Select();

                xlWorkBook.Save();
                xlWorkBook.Close();

                releaseObject(xlNewSheet);
                releaseObject(worksheets);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);

                MessageBox.Show("New Worksheet Created!");
                panel_z.Visibility = Visibility.Collapsed;

                LoggedDateandTimeEnd("excelgen");
                //System.Diagnostics.Process.Start(Application.ResourceAssembly.Location);
                //Application.Current.Shutdown();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                
            }


        }


        private void releaseObject(object obj)
        {
            try
            {
                Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
        
        private void btnPost_PreviewMouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (txtUserId.Text == String.Empty)
            {
                MessageBox.Show("User id cannot be empty");
                return;
            }
            if(txtLang.Text == String.Empty)
            {
                MessageBox.Show("Language cannot be empty");
                return;
            }
            //dgErrorRows.ItemsSource = null;
            //dgExcel.ItemsSource = null;
            dt.Rows.Clear();
            skippedDt.Clear();
            panel_z.Visibility = Visibility.Visible;
        }

        private async void btnPost_Click(object sender, RoutedEventArgs e)
        {
            btnPost.IsEnabled = false;
            bool x = await postData();
            if (x)
            {
                panel_z.Visibility = Visibility.Collapsed;
                txtCount.Content = "Row Count "+ 0;
            }
        }

        private void cboSheetName_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if(cboSheetName.Items.Count > 0)
            {
                Console.WriteLine(cboSheetName.SelectedIndex);
                loadDatathroughSheets(cboSheetName.SelectedIndex);
            }
            
        }

        /*
        private async Task postTrigger()
        {
            panel_z.Visibility = Visibility.Visible;
            bool x = await postData();
            if (x)
            {
                panel_z.Visibility = Visibility.Collapsed;
            }
        }*/


    }
}
