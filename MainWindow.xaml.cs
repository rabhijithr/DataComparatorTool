using Microsoft.Win32;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.OracleClient;
using System.IO;
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
using System.Data.SqlClient;
using System.Data;
using OfficeOpenXml;
using Application = Microsoft.Office.Interop.Excel.Application;
using System.Runtime.InteropServices;
using System.Windows.Interop;
using System.Diagnostics;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Data.OleDb;    
using OfficeOpenXml.Style;

namespace DataComparator
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        [DllImport("user32.dll")]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        private const int GWL_STYLE = -16;
        private const int WS_MAXIMIZEBOX = 0x10000; //maximize button
        //private const int WS_MINIMIZEBOX = 0x20000; //uncomment this to disable minimize button


        OracleConnection OracleConnection = new OracleConnection();
        SqlConnection SQLconnection = new SqlConnection();


        bool ConnectionFlag1, ConnectionFlag2, FileUploadFlag, SchemaCheckFlag = false;
        int ErrorCount = 0;        
        string db1,db2, CompareType, filePath, exception, T1,T2, ExceptionsOutput, OutputFile = string.Empty;
        string ConnectionString = "";
        List<int> binary = new List<int>();
        List<string> Mismatch = new List<string>();
        List<string> BlankValues = new List<string>(); 


        private IntPtr _windowHandle;
        
        public MainWindow()
        {
            InitializeComponent();
            this.SourceInitialized += MainWindow_SourceInitialized;
            System.Windows.Application.Current.MainWindow.Height = 270;
            GrpBxDBToFile.Visibility = Visibility.Hidden;
            GrpBxDBToDB.Visibility = Visibility.Hidden;
            GrpBxFileToFile.Visibility = Visibility.Hidden;
        }


        //Handles the Minimize and Maximize button
        private void MainWindow_SourceInitialized(object sender, EventArgs e)
        {
            _windowHandle = new WindowInteropHelper(this).Handle;

            //disable minimize button
            if (_windowHandle == null)
                throw new InvalidOperationException("The window has not yet been completely initialized");
            SetWindowLong(_windowHandle, GWL_STYLE, GetWindowLong(_windowHandle, GWL_STYLE) & ~WS_MAXIMIZEBOX);
        }



        private void DBToDB_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.MainWindow.Height = 700;
            GrpBxDBToFile.Visibility = Visibility.Hidden;
            GrpBxFileToFile.Visibility = Visibility.Hidden;
            GrpBxDBToDB.Header = "DB to DB Compare";
            GrpBxDBToDB.Visibility = Visibility.Visible;
            OpenResultButton.IsEnabled = false;
            ExceptionButton.IsEnabled = false;
            CompareButton.IsEnabled = false;
            CompareType = "DBToDB_Checked";
        }
        private void FileToFile_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.MainWindow.Height = 485;
            GrpBxDBToFile.Visibility = Visibility.Hidden;
            GrpBxDBToDB.Visibility = Visibility.Hidden;
            GrpBxFileToFile.Header = "FILE To FILE Compare";
            GrpBxFileToFile.Visibility = Visibility.Visible;

            CompareType = "FileToFile_Checked";

            //Comment or delete this below code when functionality is implemented
            GrpBxFileToFile_OpenResultButton.IsEnabled = false;
            GrpBxFileToFile_CompareButton.IsEnabled = false;
            GrpBxFileToFile_BrowseButton2.IsEnabled = false;
            GrpBxFileToFile_BrowseButton1.IsEnabled = false;
        }

        private void DBToFile_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.MainWindow.Height = 700;
            GrpBxDBToDB.Visibility = Visibility.Hidden;
            GrpBxFileToFile.Visibility = Visibility.Hidden;
            GrpBxDBToFile.Header = "FILE To DB Compare";
            GrpBxDBToFile.Visibility = Visibility.Visible;

            GrpBxDBToFile_OpenResultButton.IsEnabled = false;
            GrpBxDBToFile_CompareButton.IsEnabled = false;
            GrpBxDBToFile_ExceptionButton.IsEnabled = false;

            //Comment or delete this below code when functionality is implemented
            //GrpBxDBToFile_DBdropdown.IsEnabled = false;
            //GrpBxDBToFile_HostName.IsEnabled = false;
            //GrpBxDBToFile_ServiceName.IsEnabled = false;
            //GrpBxDBToFile_UserName.IsEnabled = false;
            //GrpBxDBToFile_Password.IsEnabled = false;
            //GrpBxDBToFile_Port.IsEnabled = false;
            //GrpBxDBToFile_ConnectButton.IsEnabled = false;
            //GrpBxDBToFile_BrowseButton.IsEnabled = false;
            //GrpBxDBToFile_CompareButton.IsEnabled = false;
            //GrpBxDBToFile_OpenResultButton.IsEnabled = false;
            //GrpBxDBToFile_QueryTxtBox.IsEnabled = false;

            CompareType = "DBToFile_Checked";
        }

        private void GrpBxDBToFile_DBdropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            db1 = ((ComboBoxItem)GrpBxDBToFile_DBdropdown.SelectedItem).Content.ToString();
        }
        
        private void DBDropdown_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            db1="";
            db1 = ((ComboBoxItem)DBdropdown.SelectedItem).Content.ToString();
        }
        private void DBDropdown2_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            db2 = "";
            db2 = ((ComboBoxItem)DBdropdown2.SelectedItem).Content.ToString();
        }
        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
         
        }

        public bool ConnectOracleDB(string Host, string ServiceName, string UserName, string Password, string Port)
        {
            ConnectionString = "SERVER=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + Host + ")(PORT=" + Port + "))(CONNECT_DATA=(SERVICE_NAME=" + ServiceName + ")));uid =" + UserName + "; pwd=" + Password + ";";
            try
            {
                using (OracleConnection = new OracleConnection(ConnectionString))
                {
                    OracleConnection.Open();
                    GrpBxDBToFile_CompareButton.IsEnabled = true;
                    CompareButton.IsEnabled = true;
                        return true;
                }
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                GrpBxDBToFile_CompareButton.IsEnabled = false;
                CompareButton.IsEnabled = false;
                return false;
            }
        }
        public bool ConnectSQLDB(string Host, string ServiceName, string UserName, string Password, string Port)
        {
            ConnectionString = "SERVER=(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST=" + Host + ")(PORT=" + Port + "))(CONNECT_DATA=(SERVICE_NAME=" + ServiceName + ")));uid =" + UserName + "; pwd=" + Password + ";";
            try
            {
                using (SQLconnection = new SqlConnection(ConnectionString))
                {
                    SQLconnection.Open();
                    GrpBxDBToFile_CompareButton.IsEnabled = true;
                    CompareButton.IsEnabled = true;
                    return true;
                }
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                GrpBxDBToFile_CompareButton.IsEnabled = false;
                CompareButton.IsEnabled = false;
                return false;
            }
        }
        private void CopyDBDetails_Checked(object sender, RoutedEventArgs e)
        {
            HostName2.Text = HostName.Text;
            ServiceName2.Text = ServiceName.Text;
            UserName2.Text = UserName.Text;
            Password2.Password = Password.Password;
            Port2.Text = Port.Text;
        }
        private void NullDBDetails_Checked(object sender, RoutedEventArgs e)
        {
            HostName2.Text = "";
            ServiceName2.Text = "";
            UserName2.Text = "";
            Password2.Password = "";
            Port2.Text = "";
        }
        

        //Connects to the specific type of Databse Selected
        private void ConnectButton_Click(object sender, RoutedEventArgs e)
        {
            switch (CompareType)
            {
                case "DBToDB_Checked":
                    switch (db1)
                    {
                        case "Oracle":
                            ConnectionFlag1 = ConnectOracleDB(HostName.Text, ServiceName.Text, UserName.Text, Password.Password, Port.Text);
                            break;
                        case "MySQL":
                            ConnectionFlag1 = ConnectSQLDB(HostName.Text, ServiceName.Text, UserName.Text, Password.Password, Port.Text);
                            break;
                        default:
                            exception = "Couldn't connect";
                            break;

                    }
                    switch (db2)
                    {
                        case "Oracle":
                            ConnectionFlag2 = ConnectOracleDB(HostName2.Text, ServiceName2.Text, UserName2.Text, Password2.Password, Port2.Text);
                            break;
                        case "MySQL":
                            ConnectionFlag2 = ConnectSQLDB(HostName2.Text, ServiceName2.Text, UserName2.Text, Password2.Password, Port2.Text);
                            break;
                        default:
                            exception = "Couldn't connect";
                            break;
                    }
                    if (ConnectionFlag1 && ConnectionFlag2)
                    {
                        ConnectionStatus.Content = "Connection Established";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.Green);
                        ConnChk.Content = "Pass";
                        ConnChk.Foreground = new SolidColorBrush(Colors.PaleGreen);
                    }
                    else if (ConnectionFlag1 && !ConnectionFlag2)
                    {
                        ConnectionStatus.Content = "Source connection established. Please check Target crendentials";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.OrangeRed);
                        ConnChk.Content = "Fail";
                        ConnChk.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    else if (!ConnectionFlag1 && ConnectionFlag2)
                    {
                        ConnectionStatus.Content = "Target connection established. Please check source crendentials";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.OrangeRed);
                        ConnChk.Content = "Fail";
                        ConnChk.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    else
                    {
                        ConnectionStatus.Content = "Sorry, neither Source nor Target could be connected. Verify Credentials";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.OrangeRed);
                        ConnChk.Content = "Fail";
                        ConnChk.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    break;
                case "DBToFile_Checked":
                    switch (db1)
                    {
                        case "Oracle":
                            ConnectionFlag1=ConnectOracleDB(GrpBxDBToFile_HostName.Text, GrpBxDBToFile_ServiceName.Text, GrpBxDBToFile_UserName.Text, GrpBxDBToFile_Password.Password, GrpBxDBToFile_Port.Text);
                            break;
                        case "MySQL":
                            ConnectionFlag1 = ConnectSQLDB(GrpBxDBToFile_HostName.Text, GrpBxDBToFile_ServiceName.Text, GrpBxDBToFile_UserName.Text, GrpBxDBToFile_Password.Password, GrpBxDBToFile_Port.Text);
                            break;
                        default:                            
                            exception = "Couldn't connect";
                            break;
                    }
                    if (ConnectionFlag1)
                    {
                        GrpBxDBToFile_ConnectionStatus.Content = "Connection Established";
                        GrpBxDBToFile_ConnectionStatus.Foreground = new SolidColorBrush(Colors.PaleGreen);
                        GrpBxDBToFile_ConnChk.Content = "Pass";
                        GrpBxDBToFile_ConnChk.Foreground = new SolidColorBrush(Colors.PaleGreen);
                    }
                    else
                    {
                        GrpBxDBToFile_ConnectionStatus.Content = exception.ToString() + " ||  Please check connection";
                        GrpBxDBToFile_ConnectionStatus.Foreground = new SolidColorBrush(Colors.OrangeRed);
                        GrpBxDBToFile_ConnChk.Content = "Fail";
                        GrpBxDBToFile_ConnChk.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    break;
                default:
                    break;

            }
        }
      
        //End Connection Establishment

        //Select Excel File with queries and get it's absolute path

        //End AbsolutePath for Excel File

        //Picks up 2 queries at a time and starts comapring the results
        private void CompareButton_Click(object sender, RoutedEventArgs e)
        {
            ExcelPackage excel = new ExcelPackage();
            ExceptionsOutput = Directory.GetCurrentDirectory().ToString() + (@"\Exceptions" + DateTime.Now.ToString("yyyy:MM:dd:hh:mm:ss").Replace(@":", "") + ".txt");
            OutputFile = Directory.GetCurrentDirectory().ToString() + (@"\Results" + DateTime.Now.ToString("yyyy:MM:dd:hh:mm:ss").Replace(@":", "") + ".xlsx");

            if (CompareType == "DBToDB_Checked")
            {
                try
                {
                    if (!ConnectionFlag1 || !ConnectionFlag2)
                    {
                        ConnectionStatus.Content = "Connection is not established, Please connect to DB";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.Yellow);
                    }
                    else if (!FileUploadFlag)
                    {
                        ConnectionStatus.Content = "Please select the excel file";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.Yellow);
                    }
                    else
                    {
                        OpenResultButton.IsEnabled = false;
                        ExceptionButton.IsEnabled = false;
                        ConnectionStatus.Content = "Connection Established";
                        ConnectionStatus.Foreground = new SolidColorBrush(Colors.Green);
                        List<string> Matrix = new List<string>();
                        var headerRow = new List<string[]>()
                        {
                         new string[] { "Test Case ID","Test Case Description","Query In Source", "Query In Target", "Number of mismatches", "Status","If Exception has been raised" }
                        };
                        excel.Workbook.Worksheets.Add("ResultSummary");
                        var ResultSheet = excel.Workbook.Worksheets["ResultSummary"];
                        ResultSheet.Cells["A1:G1"].LoadFromArrays(headerRow);
                        Matrix = ReadInExcel(filePath);
                        if (Matrix.Count < 1)
                        {
                            CompareException.Content = "Looks like there are no queries in the uploaded file";
                            CompareException.Foreground = new SolidColorBrush(Colors.Yellow);
                        }
                        else
                        {
                            CompareException.Content = "";
                            int i = 0;
                            exception = "";
                            ResultSheet.Cells["A1:G1"].Style.Font.Bold = true;
                            ResultSheet.Cells["A1:G1"].Style.Font.Size = 14;
                            ResultSheet.Cells["A1:G1"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                            ResultSheet.Cells["A1:G1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                            ResultSheet.Cells["A1:G1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleGreen);
                            ResultSheet.Cells[ResultSheet.Dimension.Address].AutoFitColumns();

                            for (i = 0; i < Matrix.Count; i = i + 4)
                            {
                                Compare(Matrix[i + 2], Matrix[i + 3], ConnectionString);

                                ResultSheet.SetValue(i / 2 + 2, 1, Matrix[i]);
                                ResultSheet.SetValue(i / 2 + 2, 2, Matrix[i + 1]);
                                ResultSheet.SetValue(i / 2 + 2, 3, Matrix[i + 2]);
                                ResultSheet.SetValue(i / 2 + 2, 4, Matrix[i + 3]);
                                ResultSheet.SetValue(i / 2 + 2, 5, ErrorCount);
                                ResultSheet.Cells[i / 2 + 2, 6].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                if (ErrorCount == 0 && exception == "")
                                {
                                    ResultSheet.SetValue(i / 2 + 2, 6, "PASS");
                                    ResultSheet.Cells[i / 2 + 2, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Green);
                                    ResultSheet.Cells[i / 2 + 2, 6].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                }
                                else
                                {
                                    ResultSheet.SetValue(i / 2 + 2, 6, "FAILED");
                                    ResultSheet.Cells[i / 2 + 2, 6].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                    ResultSheet.Cells[i / 2 + 2, 6].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Red);
                                }
                                if (exception != "")
                                {
                                    ResultSheet.SetValue(i / 2 + 2, 7, exception);
                                    ResultSheet.Cells[i / 2 + 2, 7].Style.Font.Color.SetColor(System.Drawing.Color.Red);
                                }
                                else
                                {
                                    ResultSheet.SetValue(i / 2 + 2, 7, "NO");
                                    ResultSheet.Cells[i / 2 + 2, 7].Style.Font.Color.SetColor(System.Drawing.Color.Green);
                                }
                                if (ErrorCount != 0)
                                {
                                    excel.Workbook.Worksheets.Add("Testcase" + (i / 4 + 1).ToString());
                                    var Testcase = excel.Workbook.Worksheets["Testcase" + (i / 4 + 1).ToString()];
                                    headerRow = new List<string[]>()
                                    {
                                        new string[] { "Table name in Source", "Table name in Target","Source Table Value","Target Table Value", "Index compared at [i,j]" }
                                    };
                                    Testcase.Cells["A1:E1"].LoadFromArrays(headerRow);
                                    Testcase.Cells["A1:E1"].Style.Font.Bold = true;
                                    Testcase.Cells["A1:E1"].Style.Font.Size = 14;
                                    Testcase.Cells["A1:E1"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                                    Testcase.Cells["A1:E1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                    Testcase.Cells["A1:E1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.Orange);
                                    Testcase.Cells[ResultSheet.Dimension.Address].AutoFitColumns();
                                    for (int j = 0; j < Mismatch.Count; j = j + 6)
                                    {
                                        Testcase.SetValue(j / 6 + 2, 1, Mismatch[j]);
                                        Testcase.SetValue(j / 6 + 2, 2, Mismatch[j + 1]);
                                        Testcase.SetValue(j / 6 + 2, 3, Mismatch[j + 2]);
                                        Testcase.SetValue(j / 6 + 2, 4, Mismatch[j + 3]);
                                        Testcase.SetValue(j / 6 + 2, 5, (T1 + "[" + Mismatch[j + 4] + "," + Mismatch[j + 5] + "]").ToString());
                                    }
                                    Mismatch.Clear();
                                }
                                exception = "";
                                ErrorCount = 0;
                            }
                        }
                    }
                    FileInfo excelFile = new FileInfo(OutputFile);
                    excel.SaveAs(excelFile);
                    excel.Dispose();

                    OpenResultButton.IsEnabled = true;
                    if (exception != "")
                    {
                        ExceptionButton.IsEnabled = true;
                    }
                }

                catch (Exception ex)
                {
                    exception = ex.GetType().ToString();
                    GrpBxDBToFile_CompareException.Content = "Errors Logged, refer the Log File";
                    GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.Yellow);
                    File.AppendAllText(ExceptionsOutput, "Issue while dealing with query : " + " -------------" + ex.ToString() + "\n");
                }
               
                   
                //End Compare all queries taken two at a time
            }
            else if (CompareType == "DBToFile_Checked")
            {
                try
                {
                    if (!ConnectionFlag1)
                    {
                        GrpBxDBToFile_ConnectionStatus.Content = "Connection is not established, Please connect to DB";
                        GrpBxDBToFile_ConnectionStatus.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    else if (!FileUploadFlag)
                    {
                        GrpBxDBToFile_CompareException.Content = "Please select the excel file";
                        GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    else if (GrpBxDBToFile_QueryTxtBox.Text=="")
                    {
                        GrpBxDBToFile_CompareException.Content = "Please enter a query before comparing";
                        GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    }
                    else
                    {

                        ErrorCount = 0;
                        exception = "";

                        DataSet ResultTable2 = Compare(GrpBxDBToFile_QueryTxtBox.Text, filePath);

                        GrpBxDBToFile_ConnectionStatus.Content = "Connection is established";
                        GrpBxDBToFile_CompareException.Content = "";
                        GrpBxDBToFile_ConnectionStatus.Foreground = new SolidColorBrush(Colors.PaleGreen);                        
                        excel.Workbook.Worksheets.Add("ResultSummary");
                        var ResultSheet = excel.Workbook.Worksheets["ResultSummary"];
                        var headerRow = new List<string[]>()
                                    {
                                        new string[] {"DB Table Value","Data from File", "Index compared at [i,j]" }
                                    };

                        ResultSheet.Cells["A1:C1"].LoadFromArrays(headerRow);
                        ResultSheet.Cells["A1:C1"].Style.Font.Bold = true;
                        ResultSheet.Cells["A1:C1"].Style.Font.Size = 14;
                        ResultSheet.Cells["A1:C1"].Style.Font.Color.SetColor(System.Drawing.Color.Blue);
                        ResultSheet.Cells["A1:C1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ResultSheet.Cells["A1:C1"].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleGreen);
                        ResultSheet.Cells[ResultSheet.Dimension.Address].AutoFitColumns();                        

                        if (ErrorCount == 0 && exception == "")
                        {
                            GrpBxDBToFile_CompareException.Content = "There has been a complete match with 0 errors";
                            GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.Green);
                            //File.AppendAllText(@"C:\Users\rabajaj\Desktop\Errorlist.txt", "There has been a complete match with 0 errors");
                        }
                        else if(exception != "")
                        {
                            File.AppendAllText(ExceptionsOutput, exception);
                            GrpBxDBToFile_CompareException.Content = "Exception raised, refer logs";
                            GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.Red);
                            //File.AppendAllText(@"C:\Users\rabajaj\Desktop\Errorlist.txt", "Errors has been logged to the excel file");
                        }     
                        else if(exception!="" && ErrorCount==0)
                        {
                            GrpBxDBToFile_CompareException.Content = "Process Completed, refer logs";
                            GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.Red);
                        }
                        
                        if (ErrorCount != 0)
                        {
                            GrpBxDBToFile_CompareException.Content = "Errors has been logged to the Output Files";
                            GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.Green);

                            for (int j = 0; j < Mismatch.Count; j = j + 4)
                            {
                                ResultSheet.SetValue(j / 4 + 2, 1, Mismatch[j]);
                                ResultSheet.SetValue(j / 4 + 2, 2, Mismatch[j + 1]);
                                ResultSheet.SetValue(j / 4 + 2, 3, (T1 + "[" + Mismatch[j + 2] + "," + Mismatch[j + 3] + "]").ToString());
                            }

                            excel.Workbook.Worksheets.Add("MismatchesOnFile");
                            var HighlightedErrors = excel.Workbook.Worksheets["MismatchesOnFile"];
                            int k = 0;
                            for (int i=0;i<ResultTable2.Tables[0].Rows.Count;i++)
                            {
                                for(int j=0;j<ResultTable2.Tables[0].Columns.Count;j++)
                                {
                                    HighlightedErrors.SetValue(i+1,j+1,ResultTable2.Tables[0].Rows[i][j].ToString());
                                    if (binary[k] ==0)
                                    {
                                        HighlightedErrors.Cells[i+1,j+1].Style.Font.Color.SetColor(System.Drawing.Color.Black);
                                        HighlightedErrors.Cells[i+1,j+1].Style.Fill.PatternType = ExcelFillStyle.Solid;
                                        HighlightedErrors.Cells[i+1,j+1].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.PaleVioletRed);
                                    }
                                    k++;
                                }
                            }
                            k = 0;
                            Mismatch.Clear();
                            binary.Clear();
                            GrpBxDBToFile_CompareException.Content = "Process Completed Please Check Excel logs";
                            GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.YellowGreen);                           
                        }
                        
                        FileInfo excelFile = new FileInfo(OutputFile);
                        excel.SaveAs(excelFile);
                        excel.Dispose();
                        if (ErrorCount != 0) { GrpBxDBToFile_OpenResultButton.IsEnabled = true; }
                        if (exception != "") { GrpBxDBToFile_ExceptionButton.IsEnabled = true; }
                        
                    }
                   

                }
                catch (Exception ex)
                {
                    exception = ex.ToString();
                    GrpBxDBToFile_CompareException.Content = "Exception raised, please refer Exception log";
                    GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.OrangeRed);
                    File.AppendAllText(ExceptionsOutput, exception);
                    if (exception != "") { GrpBxDBToFile_ExceptionButton.IsEnabled = true; }
                    exception = "";
                    ErrorCount = 0;
                }
            }
        }

        public DataSet Compare(string queryString, string AbsolutePath)
        {
            int NumOfRow = 0;
            int NumOfCol = 0;
            Guid guid;
            // For File to Db
            DataSet ResultTable1 = new DataSet();
            DataSet ResultTable2 = new DataSet();
            DataSet SchemaTable = new DataSet();
            T1 = TableName(queryString);
            Dictionary<string, string> Schema = new Dictionary<string, string>();
            string SchemaQuery = @"SELECT  column_name,data_type FROM all_tab_columns where table_name = '" + T1 + "'";

            ResultTable2 = ReadInExcel();
            switch (db1)
            {
                case "Oracle":
                    OracleConnection.ConnectionString = ConnectionString;
                    OracleCommand OracleCommand1 = new OracleCommand(queryString, OracleConnection);
                    OracleCommand OracleCommand2 = new OracleCommand(SchemaQuery, OracleConnection);
                    FillDataSet(OracleCommand1, ResultTable1);
                    FillDataSet(OracleCommand2, SchemaTable);
                    break;
                case "SQL":
                    SQLconnection.ConnectionString = ConnectionString;
                    SqlCommand SqlCommand1 = new SqlCommand(queryString, SQLconnection);
                    SqlCommand SqlCommand2 = new SqlCommand(SchemaQuery, SQLconnection);
                    FillDataSet(SqlCommand1, ResultTable1); ;
                    FillDataSet(SqlCommand2, SchemaTable);
                    break;
            }

            try
            {
                string k = string.Empty;
                Schema = TableSchema(SchemaTable);
                string colname = "";
                NumOfRow = (ResultTable1.Tables[0].Rows.Count) < ResultTable2.Tables[0].Rows.Count ? (ResultTable1.Tables[0].Rows.Count) : ResultTable2.Tables[0].Rows.Count;
                NumOfCol = (ResultTable1.Tables[0].Columns.Count) < ResultTable2.Tables[0].Columns.Count ? (ResultTable1.Tables[0].Columns.Count) : ResultTable2.Tables[0].Columns.Count;

                if ((ResultTable1.Tables[0].Columns.Count) == ((ResultTable2.Tables[0].Columns.Count)))
                {
                    for (int i = 0; i < NumOfRow; i++)
                    {
                        for (int j = 0; j < NumOfCol; j++)
                        {
                            colname = ResultTable1.Tables[0].Columns[j].ToString();
                            switch (Schema[colname])
                            {
                                case "RAW":
                                    guid = new Guid((byte[])ResultTable1.Tables[0].Rows[i][j]);
                                    if (DotNetToOracle(guid.ToString().Replace("-", "").ToUpper()) != ResultTable2.Tables[0].Rows[i][j].ToString())
                                    {
                                        Mismatch.Add(DotNetToOracle(guid.ToString().Replace("-", "").ToUpper()));
                                        Mismatch.Add(ResultTable2.Tables[0].Rows[i][j].ToString());
                                        Mismatch.Add((i + 1).ToString());
                                        Mismatch.Add((j + 1).ToString());
                                        ErrorCount += 1;
                                        binary.Add(0);
                                    }
                                    else { binary.Add(1); }
                                    break;
                                default:
                                    if (ResultTable1.Tables[0].Rows[i][j].ToString() != ResultTable2.Tables[0].Rows[i][j].ToString())
                                    {
                                        Mismatch.Add(ResultTable1.Tables[0].Rows[i][j].ToString());
                                        Mismatch.Add(ResultTable2.Tables[0].Rows[i][j].ToString());
                                        Mismatch.Add((i + 1).ToString());
                                        Mismatch.Add((j + 1).ToString());
                                        ErrorCount += 1;
                                        binary.Add(0);
                                    }
                                    else { binary.Add(1); }
                                    break;
                            }
                        }
                    }
                    //Calls functions to  Parse any rows/columns missed for n1*m1 and n2*m2 Dataset
                    if ((ResultTable1.Tables[0].Rows.Count) != (ResultTable2.Tables[0].Rows.Count))
                    {
                        if ((ResultTable1.Tables[0].Rows.Count) > (ResultTable2.Tables[0].Rows.Count))
                        {
                            RowCheck(ResultTable1, Math.Abs((ResultTable1.Tables[0].Rows.Count) - (ResultTable2.Tables[0].Rows.Count)), NumOfRow, NumOfCol);
                        }
                        else
                        {
                            RowCheck(ResultTable2, Math.Abs((ResultTable1.Tables[0].Rows.Count) - (ResultTable2.Tables[0].Rows.Count)), NumOfRow, NumOfCol);
                        }
                    }                   
                }
                else
                {
                    exception = "Looks like the Tables have different number of columns";
                }

                //if ((ResultTable1.Tables[0].Columns.Count) != (ResultTable2.Tables[0].Columns.Count))
                //{
                //    if ((ResultTable1.Tables[0].Columns.Count) > (ResultTable2.Tables[0].Columns.Count))
                //    {
                //        ColumnCheck(ResultTable1, Math.Abs((ResultTable1.Tables[0].Columns.Count) - (ResultTable2.Tables[0].Columns.Count)), NumOfRow, NumOfCol);
                //    }
                //    else
                //    {
                //        ColumnCheck(ResultTable2, Math.Abs((ResultTable1.Tables[0].Columns.Count) - (ResultTable2.Tables[0].Columns.Count)), NumOfRow, NumOfCol);
                //    }
                //}
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                GrpBxDBToFile_CompareException.Content = "Errors Logged, refer the Log File";
                GrpBxDBToFile_CompareException.Foreground = new SolidColorBrush(Colors.Yellow);
                File.AppendAllText(ExceptionsOutput, "Issue while dealing with query : " + queryString + "                       " + " -------------" + ex.ToString() + "\n");
            }           
            return ResultTable2;
        }

        //Function to Populate two datasets and compare the results, also log the differences
        public void Compare(string queryString1, string queryString2, string ConnectionString)
        {
            
            DataSet ResultTable1 = new DataSet();
            DataSet ResultTable2 = new DataSet();
            DataSet SchemaTable1 = new DataSet();
            DataSet SchemaTable2 = new DataSet();

            Dictionary<string, string> Schema1 = new Dictionary<string, string>();
            Dictionary<string, string> Schema2 = new Dictionary<string, string>();
            T1 = TableName(queryString1);
            T2 = TableName(queryString2);

            string SchemaQuery1 = @"SELECT  column_name,data_type FROM all_tab_columns where table_name = '" + T1  + "'";
            string SchemaQuery2 = @"SELECT  column_name,data_type FROM all_tab_columns where table_name = '" + T2 + "'";
            switch (db1)
            {
                case "Oracle":
                    OracleConnection.ConnectionString = ConnectionString;
                    OracleCommand OracleCommand1 = new OracleCommand(queryString1, OracleConnection);
                    OracleCommand OracleCommand2 = new OracleCommand(SchemaQuery1, OracleConnection);
                    FillDataSet(OracleCommand1, ResultTable1);
                    FillDataSet(OracleCommand2, SchemaTable1);                
                    break;
                case "SQL":
                    SQLconnection.ConnectionString = ConnectionString;
                    SqlCommand SqlCommand1 = new SqlCommand(queryString1, SQLconnection);
                    SqlCommand SqlCommand2 = new SqlCommand(SchemaQuery1, SQLconnection);
                    FillDataSet(SqlCommand1, ResultTable1);
                    FillDataSet(SqlCommand2, SchemaTable1);
                    break;                    
            }

            switch (db2)
            {
                case "Oracle":
                    OracleConnection.ConnectionString = ConnectionString;
                    OracleCommand OracleCommand3 = new OracleCommand(queryString2, OracleConnection);
                    OracleCommand OracleCommand4 = new OracleCommand(SchemaQuery2, OracleConnection);
                    FillDataSet(OracleCommand3, ResultTable2);
                    FillDataSet(OracleCommand4, SchemaTable2);
                    break;
                case "SQL":
                    SQLconnection.ConnectionString = ConnectionString;
                    SqlCommand SqlCommand3 = new SqlCommand(queryString2, SQLconnection);
                    SqlCommand SqlCommand4 = new SqlCommand(SchemaQuery2, SQLconnection);
                    FillDataSet(SqlCommand3, ResultTable2);
                    FillDataSet(SqlCommand4, SchemaTable2);
                    break;
            }
            try
            {

                Schema1 = TableSchema(SchemaTable1);
                Schema2 = TableSchema(SchemaTable2);
                if (SchemaCheckFlag)
                {
                    if (Schema1.OrderBy(kvp => kvp.Key).SequenceEqual(Schema2.OrderBy(kvp => kvp.Key)))
                    {
                        DataTypeChk.Content = "Pass";
                        DataTypeChk.Foreground = new SolidColorBrush(Colors.PaleGreen);
                    }
                    else
                    {
                        DataTypeChk.Content = "Fail";
                        DataTypeChk.Foreground = new SolidColorBrush(Colors.Yellow);
                    }
                }
                else
                {
                    DataTypeChk.Content = "Skipped";
                    DataTypeChk.Foreground = new SolidColorBrush(Colors.YellowGreen);
                }

                //Comparing the Datasets now
                string colname = "";

                int NumOfRow = (ResultTable1.Tables[0].Rows.Count) < (ResultTable2.Tables[0].Rows.Count) ? (ResultTable1.Tables[0].Rows.Count) : (ResultTable2.Tables[0].Rows.Count);
                int NumOfCol = (ResultTable1.Tables[0].Columns.Count) < ((ResultTable2.Tables[0].Columns.Count)) ? ResultTable1.Tables[0].Columns.Count : ResultTable2.Tables[0].Columns.Count;

                if ((ResultTable1.Tables[0].Columns.Count) == ((ResultTable2.Tables[0].Columns.Count)))
                {
                    for (int i = 0; i < NumOfRow; i++)
                    {
                        for (int j = 0; j < NumOfCol; j++)
                        {
                            colname = ResultTable1.Tables[0].Columns[j].ToString();

                            switch (Schema1[colname])
                            {
                                case "RAW":
                                    Guid guid1 = new Guid((byte[])ResultTable1.Tables[0].Rows[i][j]);
                                    Guid guid2 = new Guid((byte[])ResultTable2.Tables[0].Rows[i][j]);

                                    if (guid1 != guid2)
                                    {
                                        Mismatch.Add(T1);
                                        Mismatch.Add(T2);
                                        Mismatch.Add(guid1.ToString());
                                        Mismatch.Add(guid2.ToString());
                                        Mismatch.Add((i + 1).ToString());
                                        Mismatch.Add((j + 1).ToString());
                                        ErrorCount += 1;
                                    }
                                    break;
                                default:
                                    if (ResultTable1.Tables[0].Rows[i][j].ToString() != ResultTable2.Tables[0].Rows[i][j].ToString())
                                    {
                                        Mismatch.Add(T1);
                                        Mismatch.Add(T2);
                                        Mismatch.Add(ResultTable1.Tables[0].Rows[i][j].ToString());
                                        Mismatch.Add(ResultTable2.Tables[0].Rows[i][j].ToString());
                                        Mismatch.Add((i + 1).ToString());
                                        Mismatch.Add((j + 1).ToString());
                                        ErrorCount += 1;
                                    }
                                    break;
                            }
                        }
                    }

                    //Calls functions to  Parse any rows missed for n1*m1 and n2*m1 Dataset
                    if ((ResultTable1.Tables[0].Rows.Count) != (ResultTable2.Tables[0].Rows.Count))
                    {
                        if ((ResultTable1.Tables[0].Rows.Count) > (ResultTable2.Tables[0].Rows.Count))
                        {
                            RowCheck(ResultTable1, Math.Abs((ResultTable1.Tables[0].Rows.Count) - (ResultTable2.Tables[0].Rows.Count)), NumOfRow, NumOfCol);
                        }
                        else
                        {
                            RowCheck(ResultTable2, Math.Abs((ResultTable1.Tables[0].Rows.Count) - (ResultTable2.Tables[0].Rows.Count)), NumOfRow, NumOfCol);
                        }
                    }
                }
                else
                {
                    exception = "Looks like the Tables have different number of columns";
                }





                //if ((ResultTable1.Tables[0].Columns.Count) != (ResultTable2.Tables[0].Columns.Count))
                //{
                //    if ((ResultTable1.Tables[0].Columns.Count) > (ResultTable2.Tables[0].Columns.Count))
                //    {
                //        ColumnCheck(ResultTable1, Math.Abs((ResultTable1.Tables[0].Columns.Count) - (ResultTable2.Tables[0].Columns.Count)), NumOfRow, NumOfCol);
                //    }
                //    else
                //    {
                //        ColumnCheck(ResultTable2, Math.Abs((ResultTable1.Tables[0].Columns.Count) - (ResultTable2.Tables[0].Columns.Count)), NumOfRow, NumOfCol);
                //    }
                //}
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                CompareException.Content = "Errors Logged, refer the Log File";
                CompareException.Foreground = new SolidColorBrush(Colors.Yellow);
                File.AppendAllText(ExceptionsOutput, "Issue with query : " + queryString1 + "      ||||      " + queryString2 + "               " + " -------------" + ex.ToString() + "\n");
            }            
            
        }



        //-------------------------------------------------------------------------------HELPER FUNCTIONS--------------------------------------------------------------------------------------------------//






        //Read the Selected Excel and populate the queries in a 2D Matrix
        public List<string> ReadInExcel(string absolutePath)
        {
            try
            {
                List<string> Matrix = new List<string>();
                List<string> track = new List<string>();
                Application xlApp = new Application();
                Workbook xlWorkbook = xlApp.Workbooks.Open(absolutePath);
                _Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                for (int i = 1; i <= rowCount; i++)
                {
                    for (int j = 1; j <= colCount; j++)
                    {
                        //write the value to the console
                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                            track.Add(xlRange.Cells[i, j].Value2.ToString());

                        //add useful things here!   
                    }
                    Matrix.AddRange(track);
                    track.Clear();
                }
                return Matrix;
            }
            catch (Exception ex)
            {
                FileUploadFlag = false;
                exception = ex.GetType().ToString();
                CompareException.Content = "Please try re-uploading the file";
                return null;
            }

        }
        //End Populate Queries in Matrix

        public DataSet ReadInExcel()
        {
            string connectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Persist Security Info=False;Data Source=" + filePath + ";Extended Properties="+ "\"Excel 12.0;HDR = YES; IMEX = 1;\"");
            DataSet data = new DataSet();

            foreach (var sheetName in GetExcelSheetNames(connectionString))
            {
                try
                {
                    using (OleDbConnection con = new OleDbConnection(connectionString))
                    {
                        var dataTable = new System.Data.DataTable();
                        string query = string.Format("SELECT * FROM [{0}]", sheetName);
                        con.Open();
                        OleDbDataAdapter adapter = new OleDbDataAdapter(query, con);
                        adapter.Fill(dataTable);
                        data.Tables.Add(dataTable);
                    }
                }
                catch(Exception ex)
                {
                    exception = ex.GetType().ToString();
                    CompareException.Foreground = new SolidColorBrush(Colors.Yellow);
                    File.AppendAllText(ExceptionsOutput, ex.ToString() + "\n");
                    return null;
                }
            }
            return data;
        }

        public string[] GetExcelSheetNames(string connectionString)
        {
            try
            {
                OleDbConnection con = null;
                System.Data.DataTable dt = null;
                con = new OleDbConnection(connectionString);
                con.Open();
                dt = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);

                if (dt == null)
                {
                    return null;
                }

                string[] excelSheetNames = new string[dt.Rows.Count];
                int i = 0;

                foreach (DataRow row in dt.Rows)
                {
                    excelSheetNames[i] = row["TABLE_NAME"].ToString();
                    i++;
                }
                return excelSheetNames;
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                CompareException.Foreground = new SolidColorBrush(Colors.Yellow);
                File.AppendAllText(ExceptionsOutput, "Issue with the uploaded excel, Please try reuploading it  \n" + "                       " + ex.ToString() + "\n");
                return null;
            }


        }


        //Parses any rows missed for n1*m1 and n2*m2 Dataset
        public void RowCheck(DataSet ResultTable, int diff, int NumOfRow, int NumOfCol)
        {
            for (int i = NumOfRow; i < NumOfRow + diff; i++)
            {
                for (int j = 0; j < NumOfCol; j++)
                {
                    BlankValues.Add(ResultTable.Tables[0].Rows[i][j].ToString());
                    //File.AppendAllText(Directory.GetCurrentDirectory().ToString() + @"\Errorlist.txt", "Uncompared value: " + ResultTable.Tables[0].Rows[i][j].ToString() + "                 " + DateTime.Now.ToString("yyyy:MM:dd:hh:mm:ss:ffff") + "\n");
                    ErrorCount += 1;
                }
            }
        }
        //Parses any columns missed for n1*m1 and n2*m2 Dataset
        //public void ColumnCheck(DataSet ResultTable, int diff, int NumOfRow, int NumOfCol)
        //{
        //    for (int i = 0; i < NumOfRow; i++)
        //    {
        //        for (int j = NumOfCol; j < NumOfCol + diff; j++)
        //        {
        //            BlankValues.Add(ResultTable.Tables[0].Rows[i][j].ToString());
        //            //File.AppendAllText(Directory.GetCurrentDirectory().ToString() + @"\Errorlist.txt", "Uncompared value: " + ResultTable.Tables[0].Rows[i][j].ToString() + "                 " + DateTime.Now.ToString("yyyy:MM:dd:hh:mm:ss:ffff") + "\n");
        //            ErrorCount += 1;
        //        }
        //    }
        //}
        //End Row and Column Check

        //Overloaded Function which sets the command depending on the type of connection
        public OracleCommand[] SetCommands(string queryString1, string queryString2, OracleConnection connection)
        {
            string SchemaQuery1 = @"SELECT  column_name,data_type,column_id FROM all_tab_columns where table_name = '" + TableName(queryString1) + "' order by column_id";
            string SchemaQuery2 = @"SELECT  column_name,data_type,column_id FROM all_tab_columns where table_name = '" + TableName(queryString2) + "' order by column_id";

            OracleCommand[] AllCommands = new OracleCommand[4];
            AllCommands[0] = new OracleCommand(queryString1, connection);
            AllCommands[1] = new OracleCommand(queryString2, connection);
            AllCommands[2] = new OracleCommand(SchemaQuery1, connection);
            AllCommands[3] = new OracleCommand(SchemaQuery2, connection);

            return AllCommands;
        }


        public SqlCommand[] SetCommands(string queryString1, string queryString2, SqlConnection connection)
        {
            string SchemaQuery1 = @"SELECT  column_name,data_type FROM all_tab_columns where table_name = '" + TableName(queryString1) + "'";
            string SchemaQuery2 = @"SELECT  column_name,data_type FROM all_tab_columns where table_name = '" + TableName(queryString2) + "'";

            SqlCommand[] AllCommands = new SqlCommand[4];
            AllCommands[0] = new SqlCommand(queryString1, connection);
            AllCommands[1] = new SqlCommand(queryString2, connection);
            AllCommands[2] = new SqlCommand(SchemaQuery1, connection);
            AllCommands[3] = new SqlCommand(SchemaQuery2, connection);

            return AllCommands;
        }


        //End SetCommands

        //Returns the TableSchema
        private Dictionary<string, string> TableSchema(DataSet SchemaTable)
        {
            try
            {
                Dictionary<string, string> Schema = new Dictionary<string, string>();
                for (int i = 0; i < SchemaTable.Tables[0].Rows.Count; i++)
                {
                    Schema.Add(SchemaTable.Tables[0].Rows[i][0].ToString(), SchemaTable.Tables[0].Rows[i][1].ToString());
                }
                return Schema;
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                File.AppendAllText(ExceptionsOutput, "Issue getting the Table Schema       " + ex.ToString() + "\n");
                return null;
            }

        }
        //End TableSchema

        //Returns the name of the Table from the query
        private string TableName(string queryString)
        {
            try
            {
                int k = 0;
                string[] temp = queryString.ToLower().Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                while (temp[k].ToLower() != "from")
                {
                    k++;
                }
                string tablename = temp[k + 1].ToString().Substring(temp[k + 1].ToString().LastIndexOf('.') + 1);
                return tablename.ToUpper();
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                File.AppendAllText(ExceptionsOutput, "Issue with DB table name" + ex.ToString() + "\n");
                return null;
            }

        }
        //End TableName

        //Overloaded Function which Fills the datasets depending on the type of connection
        public void FillDataSet(OracleCommand command, DataSet Table)
        {
            OracleDataAdapter adapter = new OracleDataAdapter(command);
            try
            {
                adapter = new OracleDataAdapter(command);
                adapter.Fill(Table);
                adapter.Dispose();
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                File.AppendAllText(ExceptionsOutput, "Issue filling data from adapter to dataset" + ex.ToString() + "     -------      " + "\n");                
            }

        }

        public void FillDataSet(SqlCommand command, DataSet Table)
        {
            SqlDataAdapter adapter = new SqlDataAdapter(command);
            try
            {
                adapter = new SqlDataAdapter(command);
                adapter.Fill(Table);
                adapter.Dispose();
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = "c:\\";
                openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == true)
                {
                    GrpBxDBToFile_CompareException.Content = "";
                    FileUploadFlag = true;
                    filePath = openFileDialog.FileName.ToString();
                    FileNameLabel.Content = filePath;
                    FileNameLabel.Foreground = new SolidColorBrush(Colors.White);
                    GrpBxDBToFile_FileNameLabel.Content = filePath;
                    GrpBxDBToFile_FileNameLabel.Foreground = new SolidColorBrush(Colors.White);
                }
            }
            catch (Exception ex)
            {
                exception = ex.GetType().ToString();
                FileNameLabel.Content = "Please select only excel file";
                FileNameLabel.Foreground = new SolidColorBrush(Colors.Yellow);
                GrpBxDBToFile_FileNameLabel.Content= "Please select only excel file";
                GrpBxDBToFile_FileNameLabel.Foreground = new SolidColorBrush(Colors.Yellow);
            }

        }

        private void GrpBxDBToFile_Port_TextInput(object sender, TextCompositionEventArgs e)
        {

        }

        private void OpenResultButton_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(OutputFile))
            {
                Process.Start(OutputFile);
            }
        }

        private void ExceptionButton_Click(object sender, RoutedEventArgs e)
        {
            if (File.Exists(ExceptionsOutput))
            {
                Process.Start(ExceptionsOutput);
            }
        }

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            SchemaCheckFlag = true;
        }

        //End FillDataSet

        private void NumericOnly(object sender, TextCompositionEventArgs e)
        {

            Regex reg = new Regex("[^0-9]");
            e.Handled = reg.IsMatch(e.Text);

        }

        static string OracleToDotNet(string text)
        {
            byte[] bytes = ParseHex(text);
            Guid guid = new Guid(bytes);
            return guid.ToString("N").ToUpperInvariant();
        }

        static string DotNetToOracle(string text)
        {
            Guid guid = new Guid(text);
            return BitConverter.ToString(guid.ToByteArray()).Replace("-", "");
        }

        static byte[] ParseHex(string text)
        {
            // Not the most efficient code in the world, but
            // it works...
            byte[] ret = new byte[text.Length / 2];
            for (int i = 0; i < ret.Length; i++)
            {
                ret[i] = Convert.ToByte(text.Substring(i * 2, 2), 16);
            }
            return ret;
        }
    }
}