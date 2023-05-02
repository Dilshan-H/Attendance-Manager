using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace AttendanceManager
{
    public partial class FindDataWindow : Form
    {
        delegate void SetTextCallback(string text);
        string FilePath = Application.LocalUserAppDataPath + @"\ApplicationData\EmpNames.dat";
        string BiometricsData = Application.LocalUserAppDataPath + @"\ApplicationData\";
        string ReportFilePath = Application.ExecutablePath.ToLower().Replace("attendance manager.exe", "Report Template.docx");
        string FullReportPath = Application.ExecutablePath.ToLower().Replace("attendance manager.exe", "FullReport.docx");
        string ReportPath;
        
        string[] DATLines = { };
        List<string> Nos = new List<string>();
        List<string> Grades = new List<string>();
        List<string> DataLines = new List<string>();
        List<string> TextLines = new List<string>();
        List<string> TempEmps = new List<string>();
        string SubSearch;
        string FillItem;
        string[] Date_1;
        string[] Time_1;
        string[] Date_2;
        string[] Time_2;
        System.DateTime date1;
        System.DateTime date2;
        System.TimeSpan diff1;
        System.TimeSpan diff2;
        System.TimeSpan LunchDiff;
        System.TimeSpan OT;
        System.TimeSpan TotalOT;
        System.TimeSpan TotalDiff;
        System.TimeSpan Diff;
        Word._Document oDataDoc;
        int TotalWorkDays = 0;
        int TotalHolidays = 0;
        int TotalSat_Sun = 0;
        int TotalLateDays = 0;
        int TotalShortLeaves = 0;
        bool bottomData = false;
        bool SatOrSun = false;
        bool success = false;
        bool fileError = false;
        DateTime ThisDay;
        int MO;
        string SelYear;
        string SelMonth;
        string SelMonthText;
        bool oldData = false;
        string compYM = "";
        bool importSuccess = false;
        string name;
        string machineNo;
        string grade;
        string selDept;
        int selEmp;

        //Time periods
        //DateTime EandC_N_0 = new DateTime (2000, 01, 01, 08, 15, 59);
        //DateTime EandC_N_1 = new DateTime(2000, 01, 01, 16, 15, 00);
        //DateTime EandC_SL_0 = new DateTime(2000, 01, 01, 08, 15, 59);
        //DateTime EandC_SL_1 = new DateTime(2000, 01, 01, 01, 01, 01);

        string status;
        Word.Application wrdApp;
        Word._Document wrdDoc;
        Object oMissing = System.Reflection.Missing.Value;
        Object oFalse = false;
        LoadingWindow ShowF4 = new LoadingWindow();
        LoadingWindow ShowF4_ = new LoadingWindow();
        bool TopData = true;
        bool MissedPunche = false;
        int lastRow = 2;

        public FindDataWindow()
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker2.WorkerReportsProgress = true;
            backgroundWorker2.WorkerSupportsCancellation = true;
        }

        private void Form2_Load(object sender, EventArgs e)
        {
            panel1.Left = (this.ClientSize.Width - panel1.Width) / 2;
            groupBox2.Left = (this.ClientSize.Width - groupBox2.Width) / 2;
            comboBox2.SelectedItem = comboBox2.Items[0];
            comboBox4.SelectedItem = comboBox4.Items[0];
            radioButton2.Checked = true;
            
            toolStripStatusLabel1.Text = "Ready - Import biometrics log data or load data.";

            //fill years
            int a = 2015;
            while (a <= 2040)
            {
                comboBox3.Items.Add(a);
                a++;
            }
            
            //add default year
            DateTime now = DateTime.Now;
            int ThisYear = now.Year;
            
            if (comboBox3.Items.Contains(ThisYear))
            {
                int index = comboBox3.Items.IndexOf(ThisYear);
                comboBox3.SelectedItem = comboBox3.Items[index];
            }
            else
            {
                MessageBox.Show("Oops! Outdated version -- Check system date or contact developer.\n\nIf you seeing this message after 2040...Yeah, this may seems like cr**! Who knows what may happen in the future???", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }

            if (!System.IO.File.Exists(FilePath))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(FilePath.Replace(@"\EmpNames.dat", ""));
                    using (StreamWriter sw = File.CreateText(FilePath)){}    
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X140X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (DirectoryNotFoundException ex) 
                {
                    MessageBox.Show("ERROR:02X145X" + ex.Message);
                    Close();
                }
            }
            else
            {
                string[] lines = System.IO.File.ReadAllLines(FilePath);
                lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                List<string> Names = new List<string>();
                String[] spearator = { "\t" };
                Int32 Wordcount = 4;
                bool ErrorStat = false;
                if (lines.Length != 0)
                {
                    Array.Sort(lines, System.StringComparer.InvariantCulture);
                }
                foreach (var item in lines)
                {
                    string[] data = item.Split(spearator, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                    try
                    {
                        Names.Add(data[0]);
                        Nos.Add(data[1]);
                        Grades.Add(data[2]);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Unable to load employee data! Please restart the programme.\nIf this error persists reset employee data through settings section.\nERROR:02X173X", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ErrorStat = true;
                        break;
                    }
                    
                }

                if (ErrorStat == true)
                {
                    //System.Threading.Thread.CurrentThread.Abort();
                    Close();
                    return;
                }

                comboBox1.DataSource = Names;
                comboBox1.Refresh();
                if (Names.Count != 0)
                {
                    comboBox1.SelectedItem = comboBox1.Items[0];
                }
                else
                {
                    MessageBox.Show("Oops! Employee name list is empty. You should add users first!\nERROR:02X196X", "Empty User List", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    EmployeeManagementWindow ShowF3 = new EmployeeManagementWindow();
                    ShowF3.Show();
                    Close();
                }
                
            }


            //#############################################
            //Properties.Settings.Default.RunValue = "Yes";
            //Properties.Settings.Default.Save();
            //#############################################
        }

        private void Form2_SizeChanged(object sender, EventArgs e)
        {
            panel1.Left = (this.ClientSize.Width - panel1.Width) / 2;
            groupBox2.Left = (this.ClientSize.Width - groupBox2.Width) / 2;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {

                Title = "Choose .DAT file",
                FileName = "",
                Multiselect = false,
                CheckFileExists = true,
                Filter = "DAT files (*.dat)|*.dat",
            };

            textBox1.Text = "";
            Array.Clear(DATLines,0,DATLines.Length);

            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
            {
                string DATPath = openFileDialog1.FileName;
                textBox2.Text = DATPath;
                groupBox2.BackColor = Color.Aquamarine;
                

                if (backgroundWorker1.IsBusy != true)
                {
                    // Start the asynchronous operation.
                    compYM = SelYear + "-" + SelMonth;
                    DialogResult rep = MessageBox.Show("Do you want to scan through the whole file to import data? - This might take a while.\n\nClick\"No\" to import only the data on selected month and year.", "Import Biometrics Log Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (rep == System.Windows.Forms.DialogResult.Yes)
                    {
                        oldData = true;
                    }
                    else
                    {
                        oldData = false;
                    }
                    backgroundWorker1.RunWorkerAsync(DATPath);
                    toolStripStatusLabel1.Text = "Please wait...";
                    ShowF4.Show();
                }
            }

            
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            
                if (worker.CancellationPending == true)
                {
                    e.Cancel = true;
                    return;
                }
                else
                {

                    // Perform a time consuming operation and report progress.
                    fileError = false;
                    string tempData = e.Argument.ToString();
                    string[] readLines = {};
                    string[] addLines = { };
                    string[] tempLines = {};
                    try
                    {
                        readLines = File.ReadAllLines(tempData);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X297X", "Import Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        fileError = true;
                        Close();
                    }
                    catch (DirectoryNotFoundException ex)
                    {
                        fileError = true;
                        MessageBox.Show("ERROR" + ex.Message);
                        Close();
                    }
                    List<string> addData = new List<string>();
                    String[] spearator = { "\t" };
                    Int32 Wordcount = 2;
                    string thisYear = ""; 

                    foreach (var item in readLines)
                    {
                        if (oldData == false)
                        {
                            if (!item.Contains(compYM))
                            {
                                continue;
                            } 
                        }
                        string[] data = item.Split(spearator, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                        thisYear = data[1].Trim().Substring(0, 4);
                            string tempPath = BiometricsData + thisYear + ".dat";
                            if (!System.IO.File.Exists(tempPath))
                            {
                                try
                                {
                                    System.IO.Directory.CreateDirectory(BiometricsData);
                                    using (StreamWriter sw = File.CreateText(tempPath)) { }

                                    tempLines = File.ReadAllLines(tempPath);
                                    if (!tempLines.Contains(item))
                                    {
                                        using (StreamWriter sw = File.AppendText(tempPath))
                                        {
                                            sw.WriteLine(item);
                                        }
                                    }
                                }
                                catch (UnauthorizedAccessException)
                                {
                                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X350X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    Close();
                                }
                                catch (DirectoryNotFoundException ex)
                                {
                                    MessageBox.Show("ERROR:02X355X" + ex.Message);
                                    Close();
                                }

                            }
                            else
                            {
                                try
                                {
                                    tempLines = File.ReadAllLines(tempPath);
                                    if (!tempLines.Contains(item))
                                    {
                                        using (StreamWriter sw = File.AppendText(tempPath))
                                            {
                                                sw.WriteLine(item);
                                            }
                                        importSuccess = true;
                                    }
                                }
                                catch (UnauthorizedAccessException)
                                {
                                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X376X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    Close();
                                    return;
                                }
                                catch (DirectoryNotFoundException ex)
                                {
                                    MessageBox.Show("ERROR:02X382X" + ex.Message);
                                    Close();
                                    return;
                                }
                                
                            }
                    }

                }
            
        }

        private TimeSpan updateOT(TimeSpan OTtime)
        {
            return new TimeSpan(OTtime.Hours, (OTtime.Minutes / 15) * 15, 0);
        }

        private void backgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {           
        }

        private void backgroundWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            groupBox1.Enabled = true;
            panel1.Enabled = true;
            textBox1.Enabled = true;
            ShowF4.Hide();
            button1.Enabled = true;
            if (importSuccess)
            {
                MessageBox.Show("Biometrics Log Data successfully imported.", "Import Status", MessageBoxButtons.OK, MessageBoxIcon.Information);
                toolStripStatusLabel1.Text = "Ready - Biometrics Data successfully imported.";
            }
            else
            {
                MessageBox.Show("Sorry. Failed to import Biometrics Log Data\nPlease try again.\nCheck the month, year you've selected.", "Import Status", MessageBoxButtons.OK, MessageBoxIcon.Error);
                toolStripStatusLabel1.Text = "Ready - Failed to import Biometrics Log Data.";
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null || comboBox2.SelectedItem == null || comboBox3.SelectedItem == null)
            {
                MessageBox.Show("Oops! Invalid Inputs detected!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            textBox1.Text = "";
            MO = comboBox2.SelectedIndex;
            MO++;
            string MO2Txt = MO.ToString();
            
            int a = comboBox3.SelectedIndex;
            SelYear = comboBox3.Items[a].ToString();
            if (MO.ToString().Length != 2)
            {
                MO2Txt = "0" + MO.ToString();
            }
            SubSearch = SelYear + "-" + MO2Txt;
            int selCount;
            string selID;
            selCount = comboBox1.SelectedIndex;
            selID = Nos[selCount] +"\t";

            name = comboBox1.SelectedItem.ToString();
            machineNo = selID;
            grade = Grades[selCount];
            string tempPath = BiometricsData + SelYear + ".dat";
            if (System.IO.File.Exists(tempPath))
            {
                try
                {
                    DATLines = File.ReadAllLines(tempPath);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.", "Failed to read files.\nERROR:02X513X", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                    
                }
            }
            else
            {
                MessageBox.Show("We couldn't find requested data.\nMake sure you've imported the biometrics logs and selected the correct month, year.\nERROR:02X520X", "Resources Not Found.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            //For Testing Purposes
            //selID = "006";
            //textBox1.Text = selCount.ToString();
            //textBox3.Text = selID;
            List<string> TOSortLines = new List<string>();
            foreach (var item in DATLines)
            {
                if (item.Trim().Contains(selID) && item.Contains(SubSearch.Trim()))
                {
                    FillItem = item;
                    TOSortLines.Add(FillItem);
                }
            }

            TOSortLines.Sort();
            foreach (var line in TOSortLines)
            {
                textBox1.Text += line + Environment.NewLine;
            }

            if (textBox1.Text == "")
            {
                textBox1.Text = Environment.NewLine + "☹" + Environment.NewLine + "Sorry. No records Found!";
            }
        }

        private void Form2_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (backgroundWorker2.IsBusy == true)
            {
                MessageBox.Show("Report generating in progress...\n\nClosing form will cause unexpected errors!", "Please wait", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                toolStripStatusLabel1.Text = "File is still loading!";
                e.Cancel = true;
            }
        }

        private void backgroundWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            if (worker.CancellationPending == true)
            {
                e.Cancel = true;
                return;
            }
            else
            {
                // Perform a time consuming operation and report progress.
                if (textBox1.Text != "" && textBox1.Text != Environment.NewLine + "☹" + Environment.NewLine + "Sorry. No records Found!") 
                {
                    string[] Del = { "\t" };
                    string[] Del2 = { " " };
                    string[] Del3 = { "-" };
                    string[] Del4 = { ":" };
                    Int32 Wordcount = 2;
                    string FDate = "";
                    int count = 1;
                    //Reset variables
                    TotalWorkDays = 0;
                    TotalHolidays = 0;
                    TotalLateDays = 0;
                    TotalShortLeaves = 0;
                    bottomData = false;
                    TotalOT = new TimeSpan(0);
                    TotalDiff = new TimeSpan(0);
                    lastRow = 2;

                    string count2Txt = count.ToString();
                    bool IsFirstDay = true;
                    try
                    {
                        wrdApp = new Word.Application();
                        wrdApp.Visible = true;
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Hmm, We couldn't launch Microsoft® Word to complete the process.\nPlease verify you've installed Microsoft® Word or try opening a new Word document!\nERROR:02X604X", "Failed to launch Microsoft Word.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        backgroundWorker2.CancelAsync();
                    }
                    
                    
                    // Create an instance of Word  and make it visible.

                    try
                    {
                        Object oName = ReportPath;
                        // Open the file to insert data.
                        oDataDoc = wrdApp.Documents.Open(ref oName, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                        ref oMissing/*, ref oMissing */);
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Oops! Unexpected error occured.\nPlease contact the developer with the error code.\nERROR:02X640X", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.ExitThread();
                        Application.Exit();
                        return;
                    }


                    // Fill in the data.
                    if (TopData == true)
                    {
                        FillRow(oDataDoc, true, 1, 2, e.Argument.ToString(), null, null, null, null, null);
                        FillRow(oDataDoc, true, 2, 1, name, null, null, null, null, null);
                        FillRow(oDataDoc, true, 2, 2, machineNo, null, null, null, null, null);
                        FillRow(oDataDoc, true, 2, 3, grade, null, null, null, null, null);
                        TopData = false;

                        //*************************************************

                        //oDataDoc.Save();

                        //*************************************************
                    }


                    #region Region_01

                    if (IsFirstDay == true)
                    {
                        string[] var_1 = TextLines[0].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                        string[] var_2 = var_1[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                        string[] var_3 = var_2[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                        count2Txt = var_3[2];
                        count = Convert.ToInt32(count2Txt);
                        IsFirstDay = false;
                    }
                    else
                    {
                        count = 1;
                    }
                    
                    while (count <= 31)
                    {
                        if (count.ToString().Length == 1)
                        {
                            count2Txt = "0" + count.ToString();
                        }
                        else
                        {
                            count2Txt = count.ToString();
                        }


                        FDate = SubSearch + "-" + count2Txt; // Format: 2020-06-01
                        try
                        {
                            ThisDay = new DateTime(Convert.ToInt16(SelYear), MO, count);
                        }
                        catch (Exception)
                        {
                            break;
                        }

                        string _Day = ThisDay.DayOfWeek.ToString();
                        if (_Day.Substring(0, 3) == "Sat" || _Day.Substring(0, 3) == "Sun")
                        {
                            SatOrSun = true;
                            TotalSat_Sun++;
                            FillRow(oDataDoc, false, 3, lastRow, null, null, null, null, null, null);
                        }

                        //FDate = "296	2020-06-01 07:23:40	1	0	1	0";
                        DataLines = TextLines.FindAll(s => s.Contains(FDate));
                        if (DataLines.Count == 0)
                        {
                            SatOrSun = false;
                            string day = ThisDay.DayOfWeek.ToString();
                            FillRow(oDataDoc, false, 3, lastRow, ThisDay.ToShortDateString() + " (" + day.Substring(0, 3) + ") ", null, null, null, null, null);
                            if (_Day.Substring(0, 3) != "Sat" && _Day.Substring(0, 3) != "Sun")
                            {
                                //TotalHolidays++;
                            }

                            count++;
                            
                            lastRow++;
                            continue;
                        }

                        if (DataLines.Count == 1)
                        {
                                SatOrSun = false;
                                MissedPunche = true;
                                
                                string[] var1 = DataLines[0].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                //string[] var2 = DataLines[DataLines.Count - 1].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                string[] var2 = var1[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                //string[] var3 = var2[0].Split(Del3, Wordcount, StringSplitOptions.RemoveEmptyEntries); 
                                
                                Date_1 = var2[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries); //Date_1 > 01 , 06 , 2020
                                Time_1 = var2[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries); //Time_1 > 07 , 42 , 10

                                //Date_2 = var4[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                                //Time_2 = var4[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries);

                                int y1 = Convert.ToInt32(Date_1[0]);
                                int mo1 = Convert.ToInt32(Date_1[1]);
                                int d1 = Convert.ToInt32(Date_1[2]);
                                int h1 = Convert.ToInt32(Time_1[0]);
                                int mi1 = Convert.ToInt32(Time_1[1]);
                                int s1 = Convert.ToInt32(Time_1[2].Substring(0, 2));

                                date1 = new System.DateTime(y1, mo1, d1, h1, mi1, s1);
                                string day = ThisDay.DayOfWeek.ToString();
                                FillRow(oDataDoc, false, 3, lastRow, date1.ToShortDateString() + " (" + day.Substring(0, 3) + ") ", date1.TimeOfDay.ToString(), null, null, null, null);
                                TotalWorkDays++;
                                lastRow++;
                                DataLines.Clear();
                                count++;
                        }
                        else
                        {
                            string[] var1 = DataLines[0].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                            string[] var2 = DataLines[DataLines.Count - 1].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                            string[] var3 = var1[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                            string[] var4 = var2[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                            Date_1 = var3[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                            Time_1 = var3[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries);
                            Date_2 = var4[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                            Time_2 = var4[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries);

                            int y1 = Convert.ToInt32(Date_1[0]);
                            int mo1 = Convert.ToInt32(Date_1[1]);
                            int d1 = Convert.ToInt32(Date_1[2]);
                            int h1 = Convert.ToInt32(Time_1[0]);
                            int mi1 = Convert.ToInt32(Time_1[1]);
                            int s1 = Convert.ToInt32(Time_1[2].Substring(0, 2));

                            int y2 = Convert.ToInt32(Date_2[0]);
                            int mo2 = Convert.ToInt32(Date_2[1]);
                            int d2 = Convert.ToInt32(Date_2[2]);
                            int h2 = Convert.ToInt32(Time_2[0]);
                            int mi2 = Convert.ToInt32(Time_2[1]);
                            int s2 = Convert.ToInt32(Time_2[2].Substring(0, 2));

                            date1 = new System.DateTime(y1, mo1, d1, h1, mi1, s1);
                            date2 = new System.DateTime(y2, mo2, d2, h2, mi2, s2);
                            Diff = date2 - date1;
                            

                            // Calculate Status

                            #region Executive&Clerical

                            if (grade == "Executive Grade" || grade == "Allied & Clerical Grade")
                            {

                                DateTime EandC_N = new DateTime(y1, mo1, d1, 08, 00, 59);
                                DateTime EandC_N_0 = new DateTime(y1, mo1, d1, 08, 15, 59);
                                DateTime EandC_N_OT = new DateTime(y1, mo1, d1, 08, 15, 00);
                                DateTime EandC_N_1 = new DateTime(y2, mo2, d2, 16, 15, 00);
                                DateTime EandC_N_ = new DateTime(y1, mo1, d1, 08, 30, 00);
                                DateTime EandC_SL = new DateTime(y1, mo1, d1, 08, 30, 59);
                                DateTime EandC_SL_0 = new DateTime(y1, mo1, d1, 09, 30, 59);
                                DateTime EandC_SL_1 = new DateTime(y1, mo1, d1, 14, 45, 00);
                                DateTime EandC_HD = new DateTime(y1, mo1, d1, 12, 15, 59);
                                DateTime OT_Lunch = new DateTime(y1, mo1, d1, 13, 00, 00);

                                int result0 = date1.CompareTo(EandC_N);
                                int result1 = date1.CompareTo(EandC_N_0);
                                int result2 = date2.CompareTo(EandC_N_1);
                                int result5 = date1.CompareTo(EandC_HD);
                                int result6 = date2.CompareTo(EandC_SL_1);

                                if (SatOrSun)
                                {
                                    status = "OT";
                                    
                                    int result7 = date2.CompareTo(EandC_N_);
                                    int result8 = date1.CompareTo(EandC_N_);
                                    int result9 = date2.CompareTo(OT_Lunch);
                                    if (result7 > 0)
                                    {
                                        if (result8 <= 0)
                                        {
                                            if (result9 > 0)
                                            {
                                                LunchDiff = date2 - OT_Lunch;
                                                if (LunchDiff >= new TimeSpan(1,0,0))
                                                {
                                                    OT = date2 - EandC_N_;
                                                    OT -= new TimeSpan(1, 0, 0);
                                                    OT = updateOT(OT);
                                                }
                                                else
                                                {
                                                    OT = date2 - EandC_N_;
                                                    OT -= LunchDiff;
                                                    OT = updateOT(OT);
                                                }
                                            }
                                            else
                                            {
                                                OT = date2 - EandC_N_;
                                                OT = updateOT(OT);
                                            }
                                        }
                                        else
                                        {
                                            
                                            if (result9 > 0)
                                            {
                                                LunchDiff = date2 - OT_Lunch;
                                                if (LunchDiff >= new TimeSpan(1, 0, 0))
                                                {
                                                    OT = date2 - date1;
                                                    OT -= new TimeSpan(1, 0, 0);
                                                    OT = updateOT(OT);
                                                }
                                                else
                                                {
                                                    OT = date2 - date1;
                                                    OT -= LunchDiff;
                                                    OT = updateOT(OT);
                                                }
                                            }
                                            else
                                            {
                                                OT = date2 - date1;
                                                OT = updateOT(OT);
                                            }
                                        }
                                        
                                    }
                                    else
                                    {
                                        OT = new TimeSpan(00, 00, 00);
                                    }
                                    SatOrSun = false;
                                    //break;
                                }
                                else
                                {
                                    if (result1 <= 0 && result2 >= 0) //on,before 8.15.59 >> on,after 4.15.00
                                    {
                                        status = "1";
                                        if (result2 > 0 && result0 <= 0) //on,before 8.00.59 >> after 4.30.00
                                        {
                                            OT = date2 - EandC_N_1;
                                            OT = updateOT(OT);
                                        }
                                        else //after 7.30.59 >> Deduct 15 min.
                                        {
                                            OT = date2 - EandC_N_1;
                                            OT -= new TimeSpan(0, 15, 0);
                                            OT = updateOT(OT);
                                        }
                                    }
                                    else
                                    {
                                        if (result1 > 0 && result2 >= 0) //after 8.15.59 >> on,after 4.15.00
                                        {
                                            int result3 = date1.CompareTo(EandC_SL); //#8.30.59
                                            int result4 = date2.CompareTo(EandC_N_1);
                                            if (result3 <= 0 && result4 >= 0) //on,before 8.30.59 >> on,after 4.15.00 (unnecess...)
                                            {
                                                diff1 = date1 - EandC_N_0;
                                                diff2 = date2 - EandC_N_1;
                                                if (diff1 <= diff2)
                                                {
                                                    status = "Late (C)";
                                                    TotalLateDays++;
                                                }
                                                else
                                                {
                                                    status = "Late";
                                                    TotalLateDays++;
                                                    //OT = new TimeSpan(00, 00, 00);
                                                }
                                            }
                                            else
                                            {
                                                if (result3 > 0 && result5 <= 0) //after 8.30.59 >> on,before 12.15.59
                                                {
                                                    status = "Short Leave";
                                                    TotalShortLeaves++;
                                                }

                                                if (result5 > 0 && result4 >= 0) //after 12.15.59 >> on,after 4.15.00
                                                {
                                                    status = "Half Day";
                                                }
                                            }

                                            //break;
                                        }
                                        else
                                        {
                                            if (result1 <= 0 && result2 < 0) //on,before 8.15.59 >> before 4.15.00
                                            {
                                                if (result6 >= 0) //on,after 2.45.59
                                                {
                                                    status = "Short Leave";
                                                    TotalShortLeaves++;
                                                }
                                                else
                                                {
                                                    status = "Half Day";
                                                }
                                            }
                                            else
                                            {
                                                //int result2 = date1.CompareTo(EandC_SL); //#8.30.59
                                                if (result1 > 0 && result2 < 0) //after 8.15.59 >> before 4.15.00
                                                {
                                                    status = "Half Day";
                                                    //if (result6 >= 0) //on,after 2.45.59
                                                    //{
                                                    //    status = "Short Leave";
                                                    //    TotalShortLeaves++;
                                                    //}
                                                    //else
                                                    //{
                                                    //    status = "Half Day";
                                                    //}
                                                }
                                            }
                                        }

                                        //OT calculation
                                        if (result2 > 0)
                                        {
                                            if (status == "Late (C)")
                                            {
                                                OT = date2 - EandC_N_1;
                                                diff1 = date1 - EandC_N_OT;
                                                OT -= diff1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                            }
                                            if (status == "Short Leave")
                                            {
                                                OT = date2 - EandC_N_1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                                //diff1 = date1 - EandC_N_OT;
                                                //OT -= diff1;
                                            }
                                            if (status == "Half Day")
                                            {
                                                OT = date2 - EandC_N_1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                                //diff1 = date1 - EandC_N_OT;
                                                //OT -= diff1;
                                            }

                                        }
                                        else
                                        {
                                            OT = new TimeSpan(00, 00, 00);
                                        }
                                        if (result2 == 0)
                                        {
                                            OT = new TimeSpan(00, 00, 00);
                                        }

                                        #region intelligentCheck

                                        //intelligent check
                                        int tempComp1 = Diff.CompareTo(new TimeSpan(7, 59, 0));
                                        int tempComp2 = Diff.CompareTo(new TimeSpan(3, 59, 0));
                                        if (status == "1" && tempComp1 < 0)
                                        {
                                            status = "#CHECK!";
                                        }
                                        if (status == "Half Day" && tempComp2 < 0)
                                        {
                                            status = "#CHECK!";
                                        }

                                        #endregion


                                        //if (result5 > 0)
                                        //{
                                        //    int result3 = date1.CompareTo(EandC_SL);
                                        //    if (result1 > 0)
                                        //    {
                                        //        status = "Short Leave";
                                        //    }
                                        //    else
                                        //    {
                                        //        diff1 = date1 - EandC_N_0;
                                        //        diff2 = date2 - EandC_N_1;
                                        //        if (diff1 <= diff2)
                                        //        {
                                        //            status = "Late (C)";
                                        //        }
                                        //    }

                                        //}
                                        //else
                                        //{
                                        //    status = "Late";
                                        //}
                                        //status = "0";
                                    } 
                                }
                            } 
                            #endregion
                            
                            // CAUTION! Comments are not corrected according to the grades! (From this point forward)

                            #region Technical

                            if (grade == "Technical Grade")
                            {
                                DateTime EandC_N = new DateTime(y1, mo1, d1, 07, 45, 59);
                                DateTime EandC_N_0 = new DateTime(y1, mo1, d1, 08, 00, 59);
                                DateTime EandC_N_OT = new DateTime(y1, mo1, d1, 08, 00, 00);
                                DateTime EandC_N_1 = new DateTime(y2, mo2, d2, 16, 30, 00);
                                DateTime EandC_N_ = new DateTime(y1, mo1, d1, 08, 30, 00);
                                DateTime EandC_SL = new DateTime(y1, mo1, d1, 08, 15, 59);
                                DateTime EandC_SL_0 = new DateTime(y1, mo1, d1, 09, 15, 59);
                                DateTime EandC_SL_1 = new DateTime(y1, mo1, d1, 15, 00, 00);
                                DateTime EandC_HD = new DateTime(y1, mo1, d1, 12, 15, 59);
                                DateTime OT_Lunch = new DateTime(y1, mo1, d1, 13, 00, 59);

                                int result0 = date1.CompareTo(EandC_N);
                                int result1 = date1.CompareTo(EandC_N_0);
                                int result2 = date2.CompareTo(EandC_N_1);
                                int result5 = date1.CompareTo(EandC_HD);
                                int result6 = date2.CompareTo(EandC_SL_1);

                                if (SatOrSun == true)
                                {
                                    status = "OT";

                                    int result7 = date2.CompareTo(EandC_N_);
                                    int result8 = date1.CompareTo(EandC_N_);
                                    int result9 = date2.CompareTo(OT_Lunch);
                                    if (result7 > 0)
                                    {
                                        if (result8 <= 0)
                                        {
                                            if (result9 > 0)
                                            {
                                                LunchDiff = date2 - OT_Lunch;
                                                if (LunchDiff >= new TimeSpan(1, 0, 0))
                                                {
                                                    OT = date2 - EandC_N_;
                                                    OT -= new TimeSpan(1, 0, 0);
                                                    OT = updateOT(OT);
                                                }
                                                else
                                                {
                                                    OT = date2 - EandC_N_;
                                                    OT -= LunchDiff;
                                                    OT = updateOT(OT);
                                                }
                                            }
                                            else
                                            {
                                                OT = date2 - EandC_N_;
                                                OT = updateOT(OT);
                                            }
                                        }
                                        else
                                        {

                                            if (result9 > 0)
                                            {
                                                LunchDiff = date2 - OT_Lunch;
                                                if (LunchDiff >= new TimeSpan(1, 0, 0))
                                                {
                                                    OT = date2 - date1;
                                                    OT -= new TimeSpan(1, 0, 0);
                                                    OT = updateOT(OT);
                                                }
                                                else
                                                {
                                                    OT = date2 - date1;
                                                    OT -= LunchDiff;
                                                    OT = updateOT(OT);
                                                }
                                            }
                                            else
                                            {
                                                OT = date2 - date1;
                                                OT = updateOT(OT);
                                            }
                                        }

                                    }
                                    else
                                    {
                                        OT = new TimeSpan(00, 00, 00);
                                    }
                                    SatOrSun = false;
                                    //break;
                                }

                                else
                                {
                                    if (result1 <= 0 && result2 >= 0) //on,before 8.00.59 >> on,after 4.30.00
                                    {
                                        status = "1";
                                        if (result2 > 0 && result0 <= 0) //on,before 7.45.59 >> after 4.30.00
                                        {
                                            OT = date2 - EandC_N_1;
                                            OT = updateOT(OT);
                                        }
                                        else //after 7.45.59 >> Deduct 15 min.
                                        {
                                            OT = date2 - EandC_N_1;
                                            OT -= new TimeSpan(0, 15, 0);
                                            OT = updateOT(OT);
                                        }
                                    }
                                    else
                                    {
                                        if (result1 > 0 && result2 >= 0) //after 8.00.59 >> on,after 4.30.00
                                        {
                                            int result3 = date1.CompareTo(EandC_SL); //#8.15.59
                                            int result4 = date2.CompareTo(EandC_N_1);
                                            if (result3 <= 0 && result4 >= 0) //on,before 8.15.59 >> on,after 4.30.00 (unnecess...)
                                            {
                                                diff1 = date1 - EandC_N_0;
                                                diff2 = date2 - EandC_N_1;
                                                if (diff1 <= diff2)
                                                {
                                                    status = "Late (C)";
                                                    TotalLateDays++;
                                                }
                                                else
                                                {
                                                    status = "Late";
                                                    TotalLateDays++;
                                                    //OT = new TimeSpan(00, 00, 00);
                                                }
                                            }
                                            else
                                            {
                                                if (result3 > 0 && result5 <= 0) //after 8.15.59 >> on,before 12.15.59
                                                {
                                                    status = "Short Leave";
                                                    TotalShortLeaves++;
                                                }

                                                if (result5 > 0 && result4 >= 0) //after 12.15.59 >> on,after 4.30.00
                                                {
                                                    status = "Half Day";
                                                }
                                            }

                                            //break;
                                        }
                                        else
                                        {
                                            if (result1 <= 0 && result2 < 0) //on,before 8.00.59 >> before 4.30.00
                                            {
                                                if (result6 >= 0) //on,after 3.00.59
                                                {
                                                    status = "Short Leave";
                                                    TotalShortLeaves++;
                                                }
                                                else
                                                {
                                                    status = "Half Day";
                                                }
                                            }
                                            else
                                            {
                                                //int result2 = date1.CompareTo(EandC_SL); //#8.30.59
                                                if (result1 > 0 && result2 < 0) //after 8.15.59 >> before 4.15.00
                                                {
                                                    status = "Half Day";
                                                    //if (result6 >= 0) //on,after 2.45.59
                                                    //{
                                                    //    status = "Short Leave";
                                                    //    TotalShortLeaves++;
                                                    //}
                                                    //else
                                                    //{
                                                    //    status = "Half Day";
                                                    //}
                                                }
                                            }
                                        }

                                        //OT calculation
                                        if (result2 > 0)
                                        {
                                            if (status == "Late (C)")
                                            {
                                                OT = date2 - EandC_N_1;
                                                diff1 = date1 - EandC_N_OT;
                                                OT -= diff1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                            }
                                            if (status == "Short Leave")
                                            {
                                                OT = date2 - EandC_N_1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                                //diff1 = date1 - EandC_N_OT;
                                                //OT -= diff1;
                                            }
                                            if (status == "Half Day")
                                            {
                                                OT = date2 - EandC_N_1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                                //diff1 = date1 - EandC_N_OT;
                                                //OT -= diff1;
                                            }

                                        }
                                        else
                                        {
                                            OT = new TimeSpan(00, 00, 00);
                                        }
                                        if (result2 == 0)
                                        {
                                            OT = new TimeSpan(00, 00, 00);
                                        }

                                        #region intelligentCheck

                                        //intelligent check
                                        int tempComp = Diff.CompareTo(new TimeSpan(7, 59, 0));
                                        int tempComp2 = Diff.CompareTo(new TimeSpan(3, 59, 0));
                                        if (status == "1" && tempComp < 0)
                                        {
                                            status = "#CHECK!";
                                        }
                                        if (status == "Half Day" && tempComp2 < 0)
                                        {
                                            status = "#CHECK!";
                                        }

                                        #endregion

                                    }
                                }
                                
                                #region OldCode
                                //DateTime EandC_N_0 = new DateTime(y1, mo1, d1, 08, 00, 59);
                                //DateTime EandC_N_1 = new DateTime(y2, mo2, d2, 16, 30, 00);
                                //DateTime EandC_SL = new DateTime(y1, mo1, d1, 08, 15, 59);
                                //DateTime EandC_SL_0 = new DateTime(y1, mo1, d1, 09, 15, 59);
                                //DateTime EandC_SL_1 = new DateTime(y1, mo1, d1, 15, 00, 59);
                                //DateTime EandC_HD = new DateTime(y1, mo1, d1, 12, 15, 59);

                                //int result1 = date1.CompareTo(EandC_N_0);
                                //int result2 = date2.CompareTo(EandC_N_1);
                                //int result5 = date1.CompareTo(EandC_HD);
                                //int result6 = date2.CompareTo(EandC_SL_1);

                                //if (result1 <= 0 && result2 >= 0)
                                //{
                                //    status = "1";
                                //}
                                //else
                                //{
                                //    if (result1 > 0 && result2 >= 0)
                                //    {
                                //        int result3 = date1.CompareTo(EandC_SL);
                                //        int result4 = date2.CompareTo(EandC_N_1);
                                //        if (result3 <= 0 && result4 >= 0)
                                //        {
                                //            diff1 = date1 - EandC_N_0;
                                //            diff2 = date2 - EandC_N_1;
                                //            if (diff1 <= diff2)
                                //            {
                                //                status = "Late (C)";
                                //                TotalLateDays++;
                                //            }
                                //            else
                                //            {
                                //                status = "Late";
                                //                TotalLateDays++;
                                //            }
                                //        }
                                //        else
                                //        {
                                //            if (result3 >= 0 && result5 <= 0)
                                //            {
                                //                status = "Short Leave";
                                //                TotalShortLeaves++;
                                //            }

                                //            if (result5 >= 0 && result4 <= 0)
                                //            {
                                //                status = "Half Day";
                                //            }
                                //        }

                                //        //break;
                                //    }
                                //    else
                                //    {
                                //        if (result1 <= 0 && result2 < 0)
                                //        {
                                //            if (result6 >= 0)
                                //            {
                                //                status = "Short Leave";
                                //                TotalShortLeaves++;
                                //            }
                                //            else
                                //            {
                                //                status = "Half Day";
                                //            }
                                //        }
                                //    }

                                //    //if (result5 > 0)
                                //    //{
                                //    //    int result3 = date1.CompareTo(EandC_SL);
                                //    //    if (result1 > 0)
                                //    //    {
                                //    //        status = "Short Leave";
                                //    //    }
                                //    //    else
                                //    //    {
                                //    //        diff1 = date1 - EandC_N_0;
                                //    //        diff2 = date2 - EandC_N_1;
                                //    //        if (diff1 <= diff2)
                                //    //        {
                                //    //            status = "Late (C)";
                                //    //        }
                                //    //    }

                                //    //}
                                //    //else
                                //    //{
                                //    //    status = "Late";
                                //    //}
                                //    //status = "0";
                                //} 
                                #endregion
                            } 
                            #endregion

                            #region Primary
                            if (grade == "Primary Grade")
                            {
                                DateTime EandC_N = new DateTime(y1, mo1, d1, 07, 30, 59);
                                DateTime EandC_N_0 = new DateTime(y1, mo1, d1, 07, 45, 59);
                                DateTime EandC_N_OT = new DateTime(y1, mo1, d1, 07, 45, 00);
                                DateTime EandC_N_1 = new DateTime(y2, mo2, d2, 16, 30, 00);
                                DateTime EandC_N_ = new DateTime(y1, mo1, d1, 08, 30, 00);
                                DateTime EandC_SL = new DateTime(y1, mo1, d1, 08, 00, 59);
                                DateTime EandC_SL_0 = new DateTime(y1, mo1, d1, 09, 00, 59);
                                DateTime EandC_SL_1 = new DateTime(y1, mo1, d1, 15, 00, 00);
                                DateTime EandC_HD = new DateTime(y1, mo1, d1, 12, 00, 59);
                                DateTime OT_Lunch = new DateTime(y1, mo1, d1, 13, 00, 59);

                                int result0 = date1.CompareTo(EandC_N);
                                int result1 = date1.CompareTo(EandC_N_0);
                                int result2 = date2.CompareTo(EandC_N_1);
                                int result5 = date1.CompareTo(EandC_HD);
                                int result6 = date2.CompareTo(EandC_SL_1);

                                if (SatOrSun == true)
                                {
                                    status = "OT";

                                    int result7 = date2.CompareTo(EandC_N_);
                                    int result8 = date1.CompareTo(EandC_N_);
                                    int result9 = date2.CompareTo(OT_Lunch);
                                    if (result7 > 0)
                                    {
                                        if (result8 <= 0)
                                        {
                                            if (result9 > 0)
                                            {
                                                LunchDiff = date2 - OT_Lunch;
                                                if (LunchDiff >= new TimeSpan(1, 0, 0))
                                                {
                                                    OT = date2 - EandC_N_;
                                                    OT -= new TimeSpan(1, 0, 0);
                                                    OT = updateOT(OT);
                                                }
                                                else
                                                {
                                                    OT = date2 - EandC_N_;
                                                    OT -= LunchDiff;
                                                    OT = updateOT(OT);
                                                }
                                            }
                                            else
                                            {
                                                OT = date2 - EandC_N_;
                                                OT = updateOT(OT);
                                            }
                                        }
                                        else
                                        {

                                            if (result9 > 0)
                                            {
                                                LunchDiff = date2 - OT_Lunch;
                                                if (LunchDiff >= new TimeSpan(1, 0, 0))
                                                {
                                                    OT = date2 - date1;
                                                    OT -= new TimeSpan(1, 0, 0);
                                                    OT = updateOT(OT);
                                                }
                                                else
                                                {
                                                    OT = date2 - date1;
                                                    OT -= LunchDiff;
                                                    OT = updateOT(OT);
                                                }
                                            }
                                            else
                                            {
                                                OT = date2 - date1;
                                                OT = updateOT(OT);
                                            }
                                        }

                                    }
                                    else
                                    {
                                        OT = new TimeSpan(00, 00, 00);
                                    }
                                    SatOrSun = false;
                                    //break;
                                }

                                else
                                {
                                    if (result1 <= 0 && result2 >= 0) //on,before 7.45.59 >> on,after 4.30.00
                                    {
                                        status = "1";
                                        if (result2 > 0 && result0 <= 0) //on,before 7.30.59 >> after 4.30.00
                                        {
                                            OT = date2 - EandC_N_1;
                                            OT = updateOT(OT);
                                        }
                                        else //after 7.30.59 >> Deduct 15 min.
                                        {
                                            OT = date2 - EandC_N_1;
                                            OT -= new TimeSpan(0, 15, 0);
                                            OT = updateOT(OT);
                                        }
                                    }
                                    else
                                    {
                                        if (result1 > 0 && result2 >= 0) //after 7.45.59 >> on,after 4.30.00
                                        {
                                            int result3 = date1.CompareTo(EandC_SL); //#8.15.59
                                            int result4 = date2.CompareTo(EandC_N_1);
                                            if (result3 <= 0 && result4 >= 0) //on,before 8.15.59 >> on,after 4.30.00 (unnecess...)
                                            {
                                                diff1 = date1 - EandC_N_0;
                                                diff2 = date2 - EandC_N_1;
                                                if (diff1 <= diff2)
                                                {
                                                    status = "Late (C)";
                                                    TotalLateDays++;
                                                }
                                                else
                                                {
                                                    status = "Late";
                                                    TotalLateDays++;
                                                    //OT = new TimeSpan(00, 00, 00);
                                                }
                                            }
                                            else
                                            {
                                                if (result3 > 0 && result5 <= 0) //after 8.15.59 >> on,before 12.15.59
                                                {
                                                    status = "Short Leave";
                                                    TotalShortLeaves++;
                                                }

                                                if (result5 > 0 && result4 >= 0) //after 12.15.59 >> on,after 4.30.00
                                                {
                                                    status = "Half Day";
                                                }
                                            }

                                            //break;
                                        }
                                        else
                                        {
                                            if (result1 <= 0 && result2 < 0) //on,before 8.00.59 >> before 4.30.00
                                            {
                                                if (result6 >= 0) //on,after 3.00.59
                                                {
                                                    status = "Short Leave";
                                                    TotalShortLeaves++;
                                                }
                                                else
                                                {
                                                    status = "Half Day";
                                                }
                                            }
                                            else
                                            {
                                                //int result2 = date1.CompareTo(EandC_SL); //#8.30.59
                                                if (result1 > 0 && result2 < 0) //after 8.15.59 >> before 4.15.00
                                                {
                                                    status = "Half Day";
                                                    //if (result6 >= 0) //on,after 2.45.59
                                                    //{
                                                    //    status = "Short Leave";
                                                    //    TotalShortLeaves++;
                                                    //}
                                                    //else
                                                    //{
                                                    //    status = "Half Day";
                                                    //}
                                                }
                                            }
                                        }

                                        //OT calculation
                                        if (result2 > 0)
                                        {
                                            if (status == "Late (C)")
                                            {
                                                OT = date2 - EandC_N_1;
                                                diff1 = date1 - EandC_N_OT;
                                                OT -= diff1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                            }
                                            if (status == "Short Leave")
                                            {
                                                OT = date2 - EandC_N_1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                                //diff1 = date1 - EandC_N_OT;
                                                //OT -= diff1;
                                            }
                                            if (status == "Half Day")
                                            {
                                                OT = date2 - EandC_N_1;
                                                OT = updateOT(OT);
                                                //TotalOT += OT;
                                                //diff1 = date1 - EandC_N_OT;
                                                //OT -= diff1;
                                            }

                                        }
                                        else
                                        {
                                            OT = new TimeSpan(00, 00, 00);
                                        }
                                        if (result2 == 0)
                                        {
                                            OT = new TimeSpan(00, 00, 00);
                                        }

                                        #region intelligentCheck

                                        //intelligent check
                                        int tempComp = Diff.CompareTo(new TimeSpan(7, 59, 0));
                                        int tempComp2 = Diff.CompareTo(new TimeSpan(3, 59, 0));
                                        if (status == "1" && tempComp < 0)
                                        {
                                            status = "#CHECK!";
                                        }
                                        if (status == "Half Day" && tempComp2 < 0)
                                        {
                                            status = "#CHECK!";
                                        }

                                        #endregion

                                    }
                                }
                                #region OldCode_
                                //DateTime EandC_N_0 = new DateTime(y1, mo1, d1, 07, 45, 59);
                                //DateTime EandC_N_1 = new DateTime(y2, mo2, d2, 16, 30, 00);
                                //DateTime EandC_SL = new DateTime(y1, mo1, d1, 08, 00, 59);
                                //DateTime EandC_SL_0 = new DateTime(y1, mo1, d1, 09, 00, 59);
                                //DateTime EandC_SL_1 = new DateTime(y1, mo1, d1, 15, 00, 59);
                                //DateTime EandC_HD = new DateTime(y1, mo1, d1, 12, 00, 59);

                                //int result1 = date1.CompareTo(EandC_N_0);
                                //int result2 = date2.CompareTo(EandC_N_1);
                                //int result5 = date1.CompareTo(EandC_HD);
                                //int result6 = date2.CompareTo(EandC_SL_1);

                                //if (result1 <= 0 && result2 >= 0)
                                //{
                                //    status = "1";
                                //}
                                //else
                                //{
                                //    if (result1 > 0 && result2 >= 0)
                                //    {
                                //        int result3 = date1.CompareTo(EandC_SL);
                                //        int result4 = date2.CompareTo(EandC_N_1);
                                //        if (result3 <= 0 && result4 >= 0)
                                //        {
                                //            diff1 = date1 - EandC_N_0;
                                //            diff2 = date2 - EandC_N_1;
                                //            if (diff1 <= diff2)
                                //            {
                                //                status = "Late (C)";
                                //                TotalLateDays++;
                                //            }
                                //            else
                                //            {
                                //                status = "Late";
                                //                TotalLateDays++;
                                //            }
                                //        }
                                //        else
                                //        {
                                //            if (result3 >= 0 && result5 <= 0)
                                //            {
                                //                status = "Short Leave";
                                //                TotalShortLeaves++;
                                //            }

                                //            if (result5 >= 0 && result4 <= 0)
                                //            {
                                //                status = "Half Day";
                                //            }
                                //        }

                                //        //break;
                                //    }
                                //    else
                                //    {
                                //        if (result1 <= 0 && result2 < 0)
                                //        {
                                //            if (result6 >= 0)
                                //            {
                                //                status = "Short Leave";
                                //                TotalShortLeaves++;
                                //            }
                                //            else
                                //            {
                                //                status = "Half Day";
                                //            }
                                //        }
                                //    }

                                //    //if (result5 > 0)
                                //    //{
                                //    //    int result3 = date1.CompareTo(EandC_SL);
                                //    //    if (result1 > 0)
                                //    //    {
                                //    //        status = "Short Leave";
                                //    //    }
                                //    //    else
                                //    //    {
                                //    //        diff1 = date1 - EandC_N_0;
                                //    //        diff2 = date2 - EandC_N_1;
                                //    //        if (diff1 <= diff2)
                                //    //        {
                                //    //            status = "Late (C)";
                                //    //        }
                                //    //    }

                                //    //}
                                //    //else
                                //    //{
                                //    //    status = "Late";
                                //    //}
                                //    //status = "0";
                                //}


                                ////**************************************************************

                                ////      Pre Code        //


                                ////DateTime EandC_N_0 = new DateTime(y1, mo1, d1, 08, 15, 59);
                                ////DateTime EandC_N_1 = new DateTime(y2, mo2, d2, 16, 15, 00);
                                ////int result1 = date1.CompareTo(EandC_N_0);
                                ////int result2 = date2.CompareTo(EandC_N_1);

                                ////if (result1 < 0 && result2 > 0)
                                ////{
                                ////    status = "1";
                                ////}
                                ////else
                                ////{
                                ////    status = "0";
                                ////} 
                                #endregion
                            } 
                            #endregion

                            #region Trainee
                            if (grade == "Trainee")
                            {
                                DateTime Trainee_0 = new DateTime(y1, mo1, d1, 08, 00, 59);
                                DateTime Trainee_1 = new DateTime(y2, mo2, d2, 16, 15, 00);

                                int result1 = date1.CompareTo(Trainee_0);
                                int result2 = date2.CompareTo(Trainee_1);

                                if (result1 <= 0 && result2 >= 0)
                                {
                                    status = "1";
                                }
                                else
                                {
                                    status = "Half Day";
                                }
                            } 
                            #endregion

                            string day = date1.DayOfWeek.ToString();
                            int timeComp = OT.CompareTo(new TimeSpan(1, 0, 0));
                            

                            if (timeComp < 0)
                            {
                                FillRow(oDataDoc, false, 3, lastRow, date1.ToShortDateString() + " (" + day.Substring(0, 3) + ") ", date1.TimeOfDay.ToString(), date2.TimeOfDay.ToString(), status, Diff.ToString(), null);
                            }
                            else
                            {
                                TotalOT += OT;
                                FillRow(oDataDoc, false, 3, lastRow, date1.ToShortDateString() + " (" + day.Substring(0, 3) + ") ", date1.TimeOfDay.ToString(), date2.TimeOfDay.ToString(), status, Diff.ToString(), OT.ToString());
                            }
                            //TotalDiff += Diff;
                            TotalWorkDays++;
                            lastRow++;
                            DataLines.Clear();
                            count++;
                        }


                    }
                    #endregion


                    //string LengthOfText = TextLines.Count.ToString();
                    
                }
                else
                {
                    MessageBox.Show("Oops! No data to generate a report!\nERROR:02X1791X", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }




        private void FillRow(Word._Document oDoc, bool isTop, int table, int Row, string Text1,

        string Text2, string Text3, string Text4, string Text5, string Text6)
        {
            if (bottomData)
            {
                oDoc.Tables[table].Cell(1, 3).Range.InsertAfter(Text1);
                oDoc.Tables[table].Cell(1, 6).Range.InsertAfter(Text2);
                oDoc.Tables[table].Cell(2, 3).Range.InsertAfter(Text3);
                oDoc.Tables[table].Cell(2, 6).Range.InsertAfter(Text4);
                oDoc.Save();
                bottomData = false;
                success = true;
                return;
            }

            if (isTop)
            {
                oDoc.Tables[table].Cell(Row, 3).Range.InsertAfter(Text1);
                TopData = false;
                return;
            }
            if (SatOrSun)
            {
                //Selection.Shading.BackgroundPatternColor = -603923969
                Word.WdColor backColor;
                backColor = (Word.WdColor)(-603923969);
                oDoc.Tables[table].Cell(lastRow, 1).Shading.BackgroundPatternColor = backColor;
                //SatOrSun = false;
                return;
            }
            if (MissedPunche)
            {
                oDoc.Tables[table].Cell(Row, 1).Range.InsertAfter(Text1);
                oDoc.Tables[table].Cell(Row, 2).Merge(oDoc.Tables[table].Cell(Row, 3));
                oDoc.Tables[table].Cell(Row, 2).Range.InsertAfter(Text2);
                oDoc.Tables[table].Cell(Row, 3).Range.InsertAfter("MISSED");
                oDoc.Tables[table].Cell(Row, 4).Range.InsertAfter("-");

                //**************************************************************

                //oDoc.Save();

                //**************************************************************

                MissedPunche = false;
                return;
            }


            oDoc.Tables[table].Cell(Row, 1).Range.InsertAfter(Text1);
            oDoc.Tables[table].Cell(Row, 2).Range.InsertAfter(Text2);
            oDoc.Tables[table].Cell(Row, 3).Range.InsertAfter(Text3);
            oDoc.Tables[table].Cell(Row, 4).Range.InsertAfter(Text4);
            oDoc.Tables[table].Cell(Row, 5).Range.InsertAfter(Text5);
            oDoc.Tables[table].Cell(Row, 6).Range.InsertAfter(Text6);
            
        }

        private void FillData(Word._Document oDoc, bool isTop, bool isDay, string Text1, string Text2, string Text3, string Text4) 
        {
            try
            {
                if (isTop)
                {
                    foreach (Microsoft.Office.Interop.Word.Section section in oDoc.Sections)
                    {
                        Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                        headerRange.Tables[1].Cell(1, 2).Range.InsertAfter(Text1 + ", " + Text2);
                        headerRange.Tables[1].Cell(1, 4).Range.InsertAfter(Text3);
                    }
                    return;
                }
                if (isDay)
                {
                    Microsoft.Office.Interop.Word.Paragraph para1 = oDoc.Content.Paragraphs.Add(ref oMissing);

                    para1.Range.Text = Text1;
                    //para1.Range.Bold
                    para1.Range.Underline = Word.WdUnderline.wdUnderlineSingle;
                    para1.Range.InsertParagraphAfter();
                    return;
                }

                Microsoft.Office.Interop.Word.Paragraph para = oDoc.Content.Paragraphs.Add(ref oMissing);
                //object styleHeading1 = "Heading 1";
                //para1.Range.set_Style(ref styleHeading1);
                para.Range.Underline = Word.WdUnderline.wdUnderlineNone;
                //para.Range.Text = Text1 + "\t\t" + Text2 + "\t\t" + Text3 + "\t\t" + Text4;
                para.Range.Text = Text1 + "\t" + Text2 + "\t" + Text3 + "\t" + Text4;
                para.Format.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                para.Format.SpaceBefore = 2;
                para.SpaceAfter = 2;
                para.Range.InsertParagraphAfter();
                oDataDoc.Save();
            }
            catch (Exception)
            {
                success = false;
                MessageBox.Show("Oops! Unexpected error occured while generating the report!.\nERROR:02X1901X", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                backgroundWorker3.CancelAsync();
            }
        }

        private void backgroundWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //textBox1.Text += "\r\n" + FillItem + "\n\r";
        }

        private void backgroundWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (fileError)
            {
                DataLines.Clear();
                TextLines.Clear();
                lastRow = 2;
                button3.Enabled = true;
                groupBox1.Enabled = true;
                panel1.Enabled = true;
                MessageBox.Show("We don't have permission to access into the file system.\nSometimes antivirus software cause this error.\nPlease allow this software in settings on this matter.\nERROR:02X1920X", "Failed to access file system.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TopData = true;
            DataLines.Clear();
            TextLines.Clear();
            lastRow = 2;
            button3.Enabled = true;
            groupBox1.Enabled = true;
            panel1.Enabled = true;
            bottomData = true;
            //TotalHolidays -= TotalSat_Sun;
            TotalOT = new TimeSpan(TotalOT.Hours, TotalOT.Minutes, TotalOT.Seconds);
            FillRow(oDataDoc, false, 4, 0, TotalWorkDays.ToString(), TotalLateDays.ToString(), TotalOT.ToString(), TotalShortLeaves.ToString(), null, null);
            toolStripProgressBar1.Visible = false;
            if (success == true)
            {
                toolStripStatusLabel1.Text = "Report successfully generated.";
            }
            else
            {
                toolStripStatusLabel1.Text = "Failed to generate the report.";
            }
            TotalWorkDays = 0;
            TotalSat_Sun = 0;
            TotalHolidays = 0;
            TotalLateDays = 0;
            TotalShortLeaves = 0;
            bottomData = false;
            TotalDiff = new TimeSpan(0);
        }

        private void statusStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            selEmp = comboBox1.SelectedIndex;
            name = comboBox1.SelectedItem.ToString();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || textBox1.Text == Environment.NewLine + "☹" + Environment.NewLine + "Sorry. No records Found!")
            {
                MessageBox.Show("Oops! No data to generate a report!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            TextLines.Clear();
            success = false;
            toolStripStatusLabel1.Text = "Creating report... ";
            toolStripProgressBar1.Visible = true;
            panel1.Enabled = false;
            
            foreach (var item in textBox1.Lines)
            {
                TextLines.Add(item);
                //textBox3.Text += item;
            }
            button3.Enabled = false;
            TotalWorkDays = 0;
            TotalSat_Sun = 0;
            TotalLateDays = 0;
            TotalShortLeaves = 0;
            TotalDiff = new TimeSpan(0);
            //ShowF4.Show();
            
            //Check & copy the report.docx file
            try
            {
                if (System.IO.File.Exists(ReportFilePath))
                {
                    System.IO.Directory.CreateDirectory(Application.LocalUserAppDataPath + @"\ApplicationData\Reports");
                    ReportPath = Application.LocalUserAppDataPath + @"\ApplicationData\Reports\Report_" + name + "_" + SelYear + "_" + SelMonth + "_" + System.DateTime.Now.ToString().Replace(":","_") + ".docx";
                    System.IO.File.Copy(ReportFilePath, ReportPath, true);
                }
                else
                {
                    toolStripProgressBar1.Visible = false;
                    toolStripStatusLabel1.Text = "Failed to generate the report.";
                    MessageBox.Show("Oops! File missing error occured.\nPlease re-install the software or contact the developer with error code.\nERROR:02X2002X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Exit();
                }
            }
            catch (UnauthorizedAccessException)
            {
                toolStripProgressBar1.Visible = false;
                toolStripStatusLabel1.Text = "Failed to generate the report.";
                MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X2010X", "Elevated Privileges Required!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                fileError = true;
                Close();
            }
            catch (DirectoryNotFoundException ex)
            {
                fileError = true;
                toolStripProgressBar1.Visible = false;
                MessageBox.Show("ERROR:02X2018X " + ex.Message);
                Close();
            }

            backgroundWorker2.RunWorkerAsync(comboBox2.Items[comboBox2.SelectedIndex].ToString() + ", " + comboBox3.Items[comboBox3.SelectedIndex].ToString());
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            int a = comboBox3.SelectedIndex;
            SelYear = comboBox3.Items[a].ToString();
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            textBox1.Text = "";
            SelMonthText = (string)comboBox2.SelectedItem;
            MO = comboBox2.SelectedIndex;
            MO++;
            SelMonth = MO.ToString();
            if (MO.ToString().Length != 2)
            {
                SelMonth = "0" + MO.ToString();
            }
        }

        private void comboBox1_TextChanged(object sender, EventArgs e)
        {
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox4.SelectedItem == null || comboBox3.SelectedItem == null || comboBox2.SelectedItem == null)
            {
                MessageBox.Show("Please fill all fields including Year, Month and Department", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else
            {
                toolStripStatusLabel1.Text = "Creating report... ";
                toolStripProgressBar1.Visible = true;
                panel1.Enabled = false;
                selDept = (string)comboBox4.SelectedItem;
                try
                {
                    if (System.IO.File.Exists(FullReportPath))
                    {
                        System.IO.Directory.CreateDirectory(Application.LocalUserAppDataPath + @"\ApplicationData\Reports");
                        ReportPath = Application.LocalUserAppDataPath + @"\ApplicationData\Reports\Full_Report_" + SelYear + "_" + SelMonth + "_" + System.DateTime.Now.ToString().Replace(":", "_") + ".docx";
                        System.IO.File.Copy(FullReportPath, ReportPath, true);
                    }
                    else
                    {
                        toolStripProgressBar1.Visible = false;
                        toolStripStatusLabel1.Text = "Failed to generate the report.";
                        MessageBox.Show("Oops! File missing error occured.\nPlease re-install the software or contact the developer with the error code.\nERROR:02X2073X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Application.Exit();
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    toolStripProgressBar1.Visible = false;
                    toolStripStatusLabel1.Text = "Failed to generate the report.";
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X2081X", "Elevated Privileges Required!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    fileError = true;
                    Close();
                }
                catch (DirectoryNotFoundException ex)
                {
                    fileError = true;
                    toolStripProgressBar1.Visible = false;
                    MessageBox.Show("ERROR:02X2089X - " + ex.Message);
                    Close();
                }
                backgroundWorker3.RunWorkerAsync();
            }
        }

        private void backgroundWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;

            if (worker.CancellationPending == true)
            {
                e.Cancel = true;
                return;
            }
            else
            {
                List<string> IDs = new List<string>();
                List<string> EmpNames = new List<string>();
                String[] spearator = { "\t" };
                Int32 Wordcount = 4;
                bool ErrorStat = false;
                string[] lines = { };

                if (System.IO.File.Exists(FilePath))
                {
                    try
                    {
                        lines = System.IO.File.ReadAllLines(FilePath);
                        lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();

                        if (lines.Length != 0)
                        {
                            Array.Sort(lines, System.StringComparer.InvariantCulture);
                        }

                        foreach (var item in lines)
                        {
                            try
                            {
                                string[] data = item.Split(spearator, Wordcount, StringSplitOptions.RemoveEmptyEntries);

                                if (data[3] == selDept)
                                {
                                    IDs.Add(data[1]);
                                    EmpNames.Add(data[0]);
                                }
                                else
                                {
                                    continue;
                                }
                            }
                            catch (Exception)
                            {
                                MessageBox.Show("Unable to load employee data! Please restart the programme.\nIf this error persists reset employee data through settings section.\nERROR:02X2149X", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                ErrorStat = true;
                                break;
                            }

                        }
                        if (ErrorStat == true)
                        {
                            backgroundWorker3.CancelAsync();
                        }

                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X2165X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        fileError = true;
                        backgroundWorker3.CancelAsync();
                        return;
                    }
                }
                else
                {
                    MessageBox.Show("Oops! Employee name list is empty. You should add users first!\n\nIf you have a backup file, restore it from settings section.\nERROR:02X2173X", "No Employees", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    backgroundWorker3.CancelAsync();
                    return;
                    
                }

                #region Create Report

                bool isHeader = true;
                string MO2Txt = "";
                if (MO.ToString().Length != 2)
                {
                    MO2Txt = "0" + MO.ToString();
                }
                SubSearch = SelYear + "-" + MO2Txt;

                grade = Grades[selEmp];
                string tempPath = BiometricsData + SelYear + ".dat";
                if (System.IO.File.Exists(tempPath))
                {
                    try
                    {
                        DATLines = File.ReadAllLines(tempPath);
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:02X2206X", "Failed to read files.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();

                    }
                }
                else
                {
                    MessageBox.Show("We couldn't find requested data.\nMake sure you've imported the biometrics logs and selected the correct month, year.\nERROR:02X2213X", "Resources Not Found.", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                List<string> TOSortLines = new List<string>();
                foreach (var item in DATLines)
                {
                    string[] data = item.Split(spearator, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                    if (IDs.Contains(data[0].Trim()) && item.Contains(SubSearch.Trim()))
                    {
                        FillItem = item;
                        TOSortLines.Add(FillItem);
                    }
                }

                TOSortLines.Sort();
                if (TOSortLines.Count == 0)
                {
                    success = false;
                    backgroundWorker3.CancelAsync();
                    MessageBox.Show("Oops! No data to generate a report in selected time period!\nMake sure you've selected the month, year correctly.\nERROR:02X2233X", "Message", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                try
                {
                    wrdApp = new Word.Application();
                    wrdApp.Visible = true;
                }
                catch (Exception)
                {
                    MessageBox.Show("Hmm, We couldn't launch Microsoft® Word to complete the process.\nPlease verify you've installed Microsoft® Word or try opening a new Word document!\nERROR:02X2243X", "Failed to launch Microsoft Word.", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    backgroundWorker2.CancelAsync();
                    success = false;
                    //Close();
                }

                try
                {
                    Object oName = ReportPath;
                    // Open the file to insert data.
                    oDataDoc = wrdApp.Documents.Open(ref oName, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing/*, ref oMissing */);
                }
                catch (Exception)
                {
                    MessageBox.Show("Oops! Unexpected error occured.\nPlease contact the developer with the error code. \nERROR:02X2261X", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.ExitThread();
                    Application.Exit();
                    return;
                }

                if (isHeader)
                {
                    FillData(oDataDoc, true, false, SelMonthText, SelYear, selDept, null);
                    isHeader = false;
                }

                bool FirstDay = true;
                int count = 1;
                string[] Del = { "\t" };
                string[] Del2 = { " " };
                string[] Del3 = { "-" };
                string[] Del4 = { ":" };
                Int32 Wordcount2 = 2;
                string FDate;
                string count2Txt = count.ToString();

                if (FirstDay == true)
                {
                    string[] var_1 = TOSortLines[0].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                    string[] var_2 = var_1[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                    string[] var_3 = var_2[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                    count2Txt = var_3[2];
                    count = Convert.ToInt32(count2Txt);
                    FirstDay = false;
                }
                else
                {
                    count = 1;
                }

                while (count <= 31)
                {
                    if (count.ToString().Length == 1)
                    {
                        count2Txt = "0" + count.ToString();
                    }
                    else
                    {
                        count2Txt = count.ToString();
                    }
                    FDate = SubSearch + "-" + count2Txt; // Format: 2020-06-01
                    try
                    {
                        ThisDay = new DateTime(Convert.ToInt16(SelYear), MO, count);
                    }
                    catch (Exception)
                    {
                        break;
                    }

                    //string _Day = ThisDay.DayOfWeek.ToString();
                    //if (_Day.Substring(0, 3) == "Sat" || _Day.Substring(0, 3) == "Sun")
                    {
                        //SatOrSun = true;
                        //TotalSat_Sun++;
                        //string day = ThisDay.DayOfWeek.ToString();
                        //FillData(oDataDoc, false, true, ThisDay.ToShortDateString() + " (" + day.Substring(0, 3) + ") ", null, null, null);
                    }

                    //FDate = "2020-06-01"
                    DataLines = TOSortLines.FindAll(s => s.Contains(FDate));
                    
                    if (DataLines.Count == 0)
                    {
                        //SatOrSun = false;
                        count++;
                        continue;
                    }
                    else
                    {
                        //TempEmps.Clear();
                        string day = ThisDay.DayOfWeek.ToString();
                        FillData(oDataDoc, false, true, ThisDay.ToShortDateString() + " (" + day.Substring(0, 3) + ") ", null, null, null);
                        foreach (var item in IDs)
                        {
                            TempEmps.Clear();
                            TempEmps = DataLines.FindAll(n => n.Contains(item + "\t"));
                            string empno = item;
                            string selempname = EmpNames[IDs.IndexOf(item)].ToUpper();
                            //while (empno.Length < 10) { empno += " "; }
                            //while (selempname.Length < 25) { selempname += " "; }
                            
                            switch (TempEmps.Count)
                            {
                                case 0:
                                    continue;
                                case 1:
                                    string[] var1 = TempEmps[0].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                    string[] var2 = var1[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                    //Date_1 = var2[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries); //Date_1 > 01 , 06 , 2020
                                    //Time_1 = var2[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries); //Time_1 > 07 , 42 , 10
                                    
                                    FillData(oDataDoc, false, false, empno, selempname, var2[1], "MISSED");
                                    continue;
                                default:
                                    string[] var3 = TempEmps[0].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                    string[] var4 = TempEmps[TempEmps.Count - 1].Split(Del, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                    string[] var5 = var3[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                    string[] var6 = var4[1].Split(Del2, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                    //Date_1 = var5[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                                    //Time_1 = var5[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries);
                                    //Date_2 = var6[0].Split(Del3, 3, StringSplitOptions.RemoveEmptyEntries);
                                    //Time_2 = var6[1].Split(Del4, 3, StringSplitOptions.RemoveEmptyEntries);
                                    
                                    FillData(oDataDoc, false, false, empno, selempname, var5[1], var6[1]);
                                    continue;
                            }
                        }
                        count++;
                    }

                } // END WHILE #31

                //FillData(oDataDoc, false, 
                #endregion
            }
        }

        private void backgroundWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            panel1.Enabled = true;
            toolStripProgressBar1.Visible = false;
            DataLines.Clear();
            TextLines.Clear();
            
            if (fileError | !success)
            {
                toolStripStatusLabel1.Text = "Failed to generate the report!";
            }
            else
            {
                toolStripStatusLabel1.Text = "Report successfully generated!";
            }
           

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (comboBox4.SelectedIndex)
            {
                case 0:
                    selDept = "Administration Division";
                    break;
                case 1:
                    selDept = "Finance Division";
                    break;
                case 2:
                    selDept = "Department of MLS";
                    break;
                case 3:
                    selDept = "Department of Nursing";
                    break;
                case 4:
                    selDept = "Department of Pharmacy";
                    break;
                case 5:
                    selDept = "Library";
                    break;
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }


    }
}