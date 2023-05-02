using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace AttendanceManager
{
    public partial class EmployeeManagementWindow : Form
    {
        string FilePath = Application.LocalUserAppDataPath + @"\ApplicationData\EmpNames.dat";
        string[] lines = { };
        string selLine;
        public EmployeeManagementWindow()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!System.IO.File.Exists(FilePath))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(FilePath.Replace(@"\EmpNames.dat", ""));
                    using (StreamWriter sw = File.CreateText(FilePath)) { }
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:03X034X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (DirectoryNotFoundException ex)
                {
                    MessageBox.Show("ERROR:03X039X " + ex.Message);
                    Close();
                }
            }
            else
            {
                // Check validity of fields >> Update data
                if (textBox1.Text != "" && textBox2.Text != "" && (string)comboBox1.SelectedItem != null && (string)comboBox2.SelectedItem != null)
                    
                {
                    string no = textBox1.Text;
                    string name = textBox2.Text;
                    string grade = (string)comboBox1.SelectedItem;
                    string dept = (string)comboBox2.SelectedItem;

                    try
                    {
                        if (button1.Text == "ADD")
                        {
                            string[] lines = System.IO.File.ReadAllLines(FilePath);
                            lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                            String[] spearator = { "\t" };
                            List<string> UIds = new List<string>();
                            Int32 Wordcount = 4;
                            foreach (var item in lines)
                            {
                                string[] data = item.Split(spearator, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                                UIds.Add(data[1]);
                            }
                            if (UIds.Contains(no))
                            {
                                MessageBox.Show("Employee already exists!\nYou can update existing data using below section.\nERROR:03X070X", "Duplicate Entry Found", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                return;
                            }
                            using (StreamWriter sw = File.AppendText(FilePath))
                            {
                                sw.WriteLine("\r" + name + "\t" + no + "\t" + grade + "\t" + dept + Environment.NewLine);
                                MessageBox.Show("Record successfully added!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                textBox1.Text = "";
                                textBox2.Text = "";
                                comboBox1.Refresh();
                            } 
                        }
                        else
                        {
                            string tempLine = "";
                            tempLine = name + "\t" + no + "\t" + grade + "\t" + dept;
                            if (tempLine.Trim() == selLine.Trim())
                            {
                                MessageBox.Show("Oops! Looks like there's nothing to update?", "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }
                            else
                            {
                                string[] lines = System.IO.File.ReadAllLines(FilePath);
                                lines = lines.Where(x => !string.IsNullOrEmpty(x)).ToArray();
                                lines[Array.IndexOf(lines, selLine)] = tempLine;
                                
                                using (StreamWriter sw = File.CreateText(FilePath))
                                {
                                    foreach (var item in lines)
                                    {
                                        sw.WriteLine(item);
                                    }
                                }
                                selLine = tempLine;
                                MessageBox.Show("Employee data successfully updated.", "Update Successful", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            
                        }
                    }
                    catch (UnauthorizedAccessException)
                    {
                        MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:03X112X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Close();
                    }
                    catch (DirectoryNotFoundException ex)
                    {
                        MessageBox.Show("ERROR:03X117X " + ex.Message);
                        Close();
                    }
                }
                else
                {
                    MessageBox.Show("Please fill all fields!","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
            }
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            label6.Text = "";
            checkedListBox1.Items.Clear();
            List<string> selEmployees = new List<string>();
            switch (listBox1.SelectedIndex)
            {
                case 0:
                    selEmployees = manageEmployees("Administration Division");
                    break;
                case 1:
                    selEmployees = manageEmployees("Finance Division");
                    break;
                case 2:
                    selEmployees = manageEmployees("Department of MLS");
                    break;
                case 3:
                    selEmployees = manageEmployees("Department of Nursing");
                    break;
                case 4:
                    selEmployees = manageEmployees("Department of Pharmacy");
                    break;
                case 5:
                    selEmployees = manageEmployees("Library");
                    break;
            }

            foreach (var item in selEmployees)
            {
                label6.Text += item + Environment.NewLine;
            }
        }

        private List<string> manageEmployees(string group) //CODE = Admin 01, Fin 02, MLS 03, Nur 04, Phar 05, Lib 06
        {
            List<string> Names = new List<string>();
            String[] spearator = { "\t" };
            Int32 Wordcount = 4;
            bool ErrorStat = false;
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
                            
                            if (data[3] == group)
                            {
                                Names.Add(data[0]);
                                checkedListBox1.Items.Add(data[0]);
                                checkedListBox1.SetItemCheckState(checkedListBox1.Items.IndexOf(data[0]),CheckState.Checked);
                            }
                            else
                            {
                                checkedListBox1.Items.Add(data[0]);
                            }
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Unable to load employee data! Please restart the programme.\nIf this error persists reset employee data through settings section.\nERROR:03X199X", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            ErrorStat = true;
                            break;
                        }

                    }
                    if (ErrorStat == true)
                    {
                        //System.Threading.Thread.CurrentThread.Abort();
                        Close();
                    }

                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:03X214X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
            }
            else
            {
                MessageBox.Show("Oops! Employee name list is empty. You should add users first!\n\nIf you have a backup file, restore it from settings section.\nERROR:03X220X", "No Employees", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            return Names;
        }

        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string FilePath = Application.LocalUserAppDataPath + @"\ApplicationData\EmpNames.dat";
            button1.Text = "Update";
            selLine = "";
            List<string> Names = new List<string>();
            String[] spearator = { "\t" };
            Int32 Wordcount = 4;
            bool ErrorStat = false;
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
                    if (!item.Contains(checkedListBox1.SelectedItem.ToString()))
                    {
                        continue;
                    }
                    try
                    {
                        string[] data = item.Split(spearator, Wordcount, StringSplitOptions.RemoveEmptyEntries);
                        selLine = item;
                        textBox1.Text = data[1];
                        textBox2.Text = data[0];
                        switch (data[3])
                        {
                            case "Administration Division":
                                comboBox2.SelectedItem = comboBox2.Items[0];
                                break;
                            case "Finance Division":
                                comboBox2.SelectedItem = comboBox2.Items[1];
                                break;
                            case "Department of MLS":
                                comboBox2.SelectedItem = comboBox2.Items[2];
                                break;
                            case "Department of Nursing":
                                comboBox2.SelectedItem = comboBox2.Items[3];
                                break;
                            case "Department of Pharmacy":
                                comboBox2.SelectedItem = comboBox2.Items[4];
                                break;
                            case "Library":
                                comboBox2.SelectedItem = comboBox2.Items[5];
                                break;
                        }
                        comboBox1.SelectedItem = comboBox1.Items[comboBox1.Items.IndexOf(data[2])];
                        
                    }
                    catch (Exception)
                    {
                        MessageBox.Show("Unable to load employee data! Please restart the programme.\nIf this error persists reset employee data through settings section.\nERROR:03X284X", "ERROR", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        ErrorStat = true;
                    }
                }
            }
            catch (UnauthorizedAccessException)
            {
                MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:03X291X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
        }

        private void checkedListBox1_ItemCheck(object sender, ItemCheckEventArgs e)
        {
        }

        private void button2_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            button1.Text = "ADD";
        }
    }
}
