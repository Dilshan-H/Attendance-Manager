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
    public partial class SettingsForm : Form
    {
        string FilePath = Application.LocalUserAppDataPath + @"\ApplicationData\EmpNames.dat";
        string ReportsPath = Application.LocalUserAppDataPath + @"\ApplicationData\Reports";
        bool proceed = false;
        public SettingsForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            DialogResult rep = MessageBox.Show("Employee data (names and numbers) will be deleted!\nIf you have a backup file restore it instead of clearing data.\n\nProceed?", "Reset Data", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (rep == DialogResult.Yes)
            {
                try
                {
                    if (System.IO.File.Exists(FilePath))
                    {
                        System.IO.File.WriteAllText(FilePath, string.Empty);
                        MessageBox.Show("Employee data sucessfully deleted!", "Message", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:SettingsX039X", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (DirectoryNotFoundException ex)
                {
                    MessageBox.Show("ERROR:SettingsX039X " + ex.Message);
                    Close();
                }
            }
            else
            {
                //do_nothing
            }
            
            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            AboutWindow ShowAbout = new AboutWindow();
            ShowAbout.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            AboutWindow ShowAbout = new AboutWindow();
            ShowAbout.Show();
        }

        private void button2_DragDrop(object sender, DragEventArgs e)
        {
            if (proceed == true)
            {
                EditDataForm ShowEditor = new EditDataForm();
                ShowEditor.Show();
            }
            else
            {
                Close();
            }
        }

        private void button2_DragOver(object sender, DragEventArgs e)
        {
            DialogResult re = new System.Windows.Forms.DialogResult();
            re =  MessageBox.Show("You're entering to the developer mode!\nThis will allow you to alter/format DAT files.\nWarning! ONLY for research/educational purposes.\nYou would not hold HD Software responsible for any and all those actions/file alterations.\nClicking on 'Yes' means that you have accepted all the terms & conditions!\nProceed?", "CAUTION!", MessageBoxButtons.YesNo, MessageBoxIcon.Asterisk);
            if (re == System.Windows.Forms.DialogResult.Yes)
            {
                proceed = true;
            }
            else
            {
                proceed = false;
            }
            if (proceed == true)
            {
                EditDataForm ShowEditor = new EditDataForm();
                ShowEditor.Show();
                Close();
            }
            else
            {
                Close();
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            //Delete old reports @localappdata >> 'Reports' (Dir)
            DialogResult rep = MessageBox.Show("This action will delete all the previous reports you've generated!\nAre you sure?", "Confirmation", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            string thisFile = "";
            if (rep == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    if (System.IO.Directory.Exists(ReportsPath))
                    {
                        System.IO.DirectoryInfo Reports = new DirectoryInfo(ReportsPath);
                        foreach (FileInfo file in Reports.GetFiles())
                        {
                            thisFile = file.Name;
                            file.Delete();
                        }
                        foreach (DirectoryInfo dir in Reports.GetDirectories())
                        {
                            dir.Delete(true);
                        }
                        MessageBox.Show("Reports successfully deleted.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    }
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:SettingsX138X", "Elevated Privileges Required", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (IOException)
                {
                    MessageBox.Show("Oops! The process cannot access the file: [ " + thisFile + " ] because it is being used by another process.\n\nPlease close any opened reports and try again.\nERROR:SettingsX143X", "Failed to delete files", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
            }
            
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleLeft;
            label2.Text = "*** Warning! ***\n\nThis action will delete Employee Data list (including names and machine numbers) stored in the system.\nYou won't be able to generate reports until adding employees to the system again.\n\nPlease use 'Backup' or 'Restore' feature before proceed.";
        }

        private void button5_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleLeft;
            label2.Text = "This action will remove all the previous reports you've generated.";
        }

        private void button6_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleLeft;
            label2.Text = "*** Warning! - Only use this feature, if generated reports have faults.\n\nThis action will delete 'Biometrics Log Data' stored in the system.\nYou won't be able to generate reports until importing new Biometrics Log Data to the system again.";
        }

        private void button2_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Text = "Read 'Terms and Conditions' of this software.";
        }

        private void button3_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Text = "Learn more about the software - About 'Attendance Manager'";
        }

        private void button4_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Text = "Close this window.";
        }

        private void button7_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Text = "Backup Employee Data list (including names and machine numbers).";
        }

        private void button8_MouseHover(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleLeft;
            label2.Text = "Restore Employee Data list (including names and machine numbers) using previously saved backup file.";
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Text = "Attendance Manager - (BETA) \nCopyright © 2021 - HD Software \nAll rights reserved.";
        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            label2.TextAlign = ContentAlignment.MiddleCenter;
            label2.Text = "Attendance Manager - (BETA) \nCopyright © 2021 - HD Software \nAll rights reserved.\n\nSimply The Best!";
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string FilePath = Application.LocalUserAppDataPath + @"\ApplicationData\EmpNames.dat";
            saveFileDialog1.Filter = "DAT File|*.dat";
            saveFileDialog1.DefaultExt = "dat";
            saveFileDialog1.FileName = "Backup_AttendanceManager_#H24D";
            

            if (saveFileDialog1.ShowDialog() == DialogResult.OK) 
            {
                if (!saveFileDialog1.FileName.ToLower().Contains("backup_attendancemanager_#h24d.dat"))
                {
                    MessageBox.Show("You can't change the default file name.");
                    saveFileDialog1.FileName = "Backup_AttendanceManager_H24D";
                    saveFileDialog1.ShowDialog();
                    return;
                }
                try
                {
                    System.IO.File.Copy(FilePath, saveFileDialog1.FileName, true);
                    MessageBox.Show("Backup Successful!\nUse this file on restoration process whenever you needed.\n\n[ Backup Location ] " + saveFileDialog1.FileName, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:SettingsX236X", "Backup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (FileNotFoundException)
                {
                    MessageBox.Show("Oops! We can't find resource files to backup. Check if you have added employees to the system.\nERROR:SettingsX241X", "Backup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
                catch (DirectoryNotFoundException)
                {
                    MessageBox.Show("Oops! We can't find resource files to backup. Check if you have added employees to the system.\nERROR:SettingsX241X", "Backup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {

                Title = "Choose Backup File to restore",
                FileName = "",
                Multiselect = false,
                CheckFileExists = true,
                Filter = "DAT files (*.dat)|*.dat",
            };

            openFileDialog1.ShowDialog();
            if (!openFileDialog1.FileName.ToLower().Contains("backup_attendancemanager_#h24d.dat"))
            {
                MessageBox.Show("Oops! Invalid backup file selected!\n\nPlease check the file again.\nERROR:SettingsX257X", "Failed to restore data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (openFileDialog1.FileName != "") 
            {
                string FilePath = Application.LocalUserAppDataPath + @"\ApplicationData\EmpNames.dat";
                try
                {
                    System.IO.Directory.CreateDirectory(Application.LocalUserAppDataPath + @"\ApplicationData");
                    System.IO.File.Copy(openFileDialog1.FileName, FilePath, true);
                    MessageBox.Show("Data successfully restored!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (UnauthorizedAccessException)
                {
                    MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:SettingsX271X", "Failed to restore data", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Close();
                }
            }
        }
    }
}
