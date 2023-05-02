using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace AttendanceManager
{
    public partial class MainWindow : Form
    {
        public MainWindow()
        {
            
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FindDataWindow ShowF2 = new FindDataWindow();
            ShowF2.Show();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            EmployeeManagementWindow ShowF3 = new EmployeeManagementWindow();
            ShowF3.Show();
        }

        private void MainWindow_Load(object sender, EventArgs e)
        {
        }

        private void label1_MouseMove(object sender, MouseEventArgs e)
        {
            
        }

        private void label1_Click(object sender, EventArgs e)
        {
            AboutWindow ShowAbout = new AboutWindow();
            ShowAbout.Show();
        }

        ToolTip AboutData = new ToolTip { ToolTipTitle = "Read More.", IsBalloon = false, ToolTipIcon = ToolTipIcon.Info, UseFading = true, UseAnimation = true, ReshowDelay = 5000 };
        ToolTip Btn1 = new ToolTip { ToolTipTitle = "Generate reports", IsBalloon = false, ToolTipIcon = ToolTipIcon.Info, UseFading = true, UseAnimation = true };
        ToolTip Btn2 = new ToolTip { ToolTipTitle = "Add employee data", IsBalloon = false, ToolTipIcon = ToolTipIcon.Info, UseFading = true, UseAnimation = true };
        ToolTip Btn3 = new ToolTip { IsBalloon = false, UseFading = true, UseAnimation = true };



        private void label1_MouseHover(object sender, EventArgs e)
        {
            AboutData.Show("Click here to show the about section & the disclaimer notice.", label1, 5000);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            SettingsForm ShowSettings = new SettingsForm();
            ShowSettings.Show();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://forms.office.com/Pages/ResponsePage.aspx?id=DQSIkWdsW0yxEjajBLZtrQAAAAAAAAAAAAa__fBnlPxURjNHTkVTU1JSNUJBUFY3VEJNSEE1MkdDSi4u");
        }

        private void button1_MouseHover(object sender, EventArgs e)
        {
            Btn1.Show("Click here to view fingerprint details and generate monthly reports.", button1, 2000);
        }

        private void button2_MouseHover(object sender, EventArgs e)
        {
            Btn2.Show("Click here to manage employees' data.", button2, 2000);
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
           
        }

        private void button2_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void button2_MouseLeave(object sender, EventArgs e)
        {

        }

        private void button3_MouseHover(object sender, EventArgs e)
        {
            Btn3.Show("Settings", button3, 2000);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            EncryptAndUpload showForm = new EncryptAndUpload();
            showForm.Show();
        }

        private void button5_MouseHover(object sender, EventArgs e)
        {
            Btn3.Show("Upload data to AMS - NOT IMPLEMENTED", button5, 3000);
        }
    }
}
