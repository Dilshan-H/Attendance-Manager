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
    public partial class EditDataForm : Form
    {
        string filePath;
        public EditDataForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
            OpenFileDialog openFileDialog1 = new OpenFileDialog
            {
                Title = "Choose .DAT file",
                FileName = "",
                DefaultExt = "",
                Multiselect = false,
                CheckFileExists = true,
                Filter = "DAT files (*.dat)|*.dat",
            };

            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName != "")
            {
                filePath = openFileDialog1.FileName;
                string data = System.IO.File.ReadAllText(filePath);
                textBox1.Text = data;
                textBox2.Text = filePath;
                button2.Visible = true;
                button3.Visible = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Text = "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult re = new System.Windows.Forms.DialogResult();
            re =  MessageBox.Show("Save changes?", "CAUTION!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (re == System.Windows.Forms.DialogResult.Yes)
            {
                try
                {
                    System.IO.File.WriteAllText(filePath, textBox1.Text);
                }
                catch (Exception ex)
                {
                    MessageBox.Show("\nERROR:ADMINX064X" + ex.Message);
                    Close();
                }
            }
            else
            {

            }
            
        }

        private void EditDataForm_Load(object sender, EventArgs e)
        {
            
        }
    }
}
