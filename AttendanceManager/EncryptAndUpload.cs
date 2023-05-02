using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Security.Cryptography;

namespace AttendanceManager
{
    public partial class EncryptAndUpload : Form
    {
        public EncryptAndUpload()
        {
            InitializeComponent();
        }

        private void EncryptAndUpload_Load(object sender, EventArgs e)
        {
            int a = 2015;
            while (a <= 2040)
            {
                comboBox1.Items.Add(a);
                a++;
            }
            DateTime now = DateTime.Now;
            int ThisYear = now.Year;

            if (comboBox1.Items.Contains(ThisYear))
            {
                int index = comboBox1.Items.IndexOf(ThisYear);
                comboBox1.SelectedItem = comboBox1.Items[index];
            }
            else
            {
                MessageBox.Show("Oops! Outdated version -- Check system date or contact developer.\n\nPS: If you seeing this message after 2040...Yeah, this may seems like a cr**! Who knows what may happen in the future???", "Error occured", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Close();
            }
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem == null)
            {
                MessageBox.Show("Oops! Invalid Input detected!", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            else
            {
                string selYear = comboBox1.Items[comboBox1.SelectedIndex].ToString();
                string path = Application.LocalUserAppDataPath + @"\ApplicationData\" + selYear + ".dat";

                if (!System.IO.File.Exists(path))
                {
                    MessageBox.Show("Oops! No resources found related to the year \"" + selYear + "\" \nPlease check the year again.\nERROR:EncUplX043X", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    saveFileDialog1.Title = "Choose a location to save encrypted files.";
                    saveFileDialog1.Filter = "TEMP File|*.tmp";
                    saveFileDialog1.DefaultExt = ".tmp";
                    saveFileDialog1.FileName = selYear;

                    if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                    {
                        if (!saveFileDialog1.FileName.ToLower().Contains(selYear + ".tmp"))
                        {
                            MessageBox.Show("You can't change the default file name.");
                            saveFileDialog1.FileName = selYear;
                            saveFileDialog1.ShowDialog();
                            return;
                        }
                        try
                        {
                            System.IO.File.Copy(path, saveFileDialog1.FileName, true);
                            bool stat = EncryptFile(saveFileDialog1.FileName);
                            if (!stat)
                            {
                                MessageBox.Show("Encryption unsuccessful! - Please try again.\n\nERROR:EncUplX082X", "Unexpected Error Occured", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                            }
                            else
                            {
                                MessageBox.Show("Encryption successful!\nPlease use the admin login portal to upload the file.\n(After clicking 'OK' the login page will open in your browser)", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                //Open 'ADMIN' login page
                            }
                        }
                        catch (UnauthorizedAccessException)
                        {
                            MessageBox.Show("Oops! We don't have permission to access into the file system.\nPlease run the programme as administrator.\nERROR:EncUplX092X", "Backup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Close();
                        }
                        catch (FileNotFoundException)
                        {
                            MessageBox.Show("Oops! We can't find resource files to backup. Check if you have added employees to the system.\nERROR:EncUplX097X", "Backup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Close();
                        }
                        catch (DirectoryNotFoundException)
                        {
                            MessageBox.Show("Oops! We can't find resource files to backup. Check if you have added employees to the system.\nERROR:EncUplX102X", "Backup Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            Close();
                        }
                    }

                }
                 
            }
        }

        private bool EncryptFile(string inputFile)
        {

            try
            {
                //EXPERIMENTAL FEATURE == NOT IMPLEMENTED YET
                string pass = @"ylwRe1Ccx4ZTp9NK";// z7QR90Co4wHZZOO - ";//v3anLbcpqFE=";
                //UnicodeEncoding UE = new UnicodeEncoding();
                byte[] key = Encoding.ASCII.GetBytes(pass);//UE.GetBytes(pass);

                string cryptFile = inputFile.Replace(".tmp",".hdenc");
                FileStream fsCrypt = new FileStream(cryptFile, FileMode.Create);

                RijndaelManaged RMCrypto = new RijndaelManaged();

                CryptoStream cs = new CryptoStream(fsCrypt,
                    RMCrypto.CreateEncryptor(key,key),
                    CryptoStreamMode.Write);

                FileStream fsIn = new FileStream(inputFile, FileMode.Open);

                int data;
                while ((data = fsIn.ReadByte()) != -1)
                    cs.WriteByte((byte)data);


                fsIn.Close();
                cs.Close();
                fsCrypt.Close();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
