using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using DataTable = System.Data.DataTable;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace OfficeScreenShot
{
    public partial class Form1 : Form
    {
        DataTable dt;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dt = new DataTable();
            dt.Columns.Add("file");
            dt.Columns.Add("name");
            dt.Columns.Add("status");
            dt.Columns.Add("folder");
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog fb = new FolderBrowserDialog();
            if (fb.ShowDialog() == DialogResult.OK)
            {
                txtFolder.Text = fb.SelectedPath;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            dt.Rows.Clear();
            string strFilter = "";
            if (radioPpt.Checked)
                strFilter = ".ppt";
            else if (radioDoc.Checked)
                strFilter = ".doc";
            else if (radioXls.Checked)
                strFilter = ".xls";
            string[] strFiles = Directory.GetFiles(txtFolder.Text, "*" + strFilter, SearchOption.AllDirectories);
            strFiles = strFiles.Where(f => !f.Contains("~$")).ToArray();
            TextWriter writer = new StreamWriter(txtFolder.Text + "\\docer.csv", false, Encoding.UTF8);
            foreach (string f in strFiles)
            {
                writer.WriteLine(Path.GetFileName(f));
                DataRow dr = dt.NewRow();
                dr["file"] = Path.GetFileName(f);
                dr["name"] = Path.GetFileNameWithoutExtension(f);
                dr["folder"] = Path.GetDirectoryName(f);
                dt.Rows.Add(dr);
            }
            writer.Flush();
            writer.Dispose();
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = dt;
            //dataGridView1.Columns[2].Visible = false;

            int pagecount = Convert.ToInt32(txtPage.Text);

            if(radioPpt.Checked)
            {
                InterfaceScreenOriginal screen = new ScreenPowerPoint();
                dt =  screen.ScreenOriginal(dt, pagecount);
            }
            else if (radioDoc.Checked)
            {
                InterfaceScreenOriginal screen = new ScreenOriginWord();
                dt = screen.ScreenOriginal(dt, pagecount);
            }
            else if (radioXls.Checked)
            {
                InterfaceScreenOriginal screen = new ScreenExcel();
                dt = screen.ScreenOriginal(dt, pagecount);
            }

            MessageBox.Show("完成");
        }
    }
}
