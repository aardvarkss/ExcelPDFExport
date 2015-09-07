using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace WindowsFormsApplication1
{
    public partial class Form1 : Form
    {
        public string saveTo = "";
        public string getFrom = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string path = getFrom;
            notification.Text = "Collating Information";
            button1.Enabled = false;

            Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
            Workbook wb = excel.Workbooks.Open(path);

            // Display the ProgressBar control.
            pBar.Visible = true;
            // Set Minimum to 1 to represent the first file being copied.
            pBar.Minimum = 1;
            // Set Maximum to the total number of files to copy.
            pBar.Maximum = wb.Sheets.Count;
            // Set the initial value of the ProgressBar.
            pBar.Value = 1;
            // Set the Step property to a value of 1 to represent each file being copied.
            pBar.Step = 1;

            //iterate through tabs
            foreach (Worksheet s in wb.Sheets)
            {
                notification.Text = "Printing Invoices";

                // Save into a PDF.
                string filename = saveTo + "\\" + s.Name + ".pdf";
                const int xlQualityStandard = 2;

                s.ExportAsFixedFormat(
                Microsoft.Office.Interop.Excel.XlFixedFormatType.xlTypePDF,
                filename,xlQualityStandard, true, false,
                Type.Missing, Type.Missing, false, Type.Missing);

                pBar.PerformStep();
            }

            //MessageBox.Show(wb.Sheets.Count + " invoices saved.");
            button1.Enabled = true;
            notification.Text = "Process Complete." + wb.Sheets.Count + " invoices saved.";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // Show the dialog and get result.
            DialogResult result = openFileDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox1.Text = openFileDialog1.FileName;
                getFrom = openFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            // Show the dialog and get result.
            DialogResult result = folderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK) // Test result.
            {
                textBox2.Text = folderBrowserDialog1.SelectedPath;
                saveTo = folderBrowserDialog1.SelectedPath;
            }
        }
    }
}
