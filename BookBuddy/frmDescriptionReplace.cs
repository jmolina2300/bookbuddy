using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

// InterOP stuff for loading EXCEL files (xls)
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace BookBuddy
{
    public partial class frmDescriptionReplace : Form
    {
        public string SourceColumn { get; private set; }
        public string SourceContent { get; private set; }
        public string DestinationColumn { get; private set; }
        public string DestinationContent { get; private set; }
        public bool isMISO { get; private set; }
        public bool isMIMO { get; private set; }

        public frmDescriptionReplace()
        {
            InitializeComponent();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            // Capture values before closing
            SourceColumn = txtSourceColumn.Text;
            SourceContent = txtSourceContent.Text;
            DestinationColumn = txtDestinationColumn.Text;
            DestinationContent = txtDestinationContent.Text;
            
            // Close the form with OK result
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnImportCSV_Click(object sender, EventArgs e)
        {
            // Create a new OpenFileDialog instance
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                //Filter = "CSV files (*.csv)|*.csv",
                Title = "Select a CSV File"
            };

            // Show the dialog and check if the user clicked OK
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Get the file path from the dialog
                string filePath = openFileDialog.FileName;

                // Load the CSV data into the DataGridView
                LoadCSVIntoDataGridView(filePath);
            }
        }

        private void LoadCSVIntoDataGridView(string filePath)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            Workbook workbook = null;
            Worksheet worksheet = null;
            Range range = null;

            try
            {
                // Open the Excel file
                workbook = excelApp.Workbooks.Open(filePath, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                   Type.Missing, Type.Missing, Type.Missing);
                worksheet = (Worksheet)workbook.Sheets[1]; // Read the first sheet
                range = worksheet.UsedRange;

                // Clear any existing rows and columns
                dataGridView1.Rows.Clear();
                dataGridView1.Columns.Clear();

                // Add columns to the DataGridView
                dataGridView1.Columns.Add("SourceText", "Source Text");
                dataGridView1.Columns.Add("Replacement", "Replacement");

                // Loop through Excel rows and populate the DataGridView
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    // Read cell values (Excel uses 1-based index)
                    string sourceText = (range.Cells[row, 1] as Range).Value2.ToString();
                    string replacement = (range.Cells[row, 2] as Range).Value2.ToString();

                    // Only add rows that are not null
                    if (!string.IsNullOrEmpty(sourceText) && !string.IsNullOrEmpty(replacement))
                    {
                        dataGridView1.Rows.Add(sourceText, replacement);
                    }
                }

                MessageBox.Show("Excel data loaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading Excel data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up
                workbook.Close(false, Type.Missing, Type.Missing);
                Marshal.ReleaseComObject(worksheet);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(excelApp);
            }
        }


        private void LoadCSVIntoDataGridViewOLD(string filePath)
        {
            // Clear existing rows and columns
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            // Add columns
            dataGridView1.Columns.Add("SourceText", "Source Text");
            dataGridView1.Columns.Add("Replacement", "Replacement");

            // Read the CSV file line by line
            using (StreamReader reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    string[] values = line.Split(',');

                    // Only add rows that have exactly two columns
                    if (values.Length == 2)
                    {
                        dataGridView1.Rows.Add(values[0].Trim(), values[1].Trim());
                    }
                }
            }

            MessageBox.Show("CSV Loaded Successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

    }
}
