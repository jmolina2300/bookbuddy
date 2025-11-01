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
    public partial class frmDescriptionAutofill : Form
    {
        public string SourceColumn { get; private set; }
        public string Keywords { get; private set; }
        public string DestinationColumn { get; private set; }
        public string Description { get; private set; }
        public bool IsMISO { get; private set; }
        public bool IsMIMO { get; private set; }
        public List<string> KeywordList { get; private set; }
        public List<string> DescriptionList { get; private set; }

        private Color ColorERR = Color.FromArgb(255, 255, 120, 15);
        private Color ColorOK = Color.FromArgb(255, 255, 255, 255);

        private void ExtractDataFromGrid()
        {
            KeywordList = new List<string>();
            DescriptionList = new List<string>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Cells[0].Value != null && row.Cells[1].Value != null)
                {
                    KeywordList.Add(row.Cells[0].Value.ToString());
                    DescriptionList.Add(row.Cells[1].Value.ToString());
                }
            }
        }

        public frmDescriptionAutofill()
        {
            InitializeComponent();
        }


        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnCancel2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void btnImportCSV_Click(object sender, EventArgs e)
        {
            // Create a new OpenFileDialog instance
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx|CSV Files (*.csv)|*.csv",
                Title = "Select a CSV or XLS File"
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

        /*
         * LoadCSVIntoDataGridView
         * 
         * Load CSV or XLS data into the DataGridView.
         * 
         * Input:  filePath (string)
         * 
         * Output: None
         * 
         */
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
                dataGridView1.Columns.Add("Keyword", "Keyword");
                dataGridView1.Columns.Add("AutofillDescription", "Autofill Description");

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

                MessageBox.Show("Data loaded successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error loading data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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


        /*
         * BothColumnsFilled
         * 
         * Check if both column fields have been filled out.
         * 
         */
        private bool BothColumnsFilled()
        {
            bool missingColumnA = string.IsNullOrEmpty(txtSourceColumn.Text);
            bool missingColumnB = string.IsNullOrEmpty(txtDestinationColumn.Text);
            if (missingColumnA || missingColumnB)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void colorColumnBoxesBasedOnContent()
        {
            if (string.IsNullOrEmpty(txtSourceColumn.Text))
            {
                txtSourceColumn.BackColor = ColorERR;
            }
            else
            {
                txtSourceColumn.BackColor = ColorOK;
            }

            if (string.IsNullOrEmpty(txtDestinationColumn.Text))
            {
                txtDestinationColumn.BackColor = ColorERR;
            }
            else
            {
                txtDestinationColumn.BackColor = ColorOK;
            }
        }

        /*
         * btnOKMIMO_Click
         * 
         * Do the tasks for MIMO OK button.
         * 
         */
        private void btnOKMIMO_Click(object sender, EventArgs e)
        {
            // Color the text boxes based on whether they have text in them
            this.colorColumnBoxesBasedOnContent();

            // Make sure both column fields are filled out
            if (!BothColumnsFilled())
            {
                MessageBox.Show("Please choose both a source and destination column", "Warning");
                return;
            }

            // Capture values before closing
            SourceColumn = txtSourceColumn.Text;
            DestinationColumn = txtDestinationColumn.Text;

            // Close the form with OK result
            this.DialogResult = DialogResult.OK;
            this.ExtractDataFromGrid();
            this.IsMIMO = true;
            this.IsMISO = !IsMIMO;
            this.Close();
        }

        private void txtSourceColumn_TextChanged(object sender, EventArgs e)
        {
            txtSourceColumn.BackColor = ColorOK;
        }

        private void txtDestinationColumn_TextChanged(object sender, EventArgs e)
        {
            txtDestinationColumn.BackColor = ColorOK;
        }



    }
}
