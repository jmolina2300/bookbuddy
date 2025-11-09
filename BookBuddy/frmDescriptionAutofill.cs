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
        public bool UseOldMatchingAlgorithm { get; private set; }
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
                dataGridView1.Columns.Add("KEYWORD", "KEYWORD");
                dataGridView1.Columns.Add("OUTPUT1", "OUTPUT1");
                dataGridView1.Rows.Add("","");

                // Loop through Excel rows and populate the DataGridView
                for (int row = 1; row <= range.Rows.Count; row++)
                {
                    string sourceText = null;
                    string replacement = null;
                    try
                    {

                        // Read cell values (Excel uses 1-based index)
                        sourceText = (range.Cells[row, 1] as Range).Value2.ToString();
                        replacement = (range.Cells[row, 2] as Range).Value2.ToString();
                    }
                    catch
                    {
                    }

                    // Only add rows that are not null
                    if (!string.IsNullOrEmpty(sourceText) && !string.IsNullOrEmpty(replacement))
                    {
                        dataGridView1.Rows.Add(sourceText, replacement);
                    }
                    else
                    {
                        dataGridView1.Rows.Add("", "");
                    }
                }



                //string msg = String.Format("This sheet contains {0} rows and {1} columns.", range.Rows.Count, range.Columns.Count);
                //MessageBox.Show(msg,"asd",MessageBoxButtons.OKCancel,MessageBoxIcon.Asterisk);


                int nextColumn = 3;
                int nextDestinationNumber = 2;

                /* Row indices epxlanation:
                 * 
                 * DataGridView.Rows     -->  0-based indexing
                 * DataGridView.Columns  -->  0-based indexing
                 * 
                 * Excel.Range.Rows      -->  1-based indexing
                 * 
                 */

                // If there are more columns, then it is a multi-destination file
                if (range.Columns.Count > 2)
                {
                    // load up the rest of the columns
                    for (int column = nextColumn; column < range.Columns.Count+1; column++)
                    {
                        // Name the column and increment the destination number
                        string columnName = "OUTPUT" + nextDestinationNumber;
                        dataGridView1.Columns.Add(columnName, columnName);
                        nextDestinationNumber += 1;


                        // Then step down the column data and put it into the data grid
                        for (int row = 1; row <= range.Rows.Count; row++)
                        {
                            string cellValue = null;
                            try
                            {
                                /* Read cell values
                                 * 
                                 * Excel uses 1-based indexing, which works in our favor because
                                 * we want the data to start after the 1st row of the DataGridView
                                 * which is used for the name of the excel column destination.
                                 * 
                                 */
                                cellValue = (range.Cells[row, column] as Range).Value2.ToString();

                            }
                            catch
                            {
                            }

                            // Only add row values that are not empty
                            if (!string.IsNullOrEmpty(cellValue))
                            {
                                dataGridView1.Rows[row].Cells[columnName].Value = cellValue;
                                
                            }

                        }
                    }
                }

                // Make the table unsortable because it will mess up the column name thing
                MakeDataGridViewUnsortable();
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


        private void MakeDataGridViewUnsortable()
        {
            // Loop through EVERY column and make it unsortable
            foreach (DataGridViewColumn column in dataGridView1.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }

        private bool ValidateMappingHeader()
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Error: No rows in the table!", "Empty",
                                MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            System.Collections.Hashtable usedColumns = new System.Collections.Hashtable();
            int srcIndex = -1;
            for (int col = 0; col < dataGridView1.Columns.Count; col++)
            {
                string input = "";
                if (dataGridView1.Rows[0].Cells[col].Value != null)
                {
                    input = dataGridView1.Rows[0].Cells[col].Value.ToString().Trim();
                }

                // SOURCE COLUMN (col 0) REQUIRED
                if (col == 0)
                {
                    if (string.IsNullOrEmpty(input))
                    {
                        MessageBox.Show(
                            "Error: First cell (source column) cannot be empty!\n\nEnter the column letter that contains keywords (e.g. A).",
                            "Source Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }

                    // Index of source column gets assigned here
                    srcIndex = ColumnNameToIndex(input);
                    if (srcIndex < 0)
                    {
                        MessageBox.Show(
                            "Error: Source column \"" + input + "\" is invalid!",
                            "Invalid Source", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return false;
                    }
                    continue;
                }


                // Destination columns (col > 0)
                if (string.IsNullOrEmpty(input))
                {
                    // BLANK = SKIP THIS COLUMN  (allowed)
                    continue;
                }


                int destIndex = ColumnNameToIndex(input);
                if (destIndex < 0)
                {
                    MessageBox.Show(
                        "Error: Column " + (col + 1).ToString() +
                        " has invalid name: \"" + input + "\"\n\nUse A, Z, AA, XFD, or leave blank to skip.",
                        "Invalid Column", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }



                // Check for duplicates
                if (usedColumns.ContainsKey(destIndex))
                {
                    MessageBox.Show(
                        "Error: Column \"" + input + "\" is used more than once!\n\nChoose a different destination or leave one blank.",
                        "Duplicate Column", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return false;
                }

                // Check If user is trying to overwrite the source column (allowed)
                if (destIndex == srcIndex)
                {
                    DialogResult dlgResult = MessageBox.Show(
                                    "Warning: Column \"" + input + "\" is used as both the source and destination.\nIf you click OK, the source column will be overwritten.\n\nDo you want to proceed?",
                                    "Source Overwrite", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
                    if (dlgResult != DialogResult.OK)
                    {
                        return false;
                    }
                }

                usedColumns.Add(destIndex, true);
            }

            // Require at least one output column
            if (usedColumns.Count == 0)
            {
                MessageBox.Show(
                    "Warning: No output columns selected!\n\nFill at least one column letter (or cancel).",
                    "No Output", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return false;
            }

            return true;
        }

        public static int ColumnNameToIndex(string name)
        {
            var upperCaseName = name.ToUpper();
            var number = 0;
            var pow = 1;
            if (!System.Text.RegularExpressions.Regex.IsMatch(name, @"^[a-zA-Z]+$"))
            {
                return -1;
            }
            for (var i = upperCaseName.Length - 1; i >= 0; i--)
            {
                number += (upperCaseName[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        }





        /*
         * btnOKMIMO_Click
         * 
         * Do the tasks for MIMO OK button.
         * 
         */
        private void btnOKMIMO_Click(object sender, EventArgs e)
        {
            if (!ValidateMappingHeader())
            {
                return;
            }

            this.DialogResult = DialogResult.OK;
            this.ExtractDataFromGrid();
            this.Close();
        }



        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            // Only care about the first row (header row)
            if (e.RowIndex != 0) return;

            DataGridViewCell cell = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];

            if (cell.Value != null)
            {
                string text = cell.Value.ToString();
                string upper = text.ToUpper();

                // Only update if it actually changed (avoids infinite loop)
                if (text != upper)
                {
                    // Temporarily disable event to avoid event recursion
                    dataGridView1.CellValueChanged -= dataGridView1_CellValueChanged;

                    cell.Value = upper;

                    // Re-enable event
                    dataGridView1.CellValueChanged += dataGridView1_CellValueChanged;
                }
            }

            // Update the color
            ColorCodeHeaderRow();
        }

        private void dataGridView1_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)
        {
            int selectedRow = e.Row.Index;
            if (selectedRow == 0)
            {
                return;
            }
        }

        private void ColorCodeHeaderRow()
        {
            if (dataGridView1.Rows.Count == 0) return;

            DataGridViewRow headerRow = dataGridView1.Rows[0];

            // Reset all cells
            for (int col = 0; col < dataGridView1.Columns.Count; col++)
            {
                DataGridViewCell cell = headerRow.Cells[col];
                cell.Style.BackColor = System.Drawing.Color.White;
                cell.Style.ForeColor = System.Drawing.Color.Black;
                cell.ToolTipText = "";
            }

            for (int col = 0; col < dataGridView1.Columns.Count; col++)
            {
                string input = "";
                if (headerRow.Cells[col].Value != null)
                {
                    input = headerRow.Cells[col].Value.ToString().Trim().ToUpper();
                }

                DataGridViewCell currentCell = headerRow.Cells[col];

                // Source column check
                if (col == 0)
                {
                    if (string.IsNullOrEmpty(input))
                    {
                        currentCell.Style.BackColor = System.Drawing.Color.Salmon;
                        currentCell.ToolTipText = "SOURCE REQUIRED";
                    }
                    else
                    {
                        int idx = ColumnNameToIndex(input);
                        if (idx > 0 && idx <= 16384)
                            currentCell.Style.BackColor = System.Drawing.Color.LightGreen;
                        else
                            currentCell.Style.BackColor = System.Drawing.Color.Salmon;
                    }

                    // Done with source column
                    continue;
                }

                // Check all destination columns
                if (string.IsNullOrEmpty(input))
                {
                    currentCell.Style.BackColor = System.Drawing.Color.White;
                    currentCell.ToolTipText = "Skipped";
                    continue;
                }

                int currentIdx = ColumnNameToIndex(input);

                if (currentIdx <= 0 || currentIdx > 16384)
                {
                    currentCell.Style.BackColor = System.Drawing.Color.Salmon;
                    currentCell.ToolTipText = "Invalid column";
                    continue;
                }

                // CHECK DUPLICATES BY LOOPING THROUGH ALL OTHER CELLS
                bool isDuplicate = false;
                for (int otherCol = 1; otherCol < dataGridView1.Columns.Count; otherCol++)
                {
                    if (otherCol == col)
                    {
                        continue; // skip self
                    }

                    string otherInput = "";
                    if (headerRow.Cells[otherCol].Value != null)
                    {
                        otherInput = headerRow.Cells[otherCol].Value.ToString().Trim().ToUpper();
                    }

                    if (string.IsNullOrEmpty(otherInput))
                    {
                        continue;
                    }

                    int otherIdx = ColumnNameToIndex(otherInput);
                    if (otherIdx == currentIdx)
                    {
                        isDuplicate = true;
                        // Also mark the OTHER cell as duplicate
                        headerRow.Cells[otherCol].Style.BackColor = System.Drawing.Color.Orange;
                        headerRow.Cells[otherCol].ToolTipText = "DUPLICATE";
                    }
                }

                if (isDuplicate)
                {
                    currentCell.Style.BackColor = System.Drawing.Color.Orange;
                    currentCell.ToolTipText = "DUPLICATE";
                }
                else
                {
                    currentCell.Style.BackColor = System.Drawing.Color.LightGreen;
                    currentCell.ToolTipText = "Valid";
                }
            }
        }

        private void frmDescriptionAutofill_Load(object sender, EventArgs e)
        {
            
        }


    }
}
