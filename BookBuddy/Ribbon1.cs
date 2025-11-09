using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace BookBuddy
{
    public partial class Ribbon1 : OfficeRibbon
    {
        public Ribbon1()
        {
            InitializeComponent();
        }

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }
        private static List<Worksheet> undoList = new List<Worksheet>();
        private const int MAX_UNDO = 5;
        private const string REGEX_ALPHA_NUMERIC = @"[A-Za-z0-9]+";

        private static void pushChange( )
        {
            Worksheet activeSheet = Globals.ThisAddIn.GetActiveWorkSheet();
            if (undoList.Count < MAX_UNDO)
            {
                Worksheet newSheet = ObjectCopier.Clone(activeSheet);  // Copy current sheet
                undoList.Add(newSheet);             // Add it to the undo list
            }
            else
            {
                undoList.RemoveAt(0);               // Remove oldest
                Worksheet newSheet = ObjectCopier.Clone(activeSheet); // Copy current sheet
                undoList.Add(newSheet);             // Add it to the undo list
            }
        }
        private static void popChange( )
        {
            if (undoList.Count < 1)
            {
                return;
            }
            Worksheet lastSheet = undoList.Last();
            undoList.RemoveAt(undoList.Count-1);
            Globals.ThisAddIn.SetActiveWorkSheet(lastSheet);
        }

        public static int ColumnNameToIndex(string name)
        {
            var upperCaseName = name.ToUpper();
            var number = 0;
            var pow = 1;
            if (!Regex.IsMatch(name, @"^[a-zA-Z]+$")) 
            {
                return -1; //Check if input was not a letter
            }
            for (var i = upperCaseName.Length - 1; i >= 0; i--)
            {
                number += (upperCaseName[i] - 'A' + 1) * pow;
                pow *= 26;
            }
            return number;
        }

        private int getNumberOfNonNumericCells(Worksheet sheet, int colIndex)
        {
            int numRows = sheet.UsedRange.Rows.Count;
            int numRowsThatCanChange = 0;
            for (int i = 1; i <= numRows; i++)
            {
                Excel.Range cell = (Excel.Range)sheet.Cells[i, colIndex];
                try
                {
                    String cellText = cell.Value2.ToString();
                    decimal num;
                    if (decimal.TryParse(cellText, out num))
                    {
                        numRowsThatCanChange += 1;
                    }
                }
                catch (Exception ex) { /* The cellText was null */ }
            }

            // Normal case: this is zero
            return numRows - numRowsThatCanChange;
        }

        public void SignFlipColumn(Worksheet sheet, int colIndex)
        {
            String option = cb_pickSign.Text;
            if (!option.Contains('+') && !option.Contains('-'))
            {
                MessageBox.Show("Please use either + or - for the sign.", "Error");
                return;
            }

            /* Mom doesn't like error messages, so leave this commented out.
            int numCellsNonNumeric = getNumberOfNonNumericCells(sheet, colIndex);
            if (numCellsNonNumeric != 0)
            {
                MessageBox.Show("Warning: This column contains " + numCellsNonNumeric + " non-numeric cells. These won't be modified.", "Warning");
            }
            */
            int numChanges = 0;
            int numRows = sheet.UsedRange.Rows.Count;
            for (int i = 1; i <= numRows; i++)
            {
                Excel.Range cell = (Excel.Range)sheet.Cells[i, colIndex];
                try
                {
                    String cellText = cell.Value2.ToString();
                    decimal num;
                    if (decimal.TryParse(cellText, out num))
                    {
                        if (option.Contains('+'))
                        {
                            cell.Value2 = Math.Abs(num);    // Make positive
                            numChanges += 1;
                        }
                        else if (option.Contains('-'))
                        {
                            cell.Value2 = Math.Abs(num) * -1;  // make negative
                            numChanges += 1;
                        }
                        cell.NumberFormat = "0.00"; // Always two decimal places
                    }
                }
                catch (Exception ex) { /* The cellText was null */ }
            }

            MessageBox.Show(
                numChanges.ToString() + " cells in column " + ed_signFlip_colBox.Text.ToUpper() + " were modified.",
                "Notice"
                );
        }
        private void btn_go_Click(object sender, RibbonControlEventArgs e)
        {
            frmDescriptionAutofill dlg = new frmDescriptionAutofill();

            // Show the form as a dialog (blocking)
            DialogResult mainDialogResult = dlg.ShowDialog();
            if (mainDialogResult != DialogResult.OK)
            {
                return;
            }

            int numChanges = descriptionAutofillMIMO_MultiColumn(dlg.dataGridView1);

            // Tell the user how many cells were modified.
            MessageBox.Show(numChanges + " cells modified.", "Notice");
        }

        private void btn_go_signFlip_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet sheet = Globals.ThisAddIn.GetActiveWorkSheet();
            int column = ColumnNameToIndex(ed_signFlip_colBox.Text);
            if (column < 0)
            {
                MessageBox.Show("Invalid column \"" + ed_signFlip_colBox.Text+"\"!", "Error");
                return;
            }

            SignFlipColumn(sheet, column);
        }

        private void btn_undo_Click(object sender, RibbonControlEventArgs e)
        {
            popChange();
            Debug.WriteLine("Undo pressed");

        }

        private void btn_go_cellCleanup_Click(object sender, RibbonControlEventArgs e)
        {
            Worksheet sheet = Globals.ThisAddIn.GetActiveWorkSheet();     // Get the worksheet
            int column = ColumnNameToIndex(ed_cellCleanup_column.Text);    // Get the source column
            if (column < 0)
            {
                MessageBox.Show("Invalid column \"" + ed_cellCleanup_column.Text + "\"!", "Error");
                return;
            }
            removeCharacters(sheet, column);
        }

        private void NotifyChangesToColumn(string column, int numChanges)
        {
            MessageBox.Show(
                numChanges.ToString() + " cells in column " + column + " were modified.",
                "Notice"
                );
        }

        /* getRegexMatchSingle()
         * 
         * Returns a concatenated string of all Regex matches
         * 
         */
        private string getRegexMatchSingle(string input, string pattern, RegexOptions options)
        {
            string matches = "";
            foreach (Match m in Regex.Matches(input, pattern, options))
            {
                matches += m.Value.ToString();
            }
            return matches;
        }

        private void removeCharacters(Worksheet sheet, int column)
        {      
            int numChanges = 0;

            int numRows = sheet.UsedRange.Rows.Count;

            string exclusiveRegex = getExclusiveRegex(ed_cellCleanup_characters.Text);

            for (int row = 1; row <= numRows; row++)
            {
                Excel.Range cell = (Excel.Range)sheet.Cells[row, column];
                try
                {
                    string cellText = cell.Value2.ToString();
                    string newCellText = getRegexMatchSingle(cellText, exclusiveRegex, RegexOptions.Multiline);
                    cell.NumberFormat = "@";
                    cell.Value2 = newCellText;
                    numChanges++;
                }
                catch (Exception ex) { /* The cellText was null */ }
            }
            NotifyChangesToColumn(ed_cellCleanup_column.Text, numChanges);
        }

        /* getExclusiveRegex()
         * 
         * Returns a regex that excludes the given string
         * 
         */
        private string getExclusiveRegex(string userInput)
        {
            return @"[^" + userInput + "]+";
        }



        private void makeCellsAlphaNumeric()
        {
            Worksheet sheet = Globals.ThisAddIn.GetActiveWorkSheet();     // Get the worksheet
            int column = ColumnNameToIndex(ed_cellCleanup_column.Text);    // Get the source column
            if (column < 0)
            {
                MessageBox.Show("Invalid column \"" + ed_cellCleanup_column.Text + "\"!", "Error");
                return;
            }
            int numChanges = 0;
            int numRows = sheet.UsedRange.Rows.Count;
            for (int row = 1; row <= numRows; row++)
            {
                Excel.Range cell = (Excel.Range)sheet.Cells[row, column];
                try
                {
                    String cellText = cell.Value2.ToString();
                    String alphaNumeric = getRegexMatchSingle(cellText, REGEX_ALPHA_NUMERIC, RegexOptions.Multiline);
                    cell.Value2 = alphaNumeric;
                    numChanges++;
                    
                }
                catch (Exception ex) { /* The cellText was null */ }
            }
            NotifyChangesToColumn(ed_cellCleanup_column.Text, numChanges);
        }

        private bool RowsAndColumnsAreOK(string sourceColumnText, string destinationColumnText, Worksheet sheet)
        {
            int colSrc = ColumnNameToIndex(sourceColumnText);          // Get the source column
            int colDest = ColumnNameToIndex(destinationColumnText);    // Get the destination column
            int numRows = sheet.UsedRange.Rows.Count;

            if (numRows < 1)
            {
                return false;  // Return if there are no rows being used
            }
            if (colSrc < 0)
            {
                MessageBox.Show("Invalid source column \"" + sourceColumnText + "\"!", "Warning");
                return false;
            }
            if (colDest < 0)
            {
                MessageBox.Show("Invalid destination column \"" + destinationColumnText + "\"!", "Warning");
                return false;
            }

            if (colDest == colSrc)
            {
                // Check if source column and dest column are the same 
                DialogResult d1 = MessageBox.Show(
                    "Your source column will be overwritten!\n\nDo you want to continue?",
                    "Confirm Action",
                    MessageBoxButtons.YesNo
                );
                if (d1 == DialogResult.No)
                {
                    return false;
                }
            }

            return true;
        }



        /* descriptionAutofillMIMO
         * 
         * 
         * Fills Multiple descriptions for multiple keywords 
         * 
         *   keyword1 -> description1
         *   keyword2 -> description2
         *   keyword3 -> description3
         * 
         */
        private int descriptionAutofillMIMO_MultiColumn(DataGridView dgvMapping)
        {
            if (dgvMapping == null || dgvMapping.Rows.Count < 2)
                throw new ArgumentException("DataGridView must have at least 2 rows.");

            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet =(Excel.Worksheet)excelApp.ActiveSheet;
            Excel.Range usedRange = activeSheet.UsedRange;


            // 1. Read column mappings from FIRST ROW of DataGridView
            int srcColExcel = 0;
            var destColExcelList = new System.Collections.ArrayList(); // int list

            for (int c = 0; c < dgvMapping.Columns.Count; c++)
            {
                string header = (dgvMapping.Rows[0].Cells[c].Value ?? "").ToString().Trim().ToUpper();

                if (string.IsNullOrEmpty(header)) continue;

                int excelCol = ColumnNameToIndex(header); // Column index

                if (c == 0)
                {
                    srcColExcel = excelCol;
                }
                else
                {
                    destColExcelList.Add(excelCol);
                }
            }

            if (srcColExcel == 0)
                throw new Exception("First column in row 1 must contain source column letter (e.g. A).");

            if (destColExcelList.Count == 0)
                throw new Exception("At least one destination column letter required in row 1.");

            
            // 2. Build mapping list: source -> array of replacements
            var mappings = new System.Collections.ArrayList();  // { length, source, repl1, repl2, ... }

            for (int r = 1; r < dgvMapping.Rows.Count; r++)     // start at row 1 (0-based => second row)
            {
                DataGridViewRow dgvRow = dgvMapping.Rows[r];
                if (dgvRow.IsNewRow) continue;

                string sourceText = (dgvRow.Cells[0].Value ?? "").ToString().Trim();
                if (string.IsNullOrEmpty(sourceText)) continue;

                var rowData = new System.Collections.ArrayList();
                rowData.Add(sourceText.Length);     // 0: length for sorting
                rowData.Add(sourceText);            // 1: source

                for (int c = 1; c < dgvRow.Cells.Count; c++)
                {
                    string val = (dgvRow.Cells[c].Value ?? "").ToString().Trim();
                    rowData.Add(val);
                }

                mappings.Add(rowData);
            }

            // Sort: longest source first
            mappings.Sort(new LengthComparerMulti());

            

            // 3. Apply to Excel sheet
            int numChanges = 0;
            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                Excel.Range srcCell = (Excel.Range)usedRange.Cells[row, srcColExcel];

                if (srcCell.Value2 == null)
                {
                    continue;
                }

                string cellValue = srcCell.Value2.ToString().Trim();
                bool matched = false;
                foreach (System.Collections.ArrayList map in mappings)
                {
                    string mapSource = (string)map[1];

                    if (cellValue.IndexOf(mapSource, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Write to ALL destination columns
                        for (int i = 0; i < destColExcelList.Count; i++)
                        {
                            int destCol = (int)destColExcelList[i];
                            string replacement = (string)map[2 + i]; // 2 = source, 3+ = outputs

                            Excel.Range destCell = (Excel.Range)usedRange.Cells[row, destCol];

                            destCell.Value2 = replacement;
                        }

                        numChanges++;
                        matched = true;
                        break; // longest match wins
                    }
                }
            }

            return numChanges;
        }

        
        private int ColumnLetterToNumber(string columnLetter)
        {
            columnLetter = columnLetter.ToUpper();
            int result = 0;
            for (int i = 0; i < columnLetter.Length; i++)
            {
                result *= 26;
                result += (columnLetter[i] - 'A' + 1);
            }
            return result;
        }


        private int descriptionAutofillMIMO_old(int colSrc, int colDest, List<string> sourceTexts, List<string> replacements)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = Globals.ThisAddIn.Application;
            Worksheet activeSheet = (Worksheet)excelApp.ActiveSheet;

            Range usedRange = activeSheet.UsedRange;

            // Loop through each row in the used range
            int numChanges = 0;
            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                Range cellA = (Range)usedRange.Cells[row, colSrc];
                Range cellB = (Range)usedRange.Cells[row, colDest];

                if (cellA.Value2 != null)
                {
                    string cellValue = cellA.Value2.ToString();

                    // Search through the SourceTexts for a match
                    for (int i = 0; i < sourceTexts.Count; i++)
                    {
                        if (cellValue.Contains(sourceTexts[i]))
                        {
                            cellB.Value2 = replacements[i]; // Place the replacement in Column B
                            numChanges++;
                            break;                          // Stop searching after the first match
                        }
                    }
                }
            }
            return numChanges;
        }




        /* class: LengthComparerMulti
         * 
         * Simple comparer that sorts by the first element (the length) descending
         * 
         */
        private class LengthComparerMulti : System.Collections.IComparer
        {
            public int Compare(object x, object y)
            {
                System.Collections.ArrayList a = (System.Collections.ArrayList)x;
                System.Collections.ArrayList b = (System.Collections.ArrayList)y;

                int lenA = (int)a[0];
                int lenB = (int)b[0];

                return lenB.CompareTo(lenA); // longer first
            }
        }


        /* class: LengthComparer
         * 
         * Simple comparer that sorts by the first element (the length) descending
         * 
         */
        private class LengthComparer : System.Collections.IComparer
        {
            public int Compare(object x, object y)
            {
                object[] a = (object[])x;
                object[] b = (object[])y;

                int lenA = (int)a[0];
                int lenB = (int)b[0];

                // longer first
                return lenB.CompareTo(lenA);
            }
        }


    }
}
