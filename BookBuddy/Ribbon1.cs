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
            Worksheet sheet = Globals.ThisAddIn.GetActiveWorkSheet();  // Get the worksheet
            int numChanges = 0;
            int colSrc = ColumnNameToIndex(ed_colBox1.Text);    // Get the source column
            int colDest = ColumnNameToIndex(ed_colBox2.Text);   // Get the destination column
            int numRows = sheet.UsedRange.Rows.Count;

            String sourceText = ed_textBox1.Text;    // Get the source text
            String outputText = ed_textBox2.Text;    // Get the output (desired) text

            // Make sure the rows and columns are OK
            if (!RowsAndColumnsAreOK(ed_colBox1.Text, ed_colBox2.Text, sheet))
            {
                return;
            }       

            // Count the number of cells that will be changed
            String cellText = "";
            for (int i = 1; i <= numRows; i++)
            {
                Excel.Range cellSource = (Excel.Range)sheet.Cells[i, colSrc];
                Excel.Range cellDest = (Excel.Range)sheet.Cells[i, colDest];
                try
                {
                    cellText = cellSource.Value2.ToString();
                    if (cellText.Contains(sourceText))
                    {
                        numChanges += 1;
                    }
                }
                catch (Exception ex) { /* The cellText was null */ }
            }
            if (numChanges == 0) 
            {
                // No strings matched the keyword we were looking for
                MessageBox.Show("No matches found for keyword \"" + sourceText + "\" in column " + ed_colBox1.Text, "Notice");
                return;
            }
            DialogResult d2 = MessageBox.Show(
                numChanges.ToString() + " cells in column " + ed_colBox2.Text + " will be modified.\n\nDo you want to continue?",
                "Confirm Action",
                MessageBoxButtons.YesNo
                );
            if (d2 == DialogResult.No)
            {
                return;
            }
            numChanges = 0;
            // Check the Source column for the string pattern
            for (int i = 1; i <= numRows; i++)
            {
                Excel.Range cellSource = (Excel.Range)sheet.Cells[i, colSrc];
                Excel.Range cellDest = (Excel.Range)sheet.Cells[i, colDest];

                //**
                // BUGBUG: if the row contains ONLY a number, we get an error
                //  We cant cast th cell contents to system string
                //**
                try
                {
                    cellText = cellSource.Value2.ToString();

                    if (cellText.Contains(sourceText))
                    {
                        // If the cell at this row contains the PATTERN.
                        // then set the destination cell contents
                        cellDest.Value2 = outputText;
                        numChanges += 1;
                    }
                }
                catch (Exception ex) { /* The cellText was null */ }
            }
            //MessageBox.Show("Inserted " + numChanges.ToString() + " changes", "Notice");
            //pushChange();
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

        private void group1_DialogLauncherClick(object sender, RibbonControlEventArgs e)
        {
            frmDescriptionAutofill dlg = new frmDescriptionAutofill();

            // Show the form as a dialog (blocking)
            DialogResult mainDialogResult = dlg.ShowDialog();
            if (mainDialogResult != DialogResult.OK)
            {
                return;
            }

            // User clicked OK. Retrieve values from the form...
            string sourceColumnText = dlg.SourceColumn;
            string destinationColumnText = dlg.DestinationColumn;
            Worksheet sheet = Globals.ThisAddIn.GetActiveWorkSheet();  // Get the worksheet

            // Make sure the rows and columns are OK
            if (!RowsAndColumnsAreOK(sourceColumnText, destinationColumnText, sheet))
            {
                return;
            }

            // Everything is fine. Do the text replacement.
            int colSrc = ColumnNameToIndex(sourceColumnText);          // Get the source column
            int colDest = ColumnNameToIndex(destinationColumnText);    // Get the destination column
            int numRows = sheet.UsedRange.Rows.Count;
            int numChanges = 0;

            // Are we doing Multi-in/Single-Out, or Multi-in/Multi-out
            if (dlg.IsMISO) 
            {
                string description = dlg.Description;
                string keywords = dlg.Keywords;
                if (keywords.Length < 1 || description.Length < 1)
                {
                    MessageBox.Show("Please fill out all text fields!", "Warning");
                    return;
                }
                numChanges = descriptionAutofillMISO(colSrc, keywords, colDest, description, sheet);
            }
            else if (dlg.IsMIMO)
            {
                List<string> keywords = dlg.KeywordList;
                List<string> descriptions = dlg.DescriptionList;

                
                if (dlg.UseOldMatchingAlgorithm)
                {
                    numChanges = descriptionAutofillMIMO_old(colSrc, colDest, keywords, descriptions);
                }
                else
                {
                    numChanges = descriptionAutofillMIMO(colSrc, colDest, keywords, descriptions);
                }
            }

            // Tell the user how many cells were modified.
            NotifyChangesToColumn(destinationColumnText, numChanges);
        }

        /* descriptionAutofillMISO
         * 
         * 
         * Fills ONE description for multiple keywords 
         * 
         *   keyword1 ---+-> description1
         *   keyword2 --+
         *   keyword3 -+ 
         */
        private int descriptionAutofillMISO(int colSrc, string sourceKeywords, int colDest, string replacementText, Worksheet sheet)
        {
            // Split the space-separated keywords into an array
            string[] keywords = sourceKeywords.Split(' ');
            int replacements = 0;

            int numRows = sheet.UsedRange.Rows.Count;

            // Loop through each row in the source column
            for (int i = 1; i <= numRows; i++)
            {
                Range sourceCell = (Excel.Range)sheet.Cells[i, colSrc];

                // Skip empty cells or we'll get a null reference 
                if (sourceCell.Value2 == null)
                {
                    continue;
                }
                string cellValue = sourceCell.Value2.ToString() ?? "";

                // Check if any keyword is present in the cell value
                foreach (string keyword in keywords)
                {
                    if (!string.IsNullOrEmpty(keyword) && cellValue.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        // Replace the value in the destination column
                        Range destinationCell = (Range)sheet.Cells[i, colDest];
                        destinationCell.Value2 = replacementText;

                        // Increment the counter for replacements
                        replacements++;
                        break; // No need to check further keywords for this cell
                    }
                }
            }

            // Return the number of replacements made
            return replacements;
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
        private int descriptionAutofillMIMO(int colSrc, int colDest, List<string> sourceTexts, List<string> replacements)
        {


            Excel.Application excelApp = Globals.ThisAddIn.Application;
            Excel.Worksheet activeSheet = (Excel.Worksheet)excelApp.ActiveSheet;

            Excel.Range usedRange = activeSheet.UsedRange;

            // Build a simple list of (source, replacement) pairs
            System.Collections.ArrayList mapList = new System.Collections.ArrayList();

            for (int i = 0; i < sourceTexts.Count; i++)
            {
                string src = (sourceTexts[i] ?? "").Trim();
                if (src.Length == 0) continue;               // skip empty entries

                string repl = (replacements[i] ?? "").Trim();

                // store length + texts so we can sort later
                mapList.Add(new object[] { src.Length, src, repl });
            }


            // Sort by length DESCENDING. Use custom comparer defined below.
            mapList.Sort(new LengthComparer());

 
            // Walk the sheet and apply the first (longest) match
            int numChanges = 0;

            for (int row = 1; row <= usedRange.Rows.Count; row++)
            {
                Excel.Range cellSrc = (Excel.Range)usedRange.Cells[row, colSrc];

                Excel.Range cellDest = (Excel.Range)usedRange.Cells[row, colDest];

                if (cellSrc.Value2 == null) continue;

                string cellValue = cellSrc.Value2.ToString().Trim();

                
                
                for (int m = 0; m < mapList.Count; m++)
                {
                    object[] entry = (object[])mapList[m];
                    string mapSrc = (string)entry[1];

                    // case-insensitive Contains check
                    if (cellValue.IndexOf(mapSrc, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        cellDest.Value2 = (string)entry[2];
                        numChanges++;
                        break;              // stop after the first (longest) match
                    }
                }
            }

            return numChanges;
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

        private void ed_colBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ed_colBox1.Text = ed_colBox1.Text.ToUpper();
        }

        private void ed_colBox2_TextChanged(object sender, RibbonControlEventArgs e)
        {
            ed_colBox2.Text = ed_colBox2.Text.ToUpper();
        }

        
    }
}
