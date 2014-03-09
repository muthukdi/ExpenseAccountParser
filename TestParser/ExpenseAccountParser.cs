using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;
using System.IO;

namespace Dilip
{
    class ExpenseAccountParser
    {
        // Parses the specified expense account and writes the records to the output file
        // as character separated values containing the fields:
        // 1. Date
        // 2. Amount
        // 3. Description
        // 4. Type
        public static String GenerateCSVFile(string documentPath, string outputFilePath)
        {
			// Open the output file for writing
			StreamWriter sw = new StreamWriter(outputFilePath, false);
            //Open the Word document
            Document document = null;
            try
            {
                document = new Document(documentPath);
            }
            catch
            {
                return "Unable to open the specified Word document!";
            }
            DocumentObject obj = document.ChildObjects.FirstItem;
            DocumentObjectType type = obj.DocumentObjectType;
            if (type == DocumentObjectType.Section)
            {
                Section section = (Section)obj;
                if (section.Tables.Count == 1)
                {
                    // We now have access to the expense table!
                    Table table = (Table)section.Tables[0];
                    // Declare the table variables for processing
                    TableRow tableRow;
                    TableCell tableCell;
                    Paragraph paragraph;
                    // Iterate through the rows in the table
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        tableRow = table.Rows[i];
                        // Get the date
                        tableCell = tableRow.Cells[0];
                        string date = tableCell.Paragraphs[0].Text;
                        // Get the expense amounts, descriptions, and types
                        tableCell = tableRow.Cells[1];
                        string expenseType = "";
                        string amount = "";
                        string description = "";
                        for (int j = 0; j < tableCell.Paragraphs.Count; j++)
                        {
                            paragraph = tableCell.Paragraphs[j];
                            string text = paragraph.Text;
                            // Check what type of entry this is
                            TextSelection textSelection = new TextSelection(paragraph, 0, text.Length - 1);
                            TextRange textRange = textSelection.GetAsOneRange();
                            CharacterFormat characterFormat = textRange.CharacterFormat;
                            Color highlightColor = characterFormat.HighlightColor;
                            // If it's Appa's expense (grey) or no expense (white), skip it
                            if (highlightColor.R == highlightColor.G && highlightColor.G == highlightColor.B)
                            {
                                continue;
                            }
                            // if it's a cash withdrawal (blue), skip it
                            else if (highlightColor.R == 0 && highlightColor.G == 0 && highlightColor.B != 0)
                            {
                                continue;
                            }
                            // if it's a cash deposit (green), record it
                            else if (highlightColor.R == 0 && highlightColor.G != 0 && highlightColor.B == 0)
                            {
                                expenseType = "Deposit";
                                int spaceIndex = text.IndexOf(' ');
                                amount = text.Substring(0, spaceIndex);
                                description = "Cash Deposit";
                            }
                            // if it's a cash refund (red), record it
                            else if (highlightColor.R != 0 && highlightColor.G == 0 && highlightColor.B == 0)
                            {
                                expenseType = "Refund";
                                int spaceIndex = text.IndexOf(' ');
                                amount = text.Substring(0, spaceIndex);
                                description = "Cash Refund";
                            }
                            // if it's a cash expense (pink), record it
                            else if (highlightColor.R != 0 && highlightColor.G == 0 && highlightColor.B != 0)
                            {
                                expenseType = "Cash";
                                int spaceIndex = text.IndexOf(' ');
                                amount = text.Substring(0, spaceIndex);
                                description = text.Substring(spaceIndex + 3);
                            }
                            // if it's a debit/credit expense (yellow), record it
                            else if (highlightColor.R != 0 && highlightColor.G != 0 && highlightColor.B == 0)
                            {
                                expenseType = "Debit/Credit";
                                int spaceIndex = text.IndexOf(' ');
                                amount = text.Substring(0, spaceIndex);
                                description = text.Substring(spaceIndex + 3);
                            }
                            // Write the completed record to a sequential text file
                            sw.WriteLine(date + " | " + amount + " | " + description + " | " + expenseType);
                        }
                    }
					sw.Close();
                    System.Diagnostics.Process.Start(outputFilePath);
                    return "CSV file generated successfully!";
                }
                else if (section.Tables.Count == 0)
                {
                    return "This document does not have any tables!";
                }
                else
                {
                    return "This document has more than one table!";
                }
            }
            else
            {
                return "The first item of this document is not a section!";
            }
        }
    }
}
