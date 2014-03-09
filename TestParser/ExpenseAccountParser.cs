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
    enum ParserResult
    {
        ParsingSuccessful,
        FileFormatError,
        SectionNotFoundError,
        TableNotFoundError,
        MultipleTablesError,
        TableRowCountError,
        TableColumnCountError,
        DateCellFormatError,
        ExpenseCellFormatError,
        DailyCellFormatError,
        LeftoverCellFormatError,
        TableStructureIntact
    };

    enum ExpenseType
    {
        AppaExpense,
        None,
        Withdrawal,
        Deposit,
        Refund,
        Cash,
        DebitCredit
    };

    class ExpenseAccountParser
    {
        // Parses the specified expense account and writes the records to the output file
        // as character separated values containing the fields:
        // 1. Date
        // 2. Amount
        // 3. Description
        // 4. Type
        public static ParserResult GenerateCSVFile(string documentPath, string outputFilePath)
        {
            //Open the Word document
            Document document = null;
            try
            {
                document = new Document(documentPath);
            }
            catch
            {
                return ParserResult.FileFormatError;
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
                    ParserResult result = CheckTableStructure(table);
                    if (result != ParserResult.TableStructureIntact)
                    {
                        return result;
                    }
                    // Open the output file for writing
                    StreamWriter sw = new StreamWriter(outputFilePath, false);
                    // Process each row in the table and write the data to the file
                    for (int i = 1; i < table.Rows.Count; i++)
                    {
                        List<string> records = ParseTableRow(table.Rows[i]);
                        for (int j = 0; j < records.Count; j++)
                        {
                            sw.WriteLine(records[j]);
                        }
                    }
					sw.Close();
                    return ParserResult.ParsingSuccessful;
                }
                else if (section.Tables.Count == 0)
                {
                    return ParserResult.TableNotFoundError;
                }
                else
                {
                    return ParserResult.MultipleTablesError;
                }
            }
            else
            {
                return ParserResult.SectionNotFoundError;
            }
        }

        // Make sure that the table has the correct structure
        private static ParserResult CheckTableStructure(Table table)
        {
            if (table.Rows.Count > 32 || table.Rows.Count < 29)
            {
                return ParserResult.TableRowCountError;
            }
            for (int i = 1; i < table.Rows.Count; i++)
            {
                if (table.Rows[i].Cells.Count != 4)
                {
                    return ParserResult.TableColumnCountError;
                }
                else if (table.Rows[i].Cells[0].Paragraphs.Count != 1)
                {
                    return ParserResult.DateCellFormatError;
                }
                else if (table.Rows[i].Cells[1].Paragraphs.Count == 0)
                {
                    return ParserResult.ExpenseCellFormatError;
                }
                else if (table.Rows[i].Cells[2].Paragraphs.Count != 1)
                {
                    return ParserResult.DailyCellFormatError;
                }
                else if (table.Rows[i].Cells[3].Paragraphs.Count != 1)
                {
                    return ParserResult.DateCellFormatError;
                }
            }
            return ParserResult.TableStructureIntact;
        }

        private static List<string> ParseTableRow(TableRow tableRow)
        {
            // Declare the table variables for processing
            TableCell tableCell;
            Paragraph paragraph;
            // Get the date
            tableCell = tableRow.Cells[0];
            string date = tableCell.Paragraphs[0].Text;
            // Get the expense amounts, descriptions, and types
            tableCell = tableRow.Cells[1];
            string expenseType = "";
            string amount = "";
            string description = "";
            // Iterate through the list of expenses in this row
            List<string> records = new List<string>();
            for (int j = 0; j < tableCell.Paragraphs.Count; j++)
            {
                paragraph = tableCell.Paragraphs[j];
                ExpenseType et = GetExpenseType(paragraph);
                if (et == ExpenseType.None || et == ExpenseType.AppaExpense)
                {
                    // skip
                    continue;
                }
                else if (et == ExpenseType.Withdrawal)
                {
                    // skip
                    continue;
                }
                else if (et == ExpenseType.Deposit)
                {
                    expenseType = "Deposit";
                    int spaceIndex = paragraph.Text.IndexOf(' ');
                    amount = paragraph.Text.Substring(0, spaceIndex);
                    description = "Cash Deposit";
                }
                else if (et == ExpenseType.Refund)
                {
                    expenseType = "Refund";
                    int spaceIndex = paragraph.Text.IndexOf(' ');
                    amount = paragraph.Text.Substring(0, spaceIndex);
                    description = "Cash Refund";
                }
                else if (et == ExpenseType.Cash)
                {
                    expenseType = "Cash";
                    int spaceIndex = paragraph.Text.IndexOf(' ');
                    amount = paragraph.Text.Substring(0, spaceIndex);
                    description = paragraph.Text.Substring(spaceIndex + 3);
                }
                else if (et == ExpenseType.DebitCredit)
                {
                    expenseType = "Debit/Credit";
                    int spaceIndex = paragraph.Text.IndexOf(' ');
                    amount = paragraph.Text.Substring(0, spaceIndex);
                    description = paragraph.Text.Substring(spaceIndex + 3);
                }
                // Add this record to the list
                records.Add(date + " | " + amount + " | " + description + " | " + expenseType);
            }
            return records;
        }

        // Check what type of entry this is (cash, debit, refund, etc.)
        public static ExpenseType GetExpenseType(Paragraph paragraph)
        {
            TextSelection textSelection = new TextSelection(paragraph, 0, paragraph.Text.Length - 1);
            TextRange textRange = textSelection.GetAsOneRange();
            CharacterFormat characterFormat = textRange.CharacterFormat;
            Color highlightColor = characterFormat.HighlightColor;
            // If there is no expense (white)
            if (highlightColor.R == 255 && highlightColor.G == 255 && highlightColor.B == 255)
            {
                return ExpenseType.None;
            }
            // If it's Appa's expense (grey)
            else if (highlightColor.R == highlightColor.G && highlightColor.G == highlightColor.B)
            {
                return ExpenseType.AppaExpense;
            }
            // if it's a cash withdrawal (blue)
            else if (highlightColor.R == 0 && highlightColor.G == 0 && highlightColor.B != 0)
            {
                return ExpenseType.Withdrawal;
            }
            // if it's a cash deposit (green)
            else if (highlightColor.R == 0 && highlightColor.G != 0 && highlightColor.B == 0)
            {
                return ExpenseType.Deposit;
            }
            // if it's a cash refund (red)
            else if (highlightColor.R != 0 && highlightColor.G == 0 && highlightColor.B == 0)
            {
                return ExpenseType.Refund;
            }
            // if it's a cash expense (pink)
            else if (highlightColor.R != 0 && highlightColor.G == 0 && highlightColor.B != 0)
            {
                return ExpenseType.Cash;
            }
            // if it's a debit/credit expense (yellow)
            else if (highlightColor.R != 0 && highlightColor.G != 0 && highlightColor.B == 0)
            {
                return ExpenseType.DebitCredit;
            }
            else
            {
                // It should never get here!
                return ExpenseType.None;
            }
        }
    }
}
