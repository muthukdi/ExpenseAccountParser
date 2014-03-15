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
using System.Text.RegularExpressions;

namespace Dilip
{
    enum ParserResult
    {
        ParsingSuccessful,
        FileFormatError,
        SectionNotFoundError,
        TitleFormatError,
        MonthYearValuesRetrieved,
        TableNotFoundError,
        MultipleTablesError,
        TableRowCountError,
        TableColumnCountError,
        DateCellStructureError,
        ExpenseCellStructureError,
        DailyCellStructureError,
        LeftoverCellStructureError,
        TableStructureIntact,
        DateCellFormatError,
        ExpenseCellFormatError,
        DailyCellFormatError,
        LeftoverCellFormatError,
        TableDataIntact
    };

    enum ExpenseType
    {
        AppaExpense,
        None,
        Withdrawal,
        Deposit,
        Refund,
        Cash,
        DebitCredit,
        Unknown
    };

    class ExpenseAccountParser
    {
        static string month;
        static string year;
        static ParserResult result;
        // Parses the specified expense account and writes the records to the output file
        // as character separated values containing the fields: Date, Amount, Description, and Type
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
            if (obj.DocumentObjectType == DocumentObjectType.Section)
            {
                Section section = (Section)obj;
                result = extractMonthYearValues(section);
                if (result != ParserResult.MonthYearValuesRetrieved)
                {
                    return result;
                }
                // There should be exactly one table in this document!
                if (section.Tables.Count == 1)
                {
                    // We now have access to the expense table!
                    Table table = (Table)section.Tables[0];
                    result = CheckTableStructure(table);
                    if (result != ParserResult.TableStructureIntact)
                    {
                        return result;
                    }
                    result = CheckTableDataFormat(table);
                    if (result != ParserResult.TableDataIntact)
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

        // Extract the month and the year values from the title
        private static ParserResult extractMonthYearValues(Section section)
        {
            if (section.Paragraphs.Count > 1)
            {
                string text = section.Paragraphs[1].Text;
                string[] months = { "January", "February", "March", "April", "May", "June", 
                                         "July", "August", "September", "October", "November", "December" };
                foreach (string str in months)
                {
                    if (text.Contains(str))
                    {
                        month = str;
                        Console.WriteLine(month);
                        break;
                    }
                }
                if (month == null)
                {
                    return ParserResult.TitleFormatError;
                }
                for (int n = 1999; n < 2020; n++)
                {
                    if (text.Contains(Convert.ToString(n)))
                    {
                        year = Convert.ToString(n);
                        Console.WriteLine(year);
                        break;
                    }
                }
                if (year == null)
                {
                    return ParserResult.TitleFormatError;
                }
            }
            else
            {
                return ParserResult.TitleFormatError;
            }
            return ParserResult.MonthYearValuesRetrieved;
        }

        // Make sure that the table has the correct structure
        private static ParserResult CheckTableStructure(Table table)
        {
            // Wrong number of rows
            if (table.Rows.Count > 32 || table.Rows.Count < 29)
            {
                return ParserResult.TableRowCountError;
            }
            for (int i = 1; i < table.Rows.Count; i++)
            {
                // Wrong number of columns
                if (table.Rows[i].Cells.Count != 4)
                {
                    return ParserResult.TableColumnCountError;
                }
                // Must be exactly one entry in the date field
                if (table.Rows[i].Cells[0].Paragraphs.Count != 1)
                {
                    return ParserResult.DateCellStructureError;
                }
                // The entry in the date field must not be blank
                if (table.Rows[i].Cells[0].Paragraphs[0].Equals(""))
                {
                    return ParserResult.DateCellStructureError;
                }
                // Each entry in the expenses field must not be blank
                for (int j = 0; j < table.Rows[i].Cells[1].Paragraphs.Count; j++)
                {
                    if (table.Rows[i].Cells[1].Paragraphs[j].Equals(""))
                    {
                        return ParserResult.ExpenseCellStructureError;
                    }
                }
                // Must be exactly one entry in the daily field
                if (table.Rows[i].Cells[2].Paragraphs.Count != 1)
                {
                    return ParserResult.DailyCellStructureError;
                }
                // The entry in the daily field must not be blank
                if (table.Rows[i].Cells[2].Paragraphs[0].Equals(""))
                {
                    return ParserResult.DailyCellStructureError;
                }
                // Must be exactly one entry in the left-over field
                if (table.Rows[i].Cells[3].Paragraphs.Count != 1)
                {
                    return ParserResult.LeftoverCellFormatError;
                }
                // The entry in the left-over field must not be blank
                if (table.Rows[i].Cells[3].Paragraphs[0].Equals(""))
                {
                    return ParserResult.LeftoverCellStructureError;
                }
            }
            return ParserResult.TableStructureIntact;
        }

        // Make sure that the actual data in the table makes sense
        private static ParserResult CheckTableDataFormat(Table table)
        {
            Regex dateRegex = new Regex(month + " [0-9]{1,2}, " + year);
            Regex expenseRegex = new Regex("(\\$[0-9]+.[0-9]{2} )|(None)");
            Regex dailyRegex = new Regex("(\\$[0-9]+.[0-9]{2})|(None)");
            Regex leftoverRegex = new Regex("(\\$[0-9]+.[0-9]{2})|(None)");
            for (int i = 1; i < table.Rows.Count; i++)
            {
                string date = table.Rows[i].Cells[0].Paragraphs[0].Text;
                // Must be a date in the current month and year with the format (Month XX, Year)
                if (!dateRegex.IsMatch(date))
                {
                    Console.WriteLine(date);
                    return ParserResult.DateCellFormatError;
                }
                // Each entry in the expense field must have a valid format
                for (int j = 0; j < table.Rows[i].Cells[1].Paragraphs.Count; j++)
                {
                    string entry = table.Rows[i].Cells[1].Paragraphs[j].Text;
                    if (!expenseRegex.IsMatch(entry))
                    {
                        Console.WriteLine(entry);
                        return ParserResult.ExpenseCellFormatError;
                    }
                }
                string daily = table.Rows[i].Cells[2].Paragraphs[0].Text;
                // The daily field must be of the format $X.XX (or None)
                if (!dailyRegex.IsMatch(daily))
                {
                    Console.WriteLine(daily);
                    return ParserResult.DailyCellFormatError;
                }
                // The left-over field must also be of the format $X.XX (or None)
                string leftover = table.Rows[i].Cells[3].Paragraphs[0].Text;
                if (!leftoverRegex.IsMatch(leftover))
                {
                    Console.WriteLine(leftover);
                    return ParserResult.LeftoverCellFormatError;
                }
            }
            return ParserResult.TableDataIntact;
        }

        // Extract the field values from this table row into a list
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
                else if (et == ExpenseType.Unknown)
                {
                    expenseType = "Unknown";
                    int spaceIndex = paragraph.Text.IndexOf(' ');
                    amount = paragraph.Text.Substring(0, spaceIndex);
                    description = "Unknown Description";
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
            // If there is no expense (white) or if you left an expense entry unhighlighted
            if (highlightColor.R == 255 && highlightColor.G == 255 && highlightColor.B == 255)
            {
                if (paragraph.Text.Equals("None"))
                {
                    return ExpenseType.None;
                }
                else
                {
                    return ExpenseType.Unknown;
                }
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
