using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Dilip;

namespace TestParser
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnRun_Click(object sender, EventArgs e)
        {
            ParserResult result = Dilip.ExpenseAccountParser.GenerateCSVFile("..\\..\\AccountSample.doc", "output.txt");
            Console.WriteLine(result);
            if (result == ParserResult.ParsingSuccessful)
            {
                System.Diagnostics.Process.Start("output.txt");
            }
            /*for (int i = 77; i < 101; i++)
            {
                string documentPath = "C:\\Users\\dilip\\Desktop\\Data\\Account" + i + ".doc";
                string outputPath = "C:\\Users\\dilip\\Desktop\\Output\\output" + i + ".txt";
                ParserResult result = Dilip.ExpenseAccountParser.GenerateCSVFile(documentPath, outputPath);
                Console.WriteLine(result);
                Console.WriteLine();
            }*/
        }
    }
}
