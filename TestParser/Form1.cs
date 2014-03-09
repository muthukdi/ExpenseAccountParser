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
        }
    }
}
