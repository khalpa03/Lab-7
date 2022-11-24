using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;


namespace Lab_5
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text.Length!=0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "Создание word документа";
                sfd.FileName = "Word файл";
                if (sfd.ShowDialog() == DialogResult.Cancel)
                    return;
                string filename = sfd.FileName;
                var app = new Word.Application();
                app.Visible = true;
                var doc = app.Documents.Add();
                var r = doc.Range();
                r.Bold = 1;
                r.Text = textBox1.Text;
                doc.SaveAs(filename);
            }
            else
            {
                MessageBox.Show("Заполните строку!!!");
            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            string[] line= textBox2.Text.Split(' ');

            if (textBox2.Text.Length != 0)
            {
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "Создание Excel документа";
                sfd.FileName = "Excel файл";
                sfd.Filter = "xls files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                if (sfd.ShowDialog() == DialogResult.Cancel)
                    return;
                string filename = sfd.FileName;
                var ex = new Excel.Application();
                ex.Visible = true;
                ex.SheetsInNewWorkbook = 2;
                var workbook = ex.Workbooks.Add(Type.Missing);
                var sheet = (Excel.Worksheet)ex.Worksheets.get_Item(1);

                for (int j = 1; j <= line.Length; j++)
                    sheet.Cells[1, j] = String.Format(line[j - 1], 1, j);
                ex.Application.ActiveWorkbook.SaveAs(filename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                
            }
            else
            {
                MessageBox.Show("Заполните строку числами!!!");
            }
        }
    }
}
