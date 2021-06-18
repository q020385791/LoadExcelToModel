using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace LoadExcelToModel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnLoad_Click(object sender, EventArgs e)
        {
            string FilePath = "";
            object miss = System.Reflection.Missing.Value;
            OpenFileDialog FileDialog = new OpenFileDialog();

            FileDialog.Title = "選擇文件";
            FileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";

            if (FileDialog.ShowDialog() == DialogResult.OK)
            {
                FilePath = FileDialog.FileName;
            }
            else
            {
             
                return;
            }
            Excel.Application oExcel = new Excel.Application();
            oExcel.UserControl = true;
            oExcel.DisplayAlerts = false;
            Excel.Range range = null;
            Excel.Workbook WorkBook = oExcel.Workbooks.Open(FilePath, miss, miss, miss, miss,
                                             miss, miss, miss, miss,
                                             miss, miss, miss, miss,
                                             miss, miss);
            Excel.Worksheet Sheet = oExcel.Application.Worksheets["ListName"];
            range = Sheet.UsedRange;

            for (int i = 2; i <= range.Rows.Count; i++)
            {
               string test= Sheet.Cells[i, 1].Text;
                //DO something you want

            }
        }
    }
}
