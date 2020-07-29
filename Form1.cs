using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
 
namespace LogReader
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
         
    

        private System.Data.DataTable read(string filepath , int ColumnCount , int RowCount)
        {
            System.Data.DataTable dt = new System.Data.DataTable();

            try
            {
            
                //create a instance for the Excel object  
                Excel.Application oExcel = new Excel.Application();

                //pass that to workbook object  
                Excel.Workbook WB = oExcel.Workbooks.Open(filepath);

                // statement get the workbookname  
                string ExcelWorkbookname = WB.Name;

                // statement get the worksheet count  
                int worksheetcount = WB.Worksheets.Count;

                Excel.Worksheet wks = (Excel.Worksheet)WB.Worksheets[1];

                // statement get the firstworksheetname  

                string firstworksheetname = wks.Name;

                //statement get the first cell value  

                //Load Columns Header
                for(int c=1; c < ColumnCount; c++)
                {
                    var colData = ((Excel.Range)wks.Cells[1, c]).Value;
                    dt.Columns.Add(colData);
                }
                 
                //Load Rows
                for (int r = 2; r < RowCount; r++)
                {
                    for (int c = 1; c < ColumnCount; c++)
                    {
                        dt.Rows.Add();
                        var rawData = ((Excel.Range)wks.Cells[r, c]).Value;
                        dt.Rows[r-2][c-1] = rawData;
                    }
                }
            }
            catch (Exception ex)
            {
                string error = ex.Message;
            }

            return dt;
        }


        private void btn_Browse_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            if(ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    System.Data.DataTable dt = read(ofd.FileName, 10, 10);
                    dataGridView1.DataSource = dt;
                }
                catch(Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
        }
    }
}
