using System;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using DataTable = System.Data.DataTable;
using Excel =Microsoft.Office.Interop.Excel;

namespace Xl_Remove_shape_1._0.XlRemoveShape
{
    

    public partial class Form1 : Form
    {
        private readonly Excel.Application _xl = new Excel.Application();
        private Excel.Workbook _wb;
        private DataTable dt;
        public Form1()
        {
           //Below Code can be enabled for Machine Specific Program Execution
           //if (Environment.MachineName == "xxxx") //Enter your PC Name for proper functioning of the program,
           //     InitializeComponent();
           // else
           //{
           //     CloseForm(null, null);
           //     MessageBox.Show(@"Program has encountered unknown error and closing",caption: @"Error");
           // }
            InitializeComponent();
            dt = new DataTable();
            var dc1 = new DataColumn("Sheet Name", typeof(string));
            var dc2 = new DataColumn("No of shapes", typeof(int));
            dt.Columns.Add(dc1);
            dt.Columns.Add(dc2);
            
            FormClosing += CloseForm;

        }

        private void CloseForm(object sender, FormClosingEventArgs e)
        {
            if(_wb!=null)
                _wb.Close(false, Type.Missing, Type.Missing);
            if (_xl != null)
                _xl.Quit();
            GC.Collect();
            if (_wb != null) Marshal.FinalReleaseComObject(_wb);
            if (_xl != null) Marshal.FinalReleaseComObject(_xl);
            
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox1.Text = openFileDialog1.FileName; 
        }

        private void button2_Click(object sender, EventArgs e)
        {
            _wb = _xl.Workbooks.Open(textBox1.Text);
            if(dt.Rows.Count !=0)
                dt.Rows.Clear();
            foreach (Excel.Worksheet ws in _wb.Worksheets)
                dt.Rows.Add(ws.Name, ws.Shapes.Count);
            dataGridView1.DataSource = dt;
            dataGridView1.Columns.Cast<DataGridViewColumn>().ToList().ForEach(f => f.SortMode = DataGridViewColumnSortMode.NotSortable);

        }

        private void button3_Click(object sender, EventArgs e)
        {
            
            foreach (Excel.Worksheet ws in _wb.Worksheets)
            {
                var i = 0;
                var c = ws.Shapes.Count;
                foreach (Excel.Shape sp in ws.Shapes)
                {
                    sp.Delete();
                    i++;
                    dataGridView1[1,ws.Index-1].Value = c - i;
                }
            }
            
            _wb.SaveAs(textBox1.Text);
        }

       
    }
}
