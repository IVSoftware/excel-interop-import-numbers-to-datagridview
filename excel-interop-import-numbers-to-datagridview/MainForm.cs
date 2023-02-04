using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using Range = Microsoft.Office.Interop.Excel.Range;
using System.Diagnostics;

namespace excel_interop_import_numbers_to_datagridview
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
            // Create
            _xlApp = new Microsoft.Office.Interop.Excel.Application();
            // When in the future the main form closes, dispose the Excel interop.
            Disposed += (sender, e) =>
            {
                _xlBook?.Close();
                _xlApp.Quit();
            }; 
            buttonImport.Click += Import_Click_1;
        }
        private readonly Microsoft.Office.Interop.Excel.Application _xlApp;
        private Workbook _xlBook = null;

        protected override void OnLoad(EventArgs e)
        {
            base.OnLoad(e);

            dataGridView.DataSource = Records;

            #region F O R M A T    C O L U M N S
            Records.Add(new Record());
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                column.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
                if (column.Index > 1) column.DefaultCellStyle.Format = "F2";
            }
            Records.Clear();
            #endregion F O R M A T    C O L U M N S          
        }
        BindingList<Record> Records { get; } = new BindingList<Record>();

        private void Import_Click_1(object sender, EventArgs e)
        {
            Records.Clear();
            string filePath = Path.Combine(
                AppDomain.CurrentDomain.BaseDirectory,
                "Excel",
                "testdata.xlsx");

            _xlBook = _xlApp.Workbooks.Open(filePath);
            Worksheet xlSheet = _xlBook.Sheets[1];
            Range xlRange = xlSheet.UsedRange;

            Range range;
            List<string> 
                headers = new List<string>(),
                line = new List<string>();

            for (int i = 1; i <= xlRange.Rows.Count; i++)
            {
                if (i.Equals(1))
                {
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {
                        range = xlRange.Cells[i, j];
                        headers.Add(range.Value2);
                    }
                }
                else
                {
                    var record = new Record();
                    for (int j = 1; j <= xlRange.Columns.Count; j++)
                    {
                        range = xlRange.Cells[i, j];
                        var name = headers[j - 1];
                        switch(name)
                        {
                            case "Datum": record.Datum = DateTime.FromOADate(range.Value2); break;
                            case "Energia": record.Energia = range.Value2; break;
                            case "AC výkon": record.ACvýkon = range.Value2; break;
                            case "napetie siete": record.napetiesiete = range.Value2; break;
                            case "AC prud": record.ACprud = range.Value2; break;
                            case "DC napetie": record.DCnapetie = range.Value2; break;
                            default:
                                Debug.Assert(false, $"Not recognized: '{name}'");
                                break;
                        }
                    }
                    Records.Add(record);
                }
            }
        }
    }

    class Record
    {
        [DisplayName("Datum")]
        public DateTime Datum { get; set; }

        [DisplayName("Energia")]
        public double Energia { get; set; }

        [DisplayName("AC výkon")]
        public double ACvýkon { get; set; }

        [DisplayName("napetie siete")]
        public double napetiesiete { get; set; }

        [DisplayName("AC prud")]
        public double ACprud { get; set; }

        [DisplayName("DC napetie")]
        public double DCnapetie { get; set; }
    }
}
