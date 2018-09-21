using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DemoEPPlus
{
    public partial class Form1 : Form
    {
        private string excelImportPath = string.Empty;
        private DataTable dtImport = new DataTable();

        public Form1()
        {
            InitializeComponent();
            openFileDialog1.InitialDirectory = @"C:\";
            openFileDialog1.Title = "Browse Text Files";
            openFileDialog1.DefaultExt = "xlsx";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            excelImportPath = openFileDialog1.FileName;


            if (System.IO.File.Exists(excelImportPath))
            {
                using (var pck = new OfficeOpenXml.ExcelPackage())
                {
                    using (var stream = File.OpenRead(excelImportPath))
                    {
                        pck.Load(stream);
                    }
                    var ws = pck.Workbook.Worksheets.First();

                    var startRow = 2; // EPPlus dòng đầu tiên tính từ 1 => dòng đầu là Hearder, dòng 2 bắt đầu dữ liệu cần import

                    for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++) // ws.Dimension.End.Row => đọc đến row cuối cùng có định dạng
                    {
                        var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column]; // ws.Dimension.End.Column => đọc đến col cuối cùng có định dạng 

                        var maHang = wsRow[rowNum, 2].Text;
                        if (!string.IsNullOrEmpty(maHang))
                        {
                            var tenHang = wsRow[rowNum, 3].Text;

                            var dVT = wsRow[rowNum, 4].Text;

                            dtImport.Rows.Add(rowNum - 1, maHang, tenHang, dVT);
                        }
                    }

                    dataGridView1.DataSource = dtImport;

                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string excelName = "TemplateExport";
            string extenstion = ".xlsx";

            string fileTemplate = Application.StartupPath + "\\Excels\\";

            string strFileAfterExport = Application.StartupPath + "\\Excels\\Data\\" + excelName + DateTime.Now.ToString("yyyyMMddHHmmss") + extenstion;
            string strTemplate = Application.StartupPath + "\\Excels\\" + excelName + extenstion;

            FileInfo template = new FileInfo(strTemplate);
            FileInfo fNewFile = new FileInfo(strFileAfterExport);

            var package = new ExcelPackage(template);

            var data = dataGridView1.DataSource as DataTable;
            if (data.Rows.Count > 0)
            {
                var workbook = package.Workbook;

                ExcelWorksheet worksheet = workbook.Worksheets.First();

                // Insert record thêm
                worksheet.InsertRow(8, data.Rows.Count - 3, 7);
                int row = 5;
                int Stt = 1;
                foreach (DataRow item in data.Rows)
                {
                    // Số thứ tự
                    worksheet.Cells[row, 1].Value = Stt;

                    // Mã hàng
                    worksheet.Cells[row, 2].Value = item[1];

                    // tên hàng
                    worksheet.Cells[row, 3].Value = item[2];

                    // dvt
                    worksheet.Cells[row, 4].Value = item[3];

                    Stt++;
                    row++;
                }

                //Add additional info here
                package.SaveAs(fNewFile);

                MessageBox.Show("Export thanh cong");
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dtImport.Columns.Add("STT");
            dtImport.Columns.Add("MaHang");
            dtImport.Columns.Add("TenHang");
            dtImport.Columns.Add("DVT");
        }
    }
}
