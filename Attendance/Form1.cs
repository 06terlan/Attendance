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
using System.Runtime.InteropServices;
using System.IO;
using System.Data.OleDb;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace Attendance
{
    public partial class Form1 : Form
    {
        IDictionary<int, DateTime[]> dict = new Dictionary<int, DateTime[]>();
        IDictionary<int, string> data = new Dictionary<int, string>();
        DialogResult isFileSelcted = DialogResult.No;

        const string fileName = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void btn_file_Click(object sender, EventArgs e)
        {
            isFileSelcted = openFileDialog1.ShowDialog();
        }

        private void btn_report_Click(object sender, EventArgs e)
        {
            getUsersFromExcel();
            CreateReport();
            writeDataToExcel();
        }

        private void getUsersFromExcel()
        {
            var fileName = string.Format("{0}\\SWH.xls", Directory.GetCurrentDirectory());
            var connectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", fileName);

            using (OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();
                OleDbCommand command = new OleDbCommand("select ID,Name from [Sheet1$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string str = dr[0].ToString().Trim();
                        if (str != "")
                        {
                            int row1Col0 = int.Parse(str);
                            dict[row1Col0] = new DateTime[2];
                            data[row1Col0] = dr[1].ToString();
                        }
                    }
                }
            }
        }

        private void CreateReport()
        {
            string line;
            var fileName = isFileSelcted == DialogResult.OK ? openFileDialog1.FileName : string.Format("{0}\\History.txt", Directory.GetCurrentDirectory());

            System.IO.StreamReader file = new System.IO.StreamReader(fileName);
            while ((line = file.ReadLine()) != null)
            {
                string[] datas = line.Split(';');
                if (datas.Length >= 7)
                {
                    int sId = 0;
                    DateTime time = new DateTime();
                    string sTime = datas[0].Trim('\"');
                    string tur = datas[5].Trim('\"');

                    if (int.TryParse(datas[7].Trim('\"'), out sId) && sTime != "")
                    {
                        if (dict.ContainsKey(sId))
                        {
                            time = DateTime.Parse(sTime);
                            //in
                            if (tur == "RS-485/PCI-Panel1-R2" || tur == "RS-485/PCI-Panel1-R4" || tur == "RS-485/PCI-Panel4-HR-in")
                            {
                                dict[sId][0] = DateTime.Compare(dict[sId][0], new DateTime()) == 0 || DateTime.Compare(dict[sId][0], time) > 0 ? time : dict[sId][0];
                            }
                            else if (tur == "RS-485/PCI-Panel1-R1" || tur == "RS-485/PCI-Panel1-R3" || tur == "RS-485/PCI-Panel4-HR-out")
                            {
                                dict[sId][1] = DateTime.Compare(dict[sId][1], new DateTime()) == 0 || DateTime.Compare(dict[sId][1], time) < 0 ? time : dict[sId][1];
                            }
                        }
                    }
                }
            }
        }

        private void writeDataToExcel()
        {
            //Open the workbook (or create it if it doesn't exist)
            var fileName = string.Format("{0}\\SWH_"+(DateTime.Now.ToString("dd.MM.yyyy"))+".xlsx", Directory.GetCurrentDirectory());

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");

                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "ID";
                ws.Cells[1, 3].Value = "In";
                ws.Cells[1, 4].Value = "Out";
                ws.Cells["A1:D1"].Style.Font.Bold = true;

                int row = 2;
                foreach (var kvp in dict)
                {
                    ws.Cells[row, 1].Value = data[kvp.Key];
                    ws.Cells[row, 2].Value = kvp.Key;
                    ws.Cells[row, 3].Value = DateTime.Compare(kvp.Value[0], new DateTime()) == 0 ? "-" : kvp.Value[0].ToString("dd.MM.yyyy H:mm:ss");
                    ws.Cells[row, 4].Value = DateTime.Compare(kvp.Value[1], new DateTime()) == 0 ? "-" : kvp.Value[1].ToString("dd.MM.yyyy H:mm:ss");
                    row++;
                }

                //style
                ws.Cells["A1:D" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Font.Size = 14;
                ws.Cells["A1:D" + row].Style.Font.Name = "Calibri";

                ws.Column(1).AutoFit();
                ws.Column(2).AutoFit();
                ws.Column(3).AutoFit();
                ws.Column(4).AutoFit();

                p.SaveAs(new FileInfo(fileName));
            }

        }
    }
}
