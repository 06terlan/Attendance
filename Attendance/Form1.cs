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
using Attendance.Classes;

namespace Attendance
{
    public partial class Form1 : Form
    {
        Dictionary<int, DateTime[]> dict = null;
        Dictionary<int, User> allUsers = null;
        Dictionary<int, string> data = null;
        int maxIns = 0, maxOuts = 0;
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
            dict = new Dictionary<int, DateTime[]>();
            data = new Dictionary<int, string>();
            allUsers = new Dictionary<int, User>();

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

                            allUsers[row1Col0] = new User(row1Col0);
                        }
                    }
                }
            }
        }

        private void CreateReport()
        {
            maxIns = 0; maxOuts = 0;

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
                            if (tur == "RS-485/PCI-Panel1-R2" || tur == "RS-485/PCI-Panel1-R4" || tur == "RS-485/PCI-Panel4-HR-in" || tur == "Entry-office-door")
                            {
                                dict[sId][0] = DateTime.Compare(dict[sId][0], new DateTime()) == 0 || DateTime.Compare(dict[sId][0], time) > 0 ? time : dict[sId][0];

                                allUsers[sId].entered(time);
                                maxIns = Math.Max(maxIns, allUsers[sId].allIns);
                            }
                            else if (tur == "RS-485/PCI-Panel1-R1" || tur == "RS-485/PCI-Panel1-R3" || tur == "RS-485/PCI-Panel4-HR-out" || tur == "Exit-office-door")
                            {
                                dict[sId][1] = DateTime.Compare(dict[sId][1], new DateTime()) == 0 || DateTime.Compare(dict[sId][1], time) < 0 ? time : dict[sId][1];

                                allUsers[sId].exited(time);
                                maxOuts = Math.Max(maxOuts, allUsers[sId].allOuts);
                            }
                        }
                    }
                }
            }
        }

        private void writeDataToExcel()
        {
            var fileName = string.Format("{0}\\SWH_R.xlsx", Directory.GetCurrentDirectory());

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
                row--;
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
                ws.View.FreezePanes(2, 1);

                p.SaveAs(new FileInfo(fileName));
            }

        }

        private void writeDataToExcelExtended()
        {
            var fileName = string.Format("{0}\\SWH_EX.xlsx", Directory.GetCurrentDirectory());

            using (var p = new ExcelPackage())
            {
                var ws = p.Workbook.Worksheets.Add("Sheet1");
                int maxInOuts = Math.Max(maxIns, maxOuts);

                ws.Cells[1, 1].Value = "Name";
                ws.Cells[1, 2].Value = "ID";
                ws.Cells[1, 3].Value = "In";
                ws.Cells[1, 4].Value = "Out";

                ws.Cells["A1:D1"].Style.Font.Bold = true;
                ws.Cells["A1:D1"].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D1"].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D1"].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D1"].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;

                int row = 2, endRow, rrr = 1;
                string lastType;
                foreach (var kvp in allUsers)
                {
                    ws.Cells[row, 1].Value = data[kvp.Key];
                    ws.Cells[row, 2].Value = kvp.Key;
                    rrr++;
                    //if (kvp.Key != 15462) continue;

                    endRow = row;
                    lastType = "out";
                    foreach (var dd in kvp.Value.inOutType)
                    {
                        if (dd.Value == "in")
                        {
                            if (lastType == "in")
                            {
                                endRow++;
                            }
                            ws.Cells[endRow, 3].Value = dd.Key.ToString("dd.MM.yyyy H:mm:ss");
                        }
                        else if (dd.Value == "out")
                        {
                            ws.Cells[endRow, 4].Value = dd.Key.ToString("dd.MM.yyyy H:mm:ss");
                            endRow++;
                        }

                        lastType = dd.Value;
                    }

                    if (row != endRow)
                    {
                        if (lastType == "out") endRow--;
                        ws.Cells["A" + row + ":A" + endRow].Merge = true;
                        ws.Cells["B" + row + ":B" + endRow].Merge = true;

                        ws.Cells["A" + row + ":D" + endRow].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells["A" + row + ":D" + endRow].Style.Fill.BackgroundColor.SetColor(rrr%2 ==1 ? Color.LightGray : Color.White);
                    }
                    else
                    {
                        ws.Cells["A" + row + ":D" + row].Style.Fill.PatternType = ExcelFillStyle.Solid;
                        ws.Cells["A" + row + ":D" + row].Style.Fill.BackgroundColor.SetColor(rrr % 2 == 1 ? Color.LightGray : Color.White);
                    }

                    row = endRow + 1;
                }

                //style
                ws.Cells.Style.Font.Size = 14;
                ws.Cells.Style.Font.Name = "Calibri";
                ws.View.FreezePanes(2, 1);
                row--;
                ws.Cells["A1:D" + row].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                ws.Cells["A1:D" + row].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                ws.Column(1).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Column(2).Style.VerticalAlignment = ExcelVerticalAlignment.Center;
                ws.Column(1).AutoFit();
                ws.Column(2).AutoFit();
                ws.Column(3).AutoFit();
                ws.Column(4).AutoFit();

                p.SaveAs(new FileInfo(fileName));
            }

        }

        private void btn_report_extended_Click(object sender, EventArgs e)
        {
            getUsersFromExcel();
            CreateReport();
            writeDataToExcelExtended();
        }
    }
}
