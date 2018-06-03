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

namespace Attendance
{
    public partial class Form1 : Form
    {
        IDictionary<int, DateTime[]> dict = new Dictionary<int, DateTime[]>();
        IDictionary<int, bool> dictChk = new Dictionary<int, bool>();
        const string fileName = "";

        public Form1()
        {
            InitializeComponent();
        }

        private void btn_file_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
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
                OleDbCommand command = new OleDbCommand("select ID from [Sheet1$]", connection);
                using (OleDbDataReader dr = command.ExecuteReader())
                {
                    while (dr.Read())
                    {
                        string str = dr[0].ToString().Trim();
                        if (str != "")
                        {
                            int row1Col0 = int.Parse(str);
                            dict[row1Col0] = new DateTime[2];
                            dictChk[row1Col0] = false;
                        }
                    }
                }
            }
        }

        private void CreateReport()
        {
            string line;
            var fileName = string.Format("{0}\\History.txt", Directory.GetCurrentDirectory());

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
                            dictChk[sId] = true;
                        }
                    }
                }
            }
        }

        private void writeDataToExcel()
        {
            var fileName = string.Format("{0}\\SWH.xls", Directory.GetCurrentDirectory());

            Excel.Workbook MyBook = null;
            Excel.Application MyApp = null;
            Excel.Worksheet MySheet = null;

            MyApp = new Excel.Application();
            MyApp.Visible = false;
            MyBook = MyApp.Workbooks.Open(fileName);
            MySheet = MyBook.Sheets[1];
            Excel.Range usedRange = MySheet.UsedRange;
            int countRows = usedRange.Rows.Count;


            MyApp.Workbooks.Close();
        }
    }
}
