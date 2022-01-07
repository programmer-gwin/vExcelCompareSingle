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
using System.IO;
using System.Text.RegularExpressions;
using System.Net.Http;
using Newtonsoft.Json;

namespace vExcelCompare
{
    public partial class Form1 : Form
    {
        DataTable src1, src2;
        bool IsMakeSameLength;
        DataTable src3 = null;
        int i;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                SetDefaultTextColor();

                src3 = new DataTable();
                i = 0;

                SetupForTable1();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error " + ex.Message);
            }
        }

        private async void SetupForTable1()
        {
           // int index = 0;
            foreach (DataRow row1 in src1.Rows)
            {
                string row1Data = row1.ItemArray[0].ToString();
                //await CallAPIforResponse("2033JUS2");
                await CallAPIforResponse(row1Data);
               // index++;
            }
            sucess = 0; failed = 0;

            if (NotSuccessfulTID.Count != 0)
            {
                foreach (string row1Data in NotSuccessfulTID)
                {
                    await CallAPIforResponse(row1Data);
                }
            }
            sucess = 0; failed = 0;
        }

        List<string> NotSuccessfulTID = new List<string>();
        int sucess = 0; int failed = 0;
        private async Task CallAPIforResponse(string row1Data)
        {
            //https://capriconiosapi.elkanahtech.com/MPOSAPI/DownloadParameterByTID?TID=2033JYO9
            var modelV = new { TID = row1Data };
            string json = JsonConvert.SerializeObject(modelV);
            string res = await PostAPI("DownloadParameterByTID", json);
            if (string.IsNullOrEmpty(res) || res.Contains("{\"ResponseCode\": \"00")) { NotSuccessfulTID.Add(row1Data); failed += 1; }
            else sucess += 1;

            lblTotalFailed.Text = "Failed: " + failed.ToString();
            lblTotalSuccess.Text = "Success: " + sucess.ToString();
        }

        public static async Task<string> PostAPI(string actionName, string mrawData)
        {
            string mresult = "";
            try
            {
                using (var mclient = new HttpClient() { BaseAddress = new Uri("https://capriconiosapi.elkanahtech.com/MPOSAPI/") })
                {
                    // mclient.Timeout = TimeSpan.FromMinutes(1);
                    var response = await mclient.PostAsync($"{actionName}/", new StringContent(mrawData, Encoding.UTF8, "application/json")).ConfigureAwait(false);
                    if (response.IsSuccessStatusCode)
                    {
                        var result = await response.Content.ReadAsStringAsync();
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                string msg = ex.Message;
                mresult = "";
            }
            return mresult;
        }

        private void SetDefaultTextColor()
        {
            for (int a = 0; a < dataGridView1.RowCount - 1; a++)
            {
                dataGridView1.Rows[a].DefaultCellStyle.ForeColor = Color.Black;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            dataGridView1.AutoResizeColumns();
        }

        string fileName = "";
        private void button2_Click(object sender, EventArgs e)
        {
            LoadExcel(1);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LoadExcel(2);
        }

        private void LoadExcel(int p)
        {
            try
            {
                OpenFileDialog openFileDialog1 = new OpenFileDialog();  //create openfileDialog Object
                openFileDialog1.Filter = "XML Files (*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb) |*.xml; *.xls; *.xlsx; *.xlsm; *.xlsb";//open file format define Excel Files(.xls)|*.xls| Excel Files(.xlsx)|*.xlsx| 
                openFileDialog1.FilterIndex = 3;

                openFileDialog1.Multiselect = false;        //not allow multiline selection at the file selection level
                openFileDialog1.Title = "Open Text File-R13";   //define the name of openfileDialog
                openFileDialog1.InitialDirectory = @"Desktop"; //define the initial directory

                if (openFileDialog1.ShowDialog() == DialogResult.OK)        //executing when file open
                {
                    string pathName = openFileDialog1.FileName;
                    fileName = System.IO.Path.GetFileNameWithoutExtension(openFileDialog1.FileName);
                    string sheetName = "Sheet1";
                    if (p == 1)
                    {
                        src1 = new DataTable();
                        sheetName = textBox1.Text;
                    }

                    string strConn = string.Empty;
                    FileInfo file = new FileInfo(pathName);
                    if (!file.Exists) { throw new Exception("Error, file doesn't exists!"); }
                    string extension = file.Extension;
                    switch (extension)
                    {
                        case ".xls":
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                        case ".xlsx":
                            strConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + pathName + ";Extended Properties='Excel 12.0;HDR=Yes;IMEX=1;'";
                            break;
                        default:
                            strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + pathName + ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;'";
                            break;
                    }
                    OleDbConnection cnnxls = new OleDbConnection(strConn);
                    OleDbDataAdapter oda = new OleDbDataAdapter(string.Format("select * from [{0}$]", sheetName), cnnxls);
                    if (p == 1)
                    {
                        oda.Fill(src1);
                        dataGridView1.DataSource = src1;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error!" + ex.Message);
            }
        }

        private void lblTotalFailed_Click(object sender, EventArgs e)
        {

        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

    }
}
