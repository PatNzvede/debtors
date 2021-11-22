
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TestingForms
{
    public partial class Form1 : Form
    {
        int rows = 0;
        public Form1()
        {
            InitializeComponent();
        }
        public static IEnumerable<DateTime> AllDatesInMonth(int year, int month)
        {
            int days = DateTime.DaysInMonth(year, month);
            for (int day = 1; day <= days; day++)
            {
                yield return new DateTime(year, month, day);
            }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            string Url = "https://clientportal.jse.co.za/_layouts/15/DownloadHandler.ashx?FileName=/YieldX/Derivatives/Docs_DMTM";

            foreach (DateTime date in AllDatesInMonth(2021, 4))
            {
                var Datetostring = date.Date.ToString("yyyyMMdd");
                var processnow = Datetostring + "_D_DAILY MTM REPORT.xls";             
                string downloadTo = $"C:\\Web\\{processnow}";
                var fileUrl = Url + "/" + processnow;
                WebRequest request = WebRequest.Create(new Uri(fileUrl));
                request.Method = "HEAD";
                using (WebResponse response = request.GetResponse())
                {
                    if (response != null)
                    {
                        string dest = @"C:\Web\ProcessedFolder\";
                        string destFile = Path.Combine(dest, Path.GetFileName(processnow));
                        if((!File.Exists(destFile)) &&(!File.Exists(downloadTo)))
                        {
                            using (WebClient webClient = new WebClient())
                            {
                                webClient.DownloadFile(fileUrl, downloadTo);
                            }
                        }
                    }
                }
            }
            MessageBox.Show("All files have been downloaded");
        }
        
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string con = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                con = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                con = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con1 = new OleDbConnection(con))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con);
                  rows =  oleAdpt.Fill(dtexcel);
                   
                  }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            return dtexcel;
        }

        
        private void button2_Click(object sender, EventArgs e)
        {
            var connectionString = ConfigurationManager.ConnectionStrings["DailyReports"].ConnectionString;

            var fileExt = ".xls";               
            DirectoryInfo dry = new DirectoryInfo("C:\\Web\\");
            foreach (var fn in dry.GetFiles("*.xls"))
            {
                if (fn.Length == 0)
                {
                    fn.Delete();
                }
            }
            foreach (var fn in dry.GetFiles("*.xls"))
            {
                try
                {
                    string va = Regex.Match(fn.ToString(), @"\d+").Value;
                    DataTable dtExcel = new DataTable();
                    var filePath = fn.FullName;
                    DataColumn FileDate = new DataColumn();
                    FileDate.DataType = typeof(System.DateTime);
                   
                    FileDate.ColumnName = "FileDate";
                    dtExcel.Columns.Add("FileDate", typeof(System.DateTime));
                    dtExcel.Columns.Add("Contract", typeof(System.String));
                    dtExcel.Columns.Add("ExpiryDate", typeof(System.DateTime));
                    dtExcel.Columns.Add("Classification", typeof(System.String));
                    dtExcel.Columns.Add("Strike", typeof(System.Decimal));
                    dtExcel.Columns.Add("CallPut", typeof(System.String));
                    dtExcel.Columns.Add("MTMYield", typeof(System.Decimal));
                    dtExcel.Columns.Add("MarkPrice", typeof(System.Decimal));
                    dtExcel.Columns.Add("SpotRate", typeof(System.Decimal));
                    dtExcel.Columns.Add("PreviousMTM", typeof(System.Decimal));
                    dtExcel.Columns.Add("PremiumOnOption", typeof(System.Decimal));
                    dtExcel.Columns.Add("PreviousPrice", typeof(System.Decimal));
                    dtExcel.Columns.Add("Volatility", typeof(System.Decimal));
                    dtExcel.Columns.Add("Delta", typeof(System.Decimal));
                    dtExcel.Columns.Add("DeltaValue", typeof(System.Decimal));
                    dtExcel.Columns.Add("ContractsTraded", typeof(System.Decimal));
                    dtExcel.Columns.Add("OpenInterest", typeof(System.Decimal));

                    dtExcel = ReadExcel(filePath, fileExt);
                    dgv1.Visible = true;
                    dgv1.DataSource = dtExcel;
                    
                    SqlConnection conn = new SqlConnection(connectionString);
                    conn.Open();
                    for (int i = 4; i < dgv1.Rows.Count-4; i++)
                    {
                        using (SqlBulkCopy sqlbc = new SqlBulkCopy(connectionString))
                        {                                                   
                            SqlCommand insertCommand = new SqlCommand("dbo.SP_UpdateTable", conn);
                            insertCommand.CommandType = CommandType.StoredProcedure;
                            insertCommand.Parameters.AddWithValue("@FileDate", va);
                            insertCommand.Parameters.AddWithValue("@Contract", dgv1.Rows[i].Cells[0].Value);
                            insertCommand.Parameters.AddWithValue("@ExpiryDate", dgv1.Rows[i].Cells[2].Value);
                            insertCommand.Parameters.AddWithValue("@Classification", dgv1.Rows[i].Cells[3].Value);
                            insertCommand.Parameters.AddWithValue("@Strike", dgv1.Rows[i].Cells[4].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@CallPut", dgv1.Rows[i].Cells[5].Value);
                        insertCommand.Parameters.AddWithValue("@MTMYield", dgv1.Rows[i].Cells[6].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@MarkPrice", dgv1.Rows[i].Cells[7].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@SpotRate", dgv1.Rows[i].Cells[8].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@PreviousMTM", dgv1.Rows[i].Cells[9].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@PreviousPrice", dgv1.Rows[i].Cells[10].Value.ToString().Replace(",", "")); 
                        insertCommand.Parameters.AddWithValue("@PremiumOnOption", dgv1.Rows[i].Cells[11].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@Volatility", dgv1.Rows[i].Cells[12].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@Delta", dgv1.Rows[i].Cells[13].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@DeltaValue", dgv1.Rows[i].Cells[14].Value.ToString().Replace(",",""));
                            insertCommand.Parameters.AddWithValue("@ContractsTraded", dgv1.Rows[i].Cells[15].Value.ToString().Replace(",", ""));
                        insertCommand.Parameters.AddWithValue("@OpenInterest", dgv1.Rows[i].Cells[16].Value.ToString().Replace(",",""));
                            insertCommand.CommandType = CommandType.StoredProcedure;                           
                            insertCommand.ExecuteNonQuery();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message.ToString());
                }
            }
            string dest = @"C:\Web\ProcessedFolder\";
            foreach (var filer in Directory.EnumerateFiles(@"C:\Web\"))
            {
                string destFile = Path.Combine(dest, Path.GetFileName(filer));
                if (!File.Exists(destFile) && filer.Length > 0)
                {
                    File.Move(filer, destFile);
                }
            }
            MessageBox.Show("All files have been processed and files moved to the processed folder");
        }
        }

    }
  
        
      



   
