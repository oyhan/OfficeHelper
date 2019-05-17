using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Application = Microsoft.Office.Interop.Word.Application;

namespace OfficeHelper
{
    public partial class MainForm : Telerik.WinControls.UI.RadForm
    {
        public MainForm()
        {
            InitializeComponent();
            backgroundWorker1.WorkerReportsProgress = true;
            backgroundWorker1.WorkerSupportsCancellation = true;
            backgroundWorker1.DoWork += backgroundWorker1_DoWork;
            backgroundWorker1.ProgressChanged += BackgroundWorker1_ProgressChanged;
        }

        private void BackgroundWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
             radProgressBar1.Value1 =e.ProgressPercentage ;
            radProgressBar1.Text = $"{e.ProgressPercentage} %";
        }

        public string FilePath { get; set; }
        public Dictionary<string, string> Survay { get; set; }
        private void radButton1_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                FilePath =  openFileDialog1.FileName;
                MessageBox.Show("Ok press START now");
            }


        }

        private void radButton2_Click(object sender, EventArgs e)
        {


            if (backgroundWorker1.IsBusy != true)
            {
                // Start the asynchronous operation.
                backgroundWorker1.RunWorkerAsync();
            }
         
        }
        Microsoft.Office.Interop.Excel.Application oXL;
        Microsoft.Office.Interop.Excel._Workbook oWB;
        Microsoft.Office.Interop.Excel._Worksheet oSheet;
        Microsoft.Office.Interop.Excel.Range oRng;
        object misvalue = System.Reflection.Missing.Value;



        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            BackgroundWorker worker = sender as BackgroundWorker;
            //MessageBox.Show("Please be patient.It may take some time depending on your pc resources");
            Application application = new Application();
            Document document = application.Documents.Open(FilePath);
            Survay = new Dictionary<string, string>();
            var validList = new List<string>();

            radProgressBar1.Maximum = 100;

            radProgressBar1.Step = 1;
            radProgressBar1.Value1 = 0;
            var c = document.Paragraphs.Count;
            // Loop through all words in the document.
            try
            {
                for (int i = 1; i <= c; i++)
                {
                    if (worker.CancellationPending == true)
                    {
                        e.Cancel = true;
                        application.Quit();
                        return;
                        break;
                    }
                    if (document.Paragraphs[i].Range.Text != "\r")
                    {
                        validList.Add(document.Paragraphs[i].Range.Text.Replace("\r", ""));
                    }
                    worker.ReportProgress(i * 100 / c);


                }
                int count = validList.Count;
                for (int i = 0; i < count; i += 4)
                {
                    var index = 0;
                    for (int j = i; j < i + 4; j++)
                    {
                        //is number vote?
                        if (index == 2 && Regex.IsMatch(validList[j], @"^\d+$"))
                        {
                            if (!Survay.ContainsKey(validList[j - 1]))
                            {
                                Survay.Add(validList[j - 1], validList[j]);

                            }
                        }
                        index++;
                    }
                    // Write the word.
                }

                SaveToExcel(Survay);


                // Close word.
                application.Quit();
            }
            catch (Exception ex)
            {
                File.WriteAllText(AppDomain.CurrentDomain.BaseDirectory + "//errLog.txt", ex.ToString());
                application.Quit();

            }
            //for (int i = 1; i <= 10; i++)
            //{
            //    if (worker.CancellationPending == true)
            //    {
            //        e.Cancel = true;
            //        break;
            //    }
            //    else
            //    {
            //        // Perform a time consuming operation and report progress.
            //        System.Threading.Thread.Sleep(500);
            //    }
            //}
        }
        private void SaveToExcel(Dictionary<string, string> survay)
        {
            oXL = new Microsoft.Office.Interop.Excel.Application();
            oXL.Visible = true;
            oXL.UserControl = false;
            //Get a new workbook.
            oWB = (Microsoft.Office.Interop.Excel._Workbook)(oXL.Workbooks.Add(""));
            oSheet = (Microsoft.Office.Interop.Excel._Worksheet)oWB.ActiveSheet;
            oSheet.Cells[1, 1] = "User";
            oSheet.Cells[1, 2] = "Vote";

            var index = 0;
            foreach (var item in survay)
            {
                oSheet.Cells[index + 2, 1] = item.Key;
                oSheet.Cells[index + 2, 2] = item.Value;

                index++;


            }

            //Add table headers going cell by cell.



            //Format A1:D1 as bold, vertical alignment = center.
            //oSheet.get_Range("A1", "D1").Font.Bold = true;
            //oSheet.get_Range("A1", "D1").VerticalAlignment =
            //    Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            // Create an array to multiple values at once.
            //string[,] saNames = new string[5, 2];

            //saNames[0, 0] = "John";
            //saNames[0, 1] = "Smith";
            //saNames[1, 0] = "Tom";

            //saNames[4, 1] = "Johnson";

            ////Fill A2:B6 with an array of values (First and Last Names).
            //oSheet.get_Range("A2", "B6").Value2 = saNames;

            ////Fill C2:C6 with a relative formula (=A2 & " " & B2).
            //oRng = oSheet.get_Range("C2", "C6");
            //oRng.Formula = "=A2 & \" \" & B2";

            ////Fill D2:D6 with a formula(=RAND()*100000) and apply format.
            //oRng = oSheet.get_Range("D2", "D6");
            //oRng.Formula = "=RAND()*100000";
            //oRng.NumberFormat = "$0.00";

            ////AutoFit columns A:D.
            //oRng = oSheet.get_Range("A1", "D1");
            //oRng.EntireColumn.AutoFit();


            oWB.SaveAs("c:\\test505.xls", Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing,
                false, false, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

            oWB.Close();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {

        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (backgroundWorker1.WorkerSupportsCancellation == true)
            {
                // Cancel the asynchronous operation.
                backgroundWorker1.CancelAsync();
            }
        }
    }
}
