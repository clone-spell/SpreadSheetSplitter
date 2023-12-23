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
using OfficeOpenXml;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ProgressBar;

namespace SpreadSheetSplitter
{
    public partial class Form1 : Form
    {
        public string filePath = "";
        public string exportPath = "";
        public string worksheet = "";
        public int columnIndex = -1;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
            backgroundWorker.WorkerReportsProgress = true;
            backgroundWorker.WorkerSupportsCancellation = true;
        }

        private void btnChoseFile_Click(object sender, EventArgs e)
        {
            Color tempColur = btnChoseFile.ForeColor;
            string tempTxt = btnChoseFile.Text;
            btnChoseFile.ForeColor = SystemColors.WindowText;
            btnChoseFile.Text = string.Empty;


            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                cbSheet.Items.Clear();

                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                dialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm|All Files|*.*";

                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = dialog.FileName;
                    btnChoseFile.Text = Path.GetFileName(filePath);

                    //add sheet names to combobox
                    List<string> sheets = GetSheetNames(filePath);
                    foreach (string sheet in sheets)
                    {
                        cbSheet.Items.Add(sheet);
                    }
                }
                else
                {
                    btnChoseFile.Text = tempTxt;
                    btnChoseFile.ForeColor = tempColur;
                }

                cbColumn.Focus();
            }
        }

        private void cbSheet_SelectedIndexChanged(object sender, EventArgs e)
        {
            cbColumn.Items.Clear();

            List<string> colList = GetColumnNames(filePath, cbSheet.Text);
            foreach (string col in colList)
            {
                cbColumn.Items.Add(col);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            if(backgroundWorker.IsBusy)
                backgroundWorker.CancelAsync();

            btnChoseFile.Text = "Chose File";
            btnChoseFile.ForeColor = Color.Gray;

            cbSheet.Items.Clear();
            cbColumn.Items.Clear();

            filePath = string.Empty;
            exportPath = string.Empty;
            worksheet = string.Empty;
            columnIndex = -1;

            progressBar.Value = 0;
            btnSplit.Enabled = true;
            lblProgress.Text = string.Empty;
        }

        private void btnSplit_Click(object sender, EventArgs e)
        {
            if (!backgroundWorker.IsBusy)
            {
                //change variables
                worksheet = cbSheet.Text;
                columnIndex = cbColumn.SelectedIndex + 1;
                using (FolderBrowserDialog d = new FolderBrowserDialog())
                {
                    if (filePath == string.Empty)
                        return;

                    string fileName = Path.GetFileName(filePath);
                    d.SelectedPath = filePath.Substring(0, filePath.Length - fileName.Length);
                    if (d.ShowDialog() == DialogResult.OK)
                    {
                        exportPath = d.SelectedPath;
                    }
                    else
                    {
                        MessageBox.Show("Description : Select a valid export path", "Btn-Split-SelectFolderDialog", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                btnSplit.Enabled = false;



                backgroundWorker.RunWorkerAsync();
            }
        }



        //background worker
        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            List<string> distinctValues = GetDistinctColumnValues(filePath, worksheet, columnIndex);

            int tempProgress = 0;
            foreach (string value in distinctValues)
            {
                List<int> rowIndexsRead = GetMatchingRowIndexes(filePath, worksheet, columnIndex, value);

                if (rowIndexsRead.Any())
                {
                    var packageWrite = new ExcelPackage();

                    ExcelWorksheet worksheetWrite = packageWrite.Workbook.Worksheets.Add("Sheet1");

                    ExcelPackage packageRead = new ExcelPackage(new FileInfo(filePath));

                    ExcelWorksheet worksheetRead = packageRead.Workbook.Worksheets[worksheet];
                    List<int> colIndexesRead = new List<int>();
                    for (int i = 1; i <= worksheetRead.Dimension.End.Column; i++)
                    {
                        if (i != columnIndex)
                            colIndexesRead.Add(i);
                    }

                    worksheetWrite.Cells[1, 1].Value = "Variables";
                    worksheetWrite.Cells[1, 1].Style.Font.Bold = true;
                    worksheetWrite.Cells[2, 1].Value = "PAYMENT DATE RANGE";
                    worksheetWrite.Cells[2, 2].Value = $"{value} - {value}";

                    int colIndexWrite = 1;
                    foreach (int colIndexRead in colIndexesRead)
                    {
                        //set column names
                        worksheetWrite.Cells[4, colIndexWrite].Value = worksheetRead.Cells[1, colIndexRead].Value;
                        worksheetWrite.Cells[4, colIndexWrite].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheetWrite.Cells[4, colIndexWrite].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                        worksheetWrite.Cells[4, colIndexWrite].AutoFitColumns();
                        //set table values
                        int rowIndexWrite = 5;
                        foreach (int rowIndexRead in rowIndexsRead)
                        {
                            //cancel background worker
                            if (backgroundWorker.CancellationPending)
                            {
                                backgroundWorker.Dispose();
                            }

                            if (rowIndexWrite != rowIndexsRead.Count+4)
                            {
                                worksheetWrite.Cells[rowIndexWrite, 1].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                worksheetWrite.Cells[rowIndexWrite, 1].Style.Fill.BackgroundColor.SetColor(Color.LightGray);
                            }

                            worksheetWrite.Cells[rowIndexWrite, colIndexWrite].Value = worksheetRead.Cells[rowIndexRead, colIndexRead].Value;
                            rowIndexWrite++;
                        }
                        worksheetWrite.Cells[rowIndexsRead.Count + 4, colIndexWrite].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        worksheetWrite.Cells[rowIndexsRead.Count + 4, colIndexWrite].Style.Fill.BackgroundColor.SetColor(Color.Yellow);

                        colIndexWrite++;
                    }
                    string fileExportPath = Path.Combine(exportPath, $"{value}.xlsx");
                    packageWrite.SaveAs(new FileInfo(fileExportPath));
                }
                tempProgress++;

                System.Threading.Thread.Sleep(100);
                int progressPercentage = (int)((double)tempProgress / distinctValues.Count * 100);
                backgroundWorker.ReportProgress(progressPercentage);

            }
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBar.Value = e.ProgressPercentage;
            lblProgress.Text = $"Almost there {e.ProgressPercentage}% Done";
        }

        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MessageBox.Show($"Saved Into :\n\n{exportPath}", "Saved Successfuly", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }



        //functions
        public List<string> GetSheetNames(string filePath)
        {
            List<string> result = new List<string>();

            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    foreach (ExcelWorksheet excelWorksheet in package.Workbook.Worksheets)
                    {
                        result.Add(excelWorksheet.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Description : {ex.Message}", "Error in GetSheetNames()", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return result;
        }

        public List<string> GetColumnNames(string filePath, string sheetName)
        {
            List<string> result = new List<string>();
            try
            {
                using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                {
                    ExcelWorkbook workbook = package.Workbook;
                    ExcelWorksheet worksheet = workbook.Worksheets[sheetName];

                    for (int i = 1; i < worksheet.Dimension.End.Column; i++)
                    {
                        result.Add(worksheet.Cells[1, i].Text);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Description : {ex.Message}", "Error in GetColumnNames()", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return result;
        }

        public static List<string> GetDistinctColumnValues(string filePath, string sheetName, int colIndex)
        {
            List<string> distinctValues = new List<string>();

            // Load the Excel file
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // Get the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];


                // Iterate through the rows to get distinct values
                for (int row = 2; row < worksheet.Dimension.End.Row-2; row++)
                {
                    var cellValue = worksheet.Cells[row, colIndex].Text;
                    if (!distinctValues.Contains(cellValue))
                    {
                        distinctValues.Add(cellValue);
                    }
                }
            }

            return distinctValues;
        }

        public static List<int> GetMatchingRowIndexes(string filePath, string sheetName, int colIndex, string matchValue)
        {
            List<int> matchingIndexes = new List<int>();

            // Load the Excel file
            FileInfo fileInfo = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                // Get the worksheet
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];


                // Iterate through the rows to find matching values
                for (int row = 2; row <= worksheet.Dimension.End.Row-1; row++)
                {
                    var cellValue = worksheet.Cells[row, colIndex].Text;
                    if (cellValue.Equals(matchValue, StringComparison.OrdinalIgnoreCase))
                    {
                        matchingIndexes.Add(row);
                    }
                }
            }

            return matchingIndexes;
        }

    }
}
