using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Globalization;
using Gizmox.Controls;
using JDataEngine;
using JurisAuthenticator;
using JurisUtilityBase.Properties;
using System.Data.OleDb;
using Microsoft.Office.Interop.Excel;

namespace JurisUtilityBase
{
    public partial class UtilityBaseMain : Form
    {
        #region Private  members

        private JurisUtility _jurisUtility;

        #endregion

        #region Public properties

        public string CompanyCode { get; set; }

        public string JurisDbName { get; set; }

        public string JBillsDbName { get; set; }

        public List<Row> badRows = new List<Row>();

        string pathToExcelFile = "";

        Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

        #endregion

        #region Constructor

        public UtilityBaseMain()
        {
            InitializeComponent();
            _jurisUtility = new JurisUtility();

        }

        #endregion

        #region Public methods

        public void LoadCompanies()
        {
            var companies = _jurisUtility.Companies.Cast<object>().Cast<Instance>().ToList();
//            listBoxCompanies.SelectedIndexChanged -= listBoxCompanies_SelectedIndexChanged;
            listBoxCompanies.ValueMember = "Code";
            listBoxCompanies.DisplayMember = "Key";
            listBoxCompanies.DataSource = companies;
//            listBoxCompanies.SelectedIndexChanged += listBoxCompanies_SelectedIndexChanged;
            var defaultCompany = companies.FirstOrDefault(c => c.Default == Instance.JurisDefaultCompany.jdcJuris);
            if (companies.Count > 0)
            {
                listBoxCompanies.SelectedItem = defaultCompany ?? companies[0];
            }
        }

        #endregion

        #region MainForm events

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void listBoxCompanies_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (_jurisUtility.DbOpen)
            {
                _jurisUtility.CloseDatabase();
            }
            CompanyCode = "Company" + listBoxCompanies.SelectedValue;
            _jurisUtility.SetInstance(CompanyCode);
            JurisDbName = _jurisUtility.Company.DatabaseName;
            JBillsDbName = "JBills" + _jurisUtility.Company.Code;
            _jurisUtility.OpenDatabase();
            if (_jurisUtility.DbOpen)
            {
                ///GetFieldLengths();
            }

        }



        #endregion

        #region Private methods

        private void DoDaFix()
        {
            // Enter your SQL code here
            // To run a T-SQL statement with no results, int RecordsAffected = _jurisUtility.ExecuteNonQueryCommand(0, SQL);
            // To get an ADODB.Recordset, ADODB.Recordset myRS = _jurisUtility.RecordsetFromSQL(SQL);

            if (!string.IsNullOrEmpty(pathToExcelFile))
            {
                toolStripStatusLabel.Text = "Running. Please Wait...";
                xlApp.Visible = false;
                Workbook xlWorkbook = xlApp.Workbooks.Open(pathToExcelFile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                _Worksheet xlWorksheet = (_Worksheet)xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;
                Row currentRow = null;

                for (int a = 2; a < rowCount; a++)
                {
                    Microsoft.Office.Interop.Excel.Range range1 = xlWorksheet.Rows[a]; //For all columns in rows
                    //Range range1 = worksheet.Columns[1]; //for all rows in column 1

                    int col = 1;
                    currentRow = new Row(); //custom row class. I only use a few of the attributes but they are all programmed in the class if needed
                    foreach (Range r in range1.Cells) //range1.Cells represents all the columns/rows
                    {
                        currentRow.rowNumber = a;
                        if (col == 1)
                            currentRow.ID = Convert.ToString(r.Value);
                        else if (col == 2)
                            currentRow.matter = Convert.ToString(r.Value);
                        else if (col == 3)
                            currentRow.emp = Convert.ToString(r.Value);
                        else if (col == 4)
                            currentRow.tkpr = Convert.ToString(r.Value);
                        else if (col == 5)
                            currentRow.date = Convert.ToDateTime(r.Value);
                        else if (col == 6)
                            currentRow.hours = Convert.ToString(r.Value);
                        else if (col == 7)
                            currentRow.desc = Convert.ToString(r.Value);
                        else if (col == 8)
                            currentRow.oldTask = Convert.ToString(r.Value);
                        else if (col == 9)
                            currentRow.newTask = Convert.ToString(r.Value);
                        else if (col == 10)
                            currentRow.entryStatus = Convert.ToInt32(r.Value);
                        else if (col > 10)
                            break;
                        col++;
                    }

                    //if there IS an error (returns true)
                    if (anyErrorsInEidOrTaskCode(currentRow))
                    {
                        badRows.Add(currentRow);
                    }
                    else // no error so continue
                    {
                        //do sql stuff here
                        string SQL = "Update TimeEntry Set taskcode='" + currentRow.newTask.Trim() + "' where entryid=" + currentRow.ID.Trim();
                        _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                        SQL = "Update Timebatchdetail Set tbdtaskcd='" + currentRow.newTask.Trim() + "' where tbdid=(select tbdid from timeentrylink where entryid=" + currentRow.ID.Trim() + ")";
                        _jurisUtility.ExecuteNonQueryCommand(0, SQL);

                        SQL = "Update UnbilledTime Set uttaskcd='" + currentRow.newTask.Trim() + "' where utid=(select tbdid from timeentrylink where entryid=" + currentRow.ID.Trim() + ")";
                        _jurisUtility.ExecuteNonQueryCommand(0, SQL);
                    }
                }

                UpdateStatus("All Task codes updated.", 1, 1);
                toolStripStatusLabel.Text = "Status: Ready to Execute";

                if (badRows.Count == 0)
                    MessageBox.Show("The process is complete without error", "Confirmation", MessageBoxButtons.OK, MessageBoxIcon.None);
                else
                {
                    DialogResult rr = MessageBox.Show("The process is complete but there were" + "\r\n" + "errors. Would you like to see them?", "Errors", MessageBoxButtons.YesNo, MessageBoxIcon.None);
                    if (rr == DialogResult.Yes)
                    {
                        DataSet ds = displayErrors();
                        ReportDisplay rpds = new ReportDisplay(ds);
                        rpds.Show();
                    }
                }



                xlWorkbook.Close(false, Type.Missing, Type.Missing);
                xlApp.Quit();
            }
            else
                MessageBox.Show("Please browse to your Excel file first", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);

            
        }

        //returns false if EID exists in timeentry, timebatchdetail and unbilledtime table as well as the taskcode existing in taskcode, otherwise returns true
        //which means at least one of these tests were failed and they need to be fixed
        private bool anyErrorsInEidOrTaskCode(Row currRow)
        {
            DataSet ds1;
            //timeentry
            string SQL = "Select * from TimeEntry where entryid=" + currRow.ID.Trim();
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1.Tables[0].Rows.Count == 0)
            {
                currRow.error = "The EID is not in the TimeEntry Table";
                return true;
            }
            ds1.Clear();

            //timebatchdetail
            SQL = "Select * from Timebatchdetail where tbdid=(select tbdid from timeentrylink where entryid=" + currRow.ID.Trim() + ")";
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1.Tables[0].Rows.Count == 0)
            {
                currRow.error = "The EID is not in the TimeBatchDetail Table";
                return true;
            }
            ds1.Clear();

            //unbilledtime
            SQL = "Select * from UnbilledTime where utid=(select tbdid from timeentrylink where entryid=" + currRow.ID.Trim() + ")";
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1.Tables[0].Rows.Count == 0)
            {
                currRow.error = "The EID is not in the UnbilledTime Table";
                return true;
            }
            ds1.Clear();

            //taskcode
            SQL = "Select * from taskcode where TaskCdCode = '" + currRow.newTask.Trim() + "'";
            ds1 = _jurisUtility.ExecuteSqlCommand(0, SQL);
            if (ds1.Tables[0].Rows.Count == 0)
            {
                currRow.error = "The new task code is not valid";
                return true;
            }
            ds1.Clear();

            //only reachable if all of these sql queries return at least one row (they should all only return 1 row btw) :)
            return false;
        }



        private bool VerifyFirmName()
        {
            //    Dim SQL     As String
            //    Dim rsDB    As ADODB.Recordset
            //
            //    SQL = "SELECT CASE WHEN SpTxtValue LIKE '%firm name%' THEN 'Y' ELSE 'N' END AS Firm FROM SysParam WHERE SpName = 'FirmName'"
            //    Cmd.CommandText = SQL
            //    Set rsDB = Cmd.Execute
            //
            //    If rsDB!Firm = "Y" Then
            return true;
            //    Else
            //        VerifyFirmName = False
            //    End If

        }

        private bool FieldExistsInRS(DataSet ds, string fieldName)
        {

            foreach (DataColumn column in ds.Tables[0].Columns)
            {
                if (column.ColumnName.Equals(fieldName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }


        private static bool IsDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static bool IsNumeric(object Expression)
        {
            double retNum;

            bool isNum = Double.TryParse(Convert.ToString(Expression), System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out retNum);
            return isNum; 
        }

        private void WriteLog(string comment)
        {
            var sql =
                string.Format("Insert Into UtilityLog(ULTimeStamp,ULWkStaUser,ULComment) Values('{0}','{1}', '{2}')",
                    DateTime.Now, GetComputerAndUser(), comment);
            _jurisUtility.ExecuteNonQueryCommand(0, sql);
        }

        private string GetComputerAndUser()
        {
            var computerName = Environment.MachineName;
            var windowsIdentity = System.Security.Principal.WindowsIdentity.GetCurrent();
            var userName = (windowsIdentity != null) ? windowsIdentity.Name : "Unknown";
            return computerName + "/" + userName;
        }

        /// <summary>
        /// Update status bar (text to display and step number of total completed)
        /// </summary>
        /// <param name="status">status text to display</param>
        /// <param name="step">steps completed</param>
        /// <param name="steps">total steps to be done</param>
        private void UpdateStatus(string status, long step, long steps)
        {
            labelCurrentStatus.Text = status;

            if (steps == 0)
            {
                progressBar.Value = 0;
                labelPercentComplete.Text = string.Empty;
            }
            else
            {
                double pctLong = Math.Round(((double)step/steps)*100.0);
                int percentage = (int)Math.Round(pctLong, 0);
                if ((percentage < 0) || (percentage > 100))
                {
                    progressBar.Value = 0;
                    labelPercentComplete.Text = string.Empty;
                }
                else
                {
                    progressBar.Value = percentage;
                    labelPercentComplete.Text = string.Format("{0} percent complete", percentage);
                }
            }
        }

        private void DeleteLog()
        {
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            if (File.Exists(filePathName + ".ark5"))
            {
                File.Delete(filePathName + ".ark5");
            }
            if (File.Exists(filePathName + ".ark4"))
            {
                File.Copy(filePathName + ".ark4", filePathName + ".ark5");
                File.Delete(filePathName + ".ark4");
            }
            if (File.Exists(filePathName + ".ark3"))
            {
                File.Copy(filePathName + ".ark3", filePathName + ".ark4");
                File.Delete(filePathName + ".ark3");
            }
            if (File.Exists(filePathName + ".ark2"))
            {
                File.Copy(filePathName + ".ark2", filePathName + ".ark3");
                File.Delete(filePathName + ".ark2");
            }
            if (File.Exists(filePathName + ".ark1"))
            {
                File.Copy(filePathName + ".ark1", filePathName + ".ark2");
                File.Delete(filePathName + ".ark1");
            }
            if (File.Exists(filePathName ))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }

            

        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            string filePathName = Path.Combine(AppDir, "VoucherImportLog.txt");
            using (StreamWriter sw = File.AppendText(filePathName))
            {
                sw.WriteLine(LogLine);
            }	
        }
        #endregion

        private void button1_Click(object sender, EventArgs e)
        {
            DoDaFix();
        }

        private void buttonReport_Click(object sender, EventArgs e)
        {

            System.Environment.Exit(0);
          
        }

        private DataSet displayErrors()
        {
            DataSet ds = new DataSet();
            System.Data.DataTable errorTable = ds.Tables.Add("Errors");
            errorTable.Columns.Add("EID");
            errorTable.Columns.Add("Matsys");
            errorTable.Columns.Add("empsysnbr");
            errorTable.Columns.Add("tkpr");
            errorTable.Columns.Add("edate");
            errorTable.Columns.Add("hourswork");
            errorTable.Columns.Add("narrative");
            errorTable.Columns.Add("OLDTASK");
            errorTable.Columns.Add("NEWTASK");
            errorTable.Columns.Add("entrystatus");
            errorTable.Columns.Add("Error");
          
            foreach (Row r in badRows)
            {
                DataRow errorRow = ds.Tables["Errors"].NewRow();
                errorRow["EID"] = r.ID;
                errorRow["Matsys"] = r.matter;
                errorRow["empsysnbr"] = r.emp;
                errorRow["tkpr"] = r.tkpr;
                errorRow["edate"] = Convert.ToString(r.date);
                errorRow["hourswork"] = Convert.ToString(r.hours);
                errorRow["narrative"] = r.desc;
                errorRow["OLDTASK"] = r.oldTask;
                errorRow["NEWTASK"] = r.newTask;
                errorRow["entrystatus"] = Convert.ToString(r.entryStatus);
                errorRow["Error"] = r.error;
                ds.Tables["Errors"].Rows.Add(errorRow);
            }

            return ds;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(openFileDialog1.FileName).ToLower().Trim() == ".xls" || Path.GetExtension(openFileDialog1.FileName).ToLower().Trim() == ".xlsx")
                {
                    pathToExcelFile = openFileDialog1.FileName;
                    label2.Text = "File Chosen: " + Path.GetFileName(pathToExcelFile);

                }
                else
                    MessageBox.Show("Only valid Excel files can be seleced (.xls, .xlsx)", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            } 
        }




    }
}
