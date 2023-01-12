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

        public int FldClient { get; set; }

        public int FldMatter { get; set; }

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
            //08/16/2022

            Cursor.Current = Cursors.WaitCursor;
            toolStripStatusLabel.Text = "Executing Data Fix";
            statusStrip.Refresh();
            UpdateStatus("Executing Data Fix...", 1, 20);
            Application.DoEvents();

            //see if any of these entries are on a prebill
            string s1 = @"SELECT  distinct PBESDPreBill
                      FROM [UnbilledExpense]
                      inner join PreBillExpSumDetail on UEBatch = PBESDUEBatch and UERecNbr = PBESDUERecNbr
                      where UEDate <= '08/16/2022'";
            DataSet ds = _jurisUtility.RecordsetFromSQL(s1);

            //no open prebills for these entries is there so move on
            if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
            {
                Cursor.Current = Cursors.WaitCursor;
                toolStripStatusLabel.Text = "Executing Data Fix";
                statusStrip.Refresh();
                UpdateStatus("Executing Data Fix...", 15, 20);
                Application.DoEvents();

                s1 = @"update UnbilledExpense set UEExpCd = 'CONV'
                    where UEDate <= '08/16/2022'";
                _jurisUtility.ExecuteNonQueryCommand(0, s1);

                s1 = @"update ExpBatchDetail set EBDExpCd = 'CONV'
                     where ebdid in (select ueid from UnbilledExpense
                  where UEDate <= '08/16/2022')";
                _jurisUtility.ExecuteNonQueryCommand(0, s1);

                s1 = @"update [ExpSumByPrd] set ESPEntered = 0.00";
                _jurisUtility.ExecuteNonQueryCommand(0, s1);

                s1 = @"update [ExpSumITD] set ESICurUnbilBal = 0.00";
                _jurisUtility.ExecuteNonQueryCommand(0, s1);

                ds.Clear();
                //get expsumbyprd items
                s1 = @"SELECT  [UEMatter] ,[UEPrdYear],[UEPrdNbr],[UEExpCd] ,sum([UEAmount]) as Amt
                      FROM [UnbilledExpense]
                      where UEDate <= '08/16/2022'
                      group by [UEMatter] ,[UEPrdYear],[UEPrdNbr],[UEExpCd]";
                ds = _jurisUtility.RecordsetFromSQL(s1);

                if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                {
                    //for some reason there were no records...
                }
                else
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        //does record exist? if yes, update, if not, create
                        s1 = @"SELECT  * FROM [ExpSumByPrd]
                        where ESPMatter = " + dr[0].ToString() + " and ESPPrdYear = " + dr[1].ToString() + " and ESPPrdNbr = " + dr[2].ToString() + " and ESPExpCd = '" + dr[3].ToString() + "'";
                        DataSet fs = _jurisUtility.RecordsetFromSQL(s1);
                        if (fs == null || fs.Tables.Count == 0 || fs.Tables[0].Rows.Count == 0)
                        {
                            //doesnt exist so insert
                            s1 = @"insert into [ExpSumByPrd] ([ESPMatter],[ESPExpCd],[ESPPrdYear] ,[ESPPrdNbr],[ESPEntered] ,[ESPBilledValue],[ESPBilledAmt],[ESPReceived],[ESPAdjusted])
                            values (" + dr[0].ToString() + ",'" + dr[3].ToString() + "'," + dr[1].ToString() + "," + dr[2].ToString() + "," + dr[4].ToString() + ",0.00,0.00,0.00,0.00)";
                            _jurisUtility.ExecuteNonQueryCommand(0, s1);
                        }
                        else
                        {
                            //does exist so update
                            s1 = @"update [ExpSumByPrd] set ESPEntered = " + dr[4].ToString() +
                            " where ESPMatter = " + dr[0].ToString() + " and ESPPrdYear = " + dr[1].ToString() + " and ESPPrdNbr = " + dr[2].ToString() + " and ESPExpCd = '" + dr[3].ToString() + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s1);

                        }
                    }
                }
                    //now do it for expsumitd
                    ds.Clear();
                    s1 = @"SELECT  [UEMatter] ,[UEExpCd] ,sum([UEAmount]) as Amt
                      FROM [UnbilledExpense]
                      where UEDate <= '08/16/2022'
                      group by [UEMatter] ,[UEExpCd]";
                    ds = _jurisUtility.RecordsetFromSQL(s1);

                if (ds == null || ds.Tables.Count == 0 || ds.Tables[0].Rows.Count == 0)
                {
                    //for some reason there were no records...
                }
                else
                {
                    foreach (DataRow dr in ds.Tables[0].Rows)
                    {
                        //does record exist? if yes, update, if not, create
                        s1 = @"SELECT  * FROM [ExpSumITD]
                        where ESPMatter = " + dr[0].ToString() + " and ESIExpCd = '" + dr[1].ToString() + "'";
                        DataSet fs = _jurisUtility.RecordsetFromSQL(s1);
                        if (fs == null || fs.Tables.Count == 0 || fs.Tables[0].Rows.Count == 0)
                        {
                            //doesnt exist so insert
                            s1 = @"insert into [ExpSumITD] ([ESIMatter] ,[ESIExpCd] ,[ESICurUnbilBal] ,[ESICurARBal],[ESIEntered],[ESIBilledValue] ,[ESIBilledAmt],[ESIReceived],[ESIAdjusted])
                            values (" + dr[0].ToString() + ",'" + dr[1].ToString() + "'," + dr[2].ToString() + ",0.00,0.00,0.00,0.00,0.00,0.00)";
                            _jurisUtility.ExecuteNonQueryCommand(0, s1);
                        }
                        else
                        {
                            //does exist so update
                            s1 = @"update [ExpSumITD] set ESICurUnbilBal = " + dr[2].ToString() +
                            " where ESIMatter = " + dr[0].ToString() + " and ESIExpCd = '" + dr[1].ToString() + "'";
                            _jurisUtility.ExecuteNonQueryCommand(0, s1);

                        }
                    }


                }



                Cursor.Current = Cursors.Default;
                toolStripStatusLabel.Text = "Data Fix - Complete";
                statusStrip.Refresh();
                UpdateStatus("Data Fix - Complete", 1, 1);
                WriteLog("Data Fix " + DateTime.Now.ToShortDateString());
                Application.DoEvents();







            }
            else //there are prebills so display pb numbers
            {
                string pbNumbers = "";
                foreach (DataRow dr in ds.Tables[0].Rows)
                    pbNumbers = pbNumbers + dr[0].ToString() + ",";
                pbNumbers = pbNumbers.TrimEnd(',');

                MessageBox.Show("The following PreBills must be removed:" + "\r\n" + pbNumbers + "\r\n" + "No changes were made to the database.");
                Cursor.Current = Cursors.Default;
                toolStripStatusLabel.Text = "Fix Interrupted";
                statusStrip.Refresh();
                UpdateStatus("Fix Interrupted", 10, 20);
                Application.DoEvents();

            }
                





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
                double pctLong = Math.Round(((double)step / steps) * 100.0);
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
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
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
            if (File.Exists(filePathName))
            {
                File.Copy(filePathName, filePathName + ".ark1");
                File.Delete(filePathName);
            }

        }



        private void LogFile(string LogLine)
        {
            string AppDir = Path.GetDirectoryName(Application.ExecutablePath);
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

        private void labelDescription_Click(object sender, EventArgs e)
        {

        }

     
    }
}
