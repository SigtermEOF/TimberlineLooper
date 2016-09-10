//****************************
//#define dev
//****************************

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Threading;
using System.IO;
using System.Xml.Linq;
using System.Xml;
using System.Collections;
using System.Net;

namespace TimberlineLooper
{
    public partial class TimberlineLooper_Main : Form
    {
        // Global variables.

        TaskMethods TM = null;
        Email EM = null;

        string sTimberlineConnString = TimberlineLooper.Properties.Settings.Default.TimberlineConnString.ToString();
        string sDP2ConnString = TimberlineLooper.Properties.Settings.Default.DP2ConnString.ToString();
        string sCDSConnString = TimberlineLooper.Properties.Settings.Default.CDSConnString.ToString();

        string sEmailServer = string.Empty;
        string sMyEmail = string.Empty;
        string sEmailBCC = string.Empty;
        
        bool bCreated = false;
        DataTable dtResults = new DataTable();
        DataTable dtScanned = new DataTable();
        DataTable dtScannedAddInfo = new DataTable();
        DataTable dtFinal = new DataTable();
        bool bReturned = false;
        bool bRenderError = false;
        bool bAutoGenFileError = false;
        bool bXMLFileError = false;
        StringBuilder sBStatus = new StringBuilder();
        bool bBatchTimer = false;
        bool bNoGatherData = false;
        bool bGotCode = false;
        
        public TimberlineLooper_Main()
        {
            InitializeComponent();
            TM = new TaskMethods();
            EM = new Email();
        }

        #region Form events.

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DoWork();
        }

        private void chkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBox1.Checked == false && chkBox2.Checked == false)
            {
                this.btnDoWork.Enabled = false;
            }
            else if (chkBox1.Checked == true || chkBox2.Checked == true)
            {
                this.btnDoWork.Enabled = true;
            }
        }

        private void chkBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (chkBox1.Checked == false && chkBox2.Checked == false)
            {
                this.btnDoWork.Enabled = false;
            }
            else if (chkBox1.Checked == true || chkBox2.Checked == true)
            {
                this.btnDoWork.Enabled = true;
            }
        }

        private void rTxtBoxStatus_TextChanged(object sender, EventArgs e)
        {
            rTxtBoxStatus.SelectionStart = rTxtBoxStatus.Text.Length;
            rTxtBoxStatus.ScrollToCaret();
            rTxtBoxStatus.Refresh();
        }

        private void tmrGather_Tick(object sender, EventArgs e)
        {
            this.DoWork();
        }

        private void tmrProcess_Tick(object sender, EventArgs e)
        {
            this.DoWork();
        }

        private void tmrBatch_Tick(object sender, EventArgs e)
        {
            this.DoWork();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        #endregion

        #region Work assignment.

        private void DoWork()
        {
            this.TaskMethods();

            DateTime dTimeNowStart = DateTime.Now;
            string sDTStart = DateTime.Now.ToString("HH:mm:ss");

            this.btnDoWork.Enabled = false;

            if (chkBox1.Checked == true && chkBox2.Checked == true)
            {
                this.StatusText.Text = "Status: [Current batch cycle started: " + dTimeNowStart + "]";
                Application.DoEvents();
                this.tmrBatch.Stop();
                this.tmrBatch.Enabled = false;
                this.tmrGather.Stop();
                this.tmrGather.Enabled = false;
                this.tmrProcess.Stop();
                this.tmrProcess.Enabled = false;
                bBatchTimer = true;
                this.Gather(ref sDTStart);
                this.Clear();
                this.Process(ref sDTStart);
                this.Clear();
            }
            if (chkBox1.Checked == false && chkBox2.Checked == true)
            {
                this.StatusText.Text = "Status: [Current processing cycle started: " + dTimeNowStart + "]";
                Application.DoEvents();
                this.tmrBatch.Stop();
                this.tmrBatch.Enabled = false;
                this.tmrGather.Stop();
                this.tmrGather.Enabled = false;
                this.tmrProcess.Stop();
                this.tmrProcess.Enabled = false;
                bBatchTimer = false;
                this.Process(ref sDTStart);
                this.Clear();
            }
            if (chkBox1.Checked == true && chkBox2.Checked == false)
            {
                this.StatusText.Text = "Status: [Current gathering cycle started: " + dTimeNowStart + "]";
                Application.DoEvents();
                this.tmrBatch.Stop();
                this.tmrBatch.Enabled = false;
                this.tmrGather.Stop();
                this.tmrGather.Enabled = false;
                this.tmrProcess.Stop();
                this.tmrProcess.Enabled = false;
                bBatchTimer = false;
                this.Gather(ref sDTStart);
                this.Clear();
            }
        }

        private void TaskMethods()  
        {
            this.Clear();

            TM.EmailVariables(ref sEmailServer, ref sMyEmail, ref sEmailBCC);

            #region Check for exceptions. 

            string sMySubject = string.Empty;
            string sMyBody = string.Empty;

            TM.CheckForExceptions(ref sMySubject, ref sMyBody);

            if (sMySubject != string.Empty || sMyBody != string.Empty)
            {
                EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);
            }            
            
            #endregion

            #region Check Order Items for errors. 

            sMySubject = string.Empty;
            sMyBody = string.Empty;

            TM.CheckOrderItemsForErrors(ref sMySubject, ref sMyBody);

            if (sMySubject != string.Empty || sMyBody != string.Empty)
            {
                EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);
            }            

            #endregion

            #region Checked for shipped XMLs. 

            sMySubject = string.Empty;
            sMyBody = string.Empty;

            TM.CheckForShippedXML(ref sMySubject, ref sMyBody);

            if (sMySubject != string.Empty || sMyBody != string.Empty)
            {
                EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);
            }            

            #endregion
        }

        private void Gather(ref string sDTStart) // Work data gathering method. Gathering and saving of job and frame data. 
        {
            //this.LoopThroughVerifiedTable();
            this.SearchItems();
            if (bNoGatherData != true)
            {
                this.ExistsInDP2();
                this.SaveResultsToVerifiedTable();
                this.LoopThroughVerifiedTable();
            }
            else if (bNoGatherData == true)
            {
                return;
            }

            string sStatDTNow = string.Empty;
            string sStatTxt = string.Empty;
            string sNextCycle = string.Empty;

            if (bBatchTimer != true)
            {
                DateTime dtime = DateTime.Now;
                string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
                string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
                sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
                sStatTxt = sStatDTNow + " This gathering cycle has completed." + Environment.NewLine + sStatDTNow + " Idle for 15 minutes." + Environment.NewLine + Environment.NewLine;
                sNextCycle = dtime.AddMinutes(15).ToString("HH:mm:ss");

                DateTime dTimeDoWorkEnd = DateTime.Now;
                string sDTEnd = dTimeDoWorkEnd.ToString("HH:mm:ss");
                TimeSpan tSpanDuration = DateTime.Parse(sDTEnd).Subtract(DateTime.Parse(sDTStart));
                this.StatusText.Text = "[ Status: Idle ][ Duration of last gathering cycle: " + tSpanDuration + " ][ Next gathering cycle: " + sNextCycle + " ]";
                Application.DoEvents();
            }
            else if (bBatchTimer == true)
            {
                DateTime dtime = DateTime.Now;
                string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
                string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
                sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
                sStatTxt = sStatDTNow + " The gather cycle has completed. The processing cycle will now begin." + Environment.NewLine;
                sNextCycle = dtime.AddMinutes(15).ToString("HH:mm:ss");

                DateTime dTimeDoWorkEnd = DateTime.Now;
                string sDTEnd = dTimeDoWorkEnd.ToString("HH:mm:ss");
                TimeSpan tSpanDuration = DateTime.Parse(sDTEnd).Subtract(DateTime.Parse(sDTStart));
                this.StatusText.Text = "[ Status: Idle ][ Duration of last gather cycle: " + tSpanDuration + " ]";
                Application.DoEvents();
            }

            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            bool bFinished = true;
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);

            this.btnDoWork.Enabled = true;

            if (bBatchTimer == true)
            {
                this.tmrBatch.Enabled = true;
                this.tmrBatch.Start();
            }
            else if (bBatchTimer == false)
            {
                this.tmrGather.Enabled = true;
                this.tmrGather.Start();
            }

        }

        private void Process(ref string sDTStart) // Work data processing method. Process job and frame data. 
        {
            this.LoopThroughOrderItems();

            string sStatDTNow = string.Empty;
            string sStatTxt = string.Empty;
            string sNextCycle = string.Empty;

            if (bBatchTimer != true)
            {
                DateTime dtime = DateTime.Now;
                string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
                string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
                sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
                sStatTxt = sStatDTNow + " This processing cycle has completed." + Environment.NewLine + sStatDTNow + " Idle for 15 minutes." + Environment.NewLine + Environment.NewLine;
                sNextCycle = dtime.AddMinutes(15).ToString("HH:mm:ss");

                DateTime dTimeDoWorkEnd = DateTime.Now;
                string sDTEnd = dTimeDoWorkEnd.ToString("HH:mm:ss");
                TimeSpan tSpanDuration = DateTime.Parse(sDTEnd).Subtract(DateTime.Parse(sDTStart));
                this.StatusText.Text = "[ Status: Idle ][ Duration of last processing cycle: " + tSpanDuration + " ][ Next processing cycle: " + sNextCycle + " ]";
                Application.DoEvents();
            }
            else if (bBatchTimer == true)
            {
                DateTime dtime = DateTime.Now;
                string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
                string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
                sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
                sStatTxt = sStatDTNow + " This batch cycle has completed." + Environment.NewLine + sStatDTNow + " Idle for 15 minutes." + Environment.NewLine + Environment.NewLine;
                sNextCycle = dtime.AddMinutes(15).ToString("HH:mm:ss");

                string sDTEnd = DateTime.Now.ToString("HH:mm:ss");
                TimeSpan tSpanDuration = DateTime.Parse(sDTEnd).Subtract(DateTime.Parse(sDTStart));
                this.StatusText.Text = "[ Status: Idle ][ Duration of last batch cycle: " + tSpanDuration + " ][ Next batch cycle: " + sNextCycle + " ]";
                Application.DoEvents();
            }

            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            bool bFinished = true;
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);

            this.btnDoWork.Enabled = true;

            if (bBatchTimer == true)
            {
                this.tmrBatch.Enabled = true;
                this.tmrBatch.Start();
            }
            else if (bBatchTimer == false)
            {
                this.tmrProcess.Enabled = true;
                this.tmrProcess.Start();
            }

        }

        #endregion

        #region Gather 1 [Gather and save jobs with ready work to the Verified table.]

        private void SearchItems() // Search CDS for TimberlineColorado items in packages. 
        {
            DateTime dtime = DateTime.Now;
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Checking for Timberline Colorado items within orders.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);            

            string sCommTextValue = string.Empty;
            string sColumn = string.Empty;
            DataTable dtGenericValue = new DataTable();
            string sValueString = string.Empty;
            double dGatherDays = 0;
            sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'GatherDays'";
            sColumn = "Variables_Variable";
            TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
            dGatherDays = Convert.ToDouble(sValueString);


            DateTime dTimeNow = DateTime.Now;
            DateTime dTNowDateOnly = DateTime.Now.Date;
            DateTime dTimeNowMinusDays = DateTime.Now.AddDays(dGatherDays);
            DateTime dTimeNowMinusDaysDateOnly = dTimeNowMinusDays.Date;

            string sDTNowDateOnly = dTNowDateOnly.ToString("MM/dd/yy");
            string sDTNowMinusDaysDateOnly = dTimeNowMinusDaysDateOnly.ToString("MM/dd/yy");

            DataTable dtCodes = new DataTable();
            string sCommText = "SELECT DISTINCT CDSCode FROM [Codes]";
            TM.SQLQuery(sTimberlineConnString, sCommText, dtCodes);

            try
            {
                DataTable dtItems = new DataTable();

                foreach (DataRow dr in dtCodes.Rows)
                {
                    string sCode = Convert.ToString(dr["CDSCode"]).Trim();

                    sCommText = "SELECT Lookupnum FROM ITEMS WHERE items.d_dueout > CTOD('" + sDTNowMinusDaysDateOnly + "') AND ITEMS.PACKAGETAG IN" +
                    " (SELECT PACKAGETAG FROM LABELS WHERE (LABELS.CODE = '" + sCode + "') AND LABELS.PACKAGETAG <> '    ') ORDER BY items.d_dueout ASC";
                    TM.CDSQuery(sCDSConnString, sCommText, dtItems);
                }

                string sColName = "lookupnum";
                int iRowsCount = dtItems.Rows.Count;
                DataTable dt = new DataTable();
                dt = dtItems.Copy();

                iRowsCount = dt.Rows.Count;

                TM.RemoveDuplicateRows(dt, sColName);

                iRowsCount = dt.Rows.Count;

                dtItems.Clear();
                dtItems = dt.Copy();

                iRowsCount = dtItems.Rows.Count;

                if (dtItems.Rows.Count > 0)
                {
                    this.CheckForALaCarte(ref dtItems);
                }
                else if (dtItems.Rows.Count == 0)
                {
                    bNoGatherData = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void CheckForALaCarte(ref DataTable dtItems) // Search CDS for a la carte TimberlineColorado items. 
        {
            string sCommTextValue = string.Empty;
            string sColumn = string.Empty;
            DataTable dtGenericValue = new DataTable();
            string sValueString = string.Empty;
            double dGatherDays = 0;
            sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'GatherDays'";
            sColumn = "Variables_Variable";
            TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
            dGatherDays = Convert.ToDouble(sValueString);

            DateTime dTimeNow = DateTime.Now;
            DateTime dTNowDateOnly = DateTime.Now.Date;
            DateTime dTimeNowMinusDays = DateTime.Now.AddDays(dGatherDays);
            DateTime dTimeNowMinusDaysDateOnly = dTimeNowMinusDays.Date;

            string sDTNowDateOnly = dTNowDateOnly.ToString("MM/dd/yy");
            string sDTNowMinusDaysDateOnly = dTimeNowMinusDaysDateOnly.ToString("MM/dd/yy");

            DataTable dtCodes = new DataTable();
            string sCommText = "SELECT DISTINCT CDSCode FROM [Codes]";
            TM.SQLQuery(sTimberlineConnString, sCommText, dtCodes);

            try
            {
                DataTable dtALaCarte = new DataTable();
                DataTable dt = new DataTable();

                foreach (DataRow dr in dtCodes.Rows)
                {
                    string sCode = Convert.ToString(dr["CDSCode"]).Trim();

                    if (sCode == "CFB")
                    {
                        string sStop = string.Empty;
                    }
                    
                    sCommText = "SELECT Lookupnum FROM CODES WHERE Codes.Code IN ('" + sCode + "') AND Codes.lookupnum" +
                        " IN (Select lookupnum FROM Items WHERE items.d_dueout > CTOD('" + sDTNowMinusDaysDateOnly + "'))";

                    TM.CDSQuery(sCDSConnString, sCommText, dtALaCarte);

                    dt.Merge(dtALaCarte);
                }

                string sColName = "lookupnum";
                int iRowsCount = dtALaCarte.Rows.Count;                
                

                iRowsCount = dt.Rows.Count;

                TM.RemoveDuplicateRows(dt, sColName);

                iRowsCount = dt.Rows.Count;

                dtALaCarte.Clear();
                dtALaCarte = dt.Copy();

                iRowsCount = dtALaCarte.Rows.Count;

                dtItems.Merge(dtALaCarte);

                if (dtItems.Rows.Count > 0)
                {
                    this.ScanForTriggerPoints(dtItems);
                }
                else if (dtItems.Rows.Count == 0)
                {
                    bNoGatherData = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void ScanForTriggerPoints(DataTable dtItems) // Verify what work is ready to process. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Scanning needed trigger points for gathered orders ready to process.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);

            DataTable dt = new DataTable();
            DataColumn dColumn = new DataColumn();
            DataRow dRow;

            if (bCreated == false)
            {
                dtScanned.Columns.Clear();                
                dColumn.DataType = System.Type.GetType("System.String");
                dColumn.ColumnName = "ProdNum";
                dtScanned.Columns.Add(dColumn);
                bCreated = true;
            }

            try
            {
                foreach (DataRow dr in dtItems.Rows)
                {
                    dt.Clear();

                    string sProdNum = Convert.ToString(dr["lookupnum"]).Trim();

                    string sCommText = "SELECT * FROM STAMPS WHERE Lookupnum = '" + sProdNum +
                    "' AND (Action = " + "'DIGI PRINT'" + " OR Action = '" + "PRNT_TRAV" + "' OR Action = " + "'COLOR QC')";

                    TM.CDSQuery(sCDSConnString, sCommText, dt);

                    if (dt.Rows.Count > 0)
                    {
                        dRow = dtScanned.NewRow();
                        dRow["ProdNum"] = sProdNum;
                        dtScanned.Rows.Add(dRow);
                    }
                }

                int iRowsCount = dtScanned.Rows.Count;

                if (dtScanned.Rows.Count > 0)
                {
                    this.GetAdditionalInfoForDTScanned();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetAdditionalInfoForDTScanned()
        {
            try
            {
                DataTable dTbl = new DataTable();

                foreach (DataRow dr in dtScanned.Rows)
                {
                    dTbl.Clear();

                    string sProdNum = Convert.ToString(dr["ProdNum"]).Trim();

                    string sCommText = "SELECT Lookupnum, Order, Packagetag FROM ITEMS WHERE Lookupnum = '" + sProdNum + "'";

                    TM.CDSQuery(sCDSConnString, sCommText, dTbl);

                    if (dTbl.Rows.Count > 0)
                    {
                        dtScannedAddInfo.Merge(dTbl);
                    }
                    else if (dTbl.Rows.Count == 0)
                    {
                        return;
                    }

                    int iRowsCount = dtScannedAddInfo.Rows.Count;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void ExistsInDP2() // Verify the work exists in DP2. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Verifying that the order exists in DP2.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);

            DataTable dt = new DataTable();

            try
            {
                foreach (DataRow dr in dtScannedAddInfo.Rows)
                {
                    dt.Clear();

                    string sRefNum = Convert.ToString(dr["Order"]).Trim();
                    string sProdNum = Convert.ToString(dr["lookupnum"]).Trim();
                    string sPkgTag = Convert.ToString(dr["Packagetag"]).Trim();

                    if (sRefNum == "9236102")
                    {
                        string sStop = string.Empty;
                    }

                    string sCommText = "SELECT [ID] FROM [Orders] WHERE [ID] = '" + sRefNum + "'";

                    TM.SQLQuery(sDP2ConnString, sCommText, dt);

                    if (dt.Rows.Count > 0)
                    {
                        this.CheckForFrames(sRefNum);
                        this.CheckCodes(sProdNum, sPkgTag);
                    }
                    else if (dt.Rows.Count == 0)
                    {
                        //return;
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void CheckForFrames(string sRefNum)
        {
            if (sRefNum == "9236102")
            {
                string sStop = string.Empty;
            }

            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [Images] WHERE [OrderID] = '" + sRefNum + "'";

                TM.SQLQuery(sDP2ConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    // Do nothing. Just checking for records in the Images table of DP2.
                }
                else if(dTbl.Rows.Count == 0)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void CheckCodes(string sProdNum, string sPkgTag) // Gather codes containing TimberlineColorado items. 
        {
            if (sProdNum == "15371")
            {
                string sStop = string.Empty;
            }

            DataTable dtCodes = new DataTable();
            string sCommText = "SELECT DISTINCT CDSCode FROM [Codes]";
            TM.SQLQuery(sTimberlineConnString, sCommText, dtCodes);

            try
            {
                DataTable dTbl = new DataTable();

                foreach (DataRow dr in dtCodes.Rows)
                {
                    dTbl.Clear();

                    string sCode = Convert.ToString(dr["CDSCode"]).Trim();

                    sCommText = "SELECT * FROM Codes WHERE Lookupnum = '" + sProdNum + "' AND ((Codes.Code = '" + sCode + "' AND Package = .F.)" +
                        " OR (Package = .T. AND Code IN (SELECT DISTINCT Packagecod FROM Labels WHERE Labels.packagetag = '"
                        + sPkgTag + "' AND (Labels.Code = '" + sCode + "'))))" +
                    " ORDER BY Sequence ASC";

                    TM.CDSQuery(sCDSConnString, sCommText, dTbl);

                    if (dTbl.Rows.Count > 0)
                    {
                        dtResults.Merge(dTbl);
                    }                    
                }

                int iRowsCount = dtResults.Rows.Count;
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void SaveResultsToVerifiedTable() // Save gathered results to database. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Saving gathered orders ready to process to database.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);            

            int iCount = dtResults.Rows.Count;

            string sCommTxt = "INSERT INTO [Timberline].[dbo].[Verified] ([ProdNum], [RefNum], [FrameNum], [InDP2], [InDP2Check], [CustNum], [PkgTag], [ServiceType])" +
                " VALUES (@A, @B, @C, 'T', @D, @E, @F, @G)";

            if (dtResults.Rows.Count > 0)
            {
                try
                {
                    using (SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString))
                    {
                        myConn.Open();

                        foreach (DataRow dr in dtResults.Rows)
                        {
                            try
                            {
                                string sProdNum = Convert.ToString(dr["lookupnum"]).Trim();
                                string sFrameNum = Convert.ToString(dr["sequence"]).Trim(); 
                                sFrameNum = sFrameNum.PadLeft(4, '0');

                                bool bGotRefNum = true;
                                DataTable dTbl = new DataTable();
                                this.GetRefNumForVerified(sProdNum, ref dTbl, ref bGotRefNum);

                                if (bGotRefNum == true)
                                {
                                    string sRefNum = Convert.ToString(dTbl.Rows[0]["order"]).Trim();
                                    string sCustNum = Convert.ToString(dTbl.Rows[0]["customer"]).Trim();
                                    string sPkgTag = Convert.ToString(dTbl.Rows[0]["packagetag"]).Trim();
                                    string sServiceType = Convert.ToString(dTbl.Rows[0]["Sertype"]).Trim();

                                    SqlCommand myCommand = myConn.CreateCommand();
                                    myCommand.CommandText = sCommTxt;
                                    myCommand.Parameters.AddWithValue("@A", sProdNum);
                                    myCommand.Parameters.AddWithValue("@B", sRefNum);
                                    myCommand.Parameters.AddWithValue("@C", sFrameNum);
                                    string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                                    myCommand.Parameters.AddWithValue("@D", sDTNow);
                                    myCommand.Parameters.AddWithValue("@E", sCustNum);
                                    myCommand.Parameters.AddWithValue("@F", sPkgTag);
                                    myCommand.Parameters.AddWithValue("@G", sServiceType);
                                    myCommand.ExecuteNonQuery();
                                }
                                else if (bGotRefNum == false)
                                {
                                    string sMySubject = "Timberline Error: No data in the Items table.";
                                    string sMyBody = "No data in the Items table for Production #: " + sProdNum + " and Frame #: " + sFrameNum + "'";

                                    EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);

                                }
                            }
                            catch (System.Data.SqlClient.SqlException)
                            {
                                // Do nothing but fall back into the above foreach loop.
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    TM.SaveExceptionToDB(ex);
                }
            }
        }        

        private void GetRefNumForVerified(string sProdNum, ref DataTable dTbl, ref bool bGotRefNum)
        {
            try
            {
                string sCommText = "SELECT Order, Customer, Packagetag, Sertype FROM ITEMS WHERE Lookupnum = '" + sProdNum + "'";

                TM.CDSQuery(sCDSConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    bGotRefNum = true;
                }
                else if (dTbl.Rows.Count == 0)
                {
                    bGotRefNum = false;
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #region Gather 2 [Loop through records saved to the Verified table in Gather 1, gather and save frame and order item data to the OrderItems table.]

        private void LoopThroughVerifiedTable() // Foreach through the Verified table and gather frame data. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Processing verified orders.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);            

            // This will grab the first record (reference number) in the Verified table where Processed = NULL and gather order item and frame data for the entire order to save to the OrderItems table.
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [Verified] WHERE ([Processed] IS NULL OR [Processed] = 'F')";

                TM.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    foreach (DataRow dRow in dTbl.Rows)
                    {
                        string sRefNum = Convert.ToString(dRow["RefNum"]).Trim();
                        string sProdNum = Convert.ToString(dRow["ProdNum"]).Trim();
                        string sPkgTag = Convert.ToString(dRow["PkgTag"]).Trim();
                        string sFrameNum = Convert.ToString(dRow["FrameNum"]).Trim();

                        string sSequence = sFrameNum.TrimStart('0');

                        this.GetOrderItemsToRender(sRefNum, sProdNum, sPkgTag, sSequence);

                        if (dtFinal.Rows.Count > 0)
                        {
                            int iFinalRowCount = dtFinal.Rows.Count;

                            DataTable dtGeneric = new DataTable();
                            sCommText = "SELECT * FROM [Verified] WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NULL";
                            TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);
                            int iIsNullRowCount = dtGeneric.Rows.Count;

                            DataView dv = dtFinal.DefaultView;
                            dv.Sort = "ItemID DESC";
                            dtFinal = dv.ToTable();

                            TM.RemoveDuplicatesFromDataTable(ref dtFinal);

                            this.SaveToOrderItems(ref dtFinal);

                            // If notnullrowcount < isnullrowcount flag records that processed = is null as processed.
                            // Example:
                            //      1 frame out of 5 has an order item in dp2 that is listed in Codes
                            //      the other 4 do not contain any order items in dp2 that are listed in the Codes table
                            //      this prevents the picking up of the ref num in the next cycle which becomes a loop at that point.

                            dtGeneric.Clear();
                            sCommText = "SELECT * FROM [Verified] WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NOT NULL";
                            TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);
                            int iIsNotNullRowCount = dtGeneric.Rows.Count;

                            if (iIsNotNullRowCount < iIsNullRowCount)
                            {
                                string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                                sCommText = "UPDATE [Verified] SET [Processed] = 'T', [ProcessedTime] = '" + sDTNow + "' WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NULL";
                                TM.SQLNonQuery(sTimberlineConnString, sCommText);
                            }
                        }
                        else if (dtFinal.Rows.Count == 0)
                        {
                            // If there are no order items within an order that match the codes from the Codes table then flag entire job as processed.

                            string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                            sCommText = "UPDATE [Verified] SET [Processed] = 'T', [ProcessedTime] = '" + sDTNow + "' WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NULL";
                            TM.SQLNonQuery(sTimberlineConnString, sCommText);

                            // Note: Send an email here indicating no matching Codes in DP2.
                        }
                    }
                }

                #region old code

                //if (dTbl.Rows.Count > 0)
                //{
                //    string sRefNum = Convert.ToString(dTbl.Rows[0]["RefNum"]).Trim();
                //    string sProdNum = Convert.ToString(dTbl.Rows[0]["ProdNum"]).Trim();
                //    string sPkgTag = Convert.ToString(dTbl.Rows[0]["PkgTag"]).Trim();
                //    string sFrameNum = Convert.ToString(dTbl.Rows[0]["FrameNum"]).Trim();

                //    string sSequence = sFrameNum.TrimStart('0');

                //    this.GetOrderItemsToRender(sRefNum, sProdNum, sPkgTag, sSequence);

                //    if (dtFinal.Rows.Count > 0)
                //    {
                //        int iFinalRowCount = dtFinal.Rows.Count;

                //        DataTable dtGeneric = new DataTable();
                //        sCommText = "SELECT * FROM [Verified] WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NULL";
                //        TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);
                //        int iIsNullRowCount = dtGeneric.Rows.Count;

                //        DataView dv = dtFinal.DefaultView;
                //        dv.Sort = "ItemID DESC";
                //        dtFinal = dv.ToTable();

                //        TM.RemoveDuplicatesFromDataTable(ref dtFinal);

                //        this.SaveToOrderItems(ref dtFinal);

                //        // If notnullrowcount < isnullrowcount flag records that processed = is null as processed.
                //        // Example:
                //        //      1 frame out of 5 has an order item in dp2 that is listed in Codes
                //        //      the other 4 do not contain any order items in dp2 that are listed in the Codes table
                //        //      this prevents the picking up of the ref num in the next cycle which becomes a loop at that point.

                //        dtGeneric.Clear();
                //        sCommText = "SELECT * FROM [Verified] WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NOT NULL";
                //        TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);
                //        int iIsNotNullRowCount = dtGeneric.Rows.Count;

                //        if (iIsNotNullRowCount < iIsNullRowCount)
                //        {
                //            string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                //            sCommText = "UPDATE [Verified] SET [Processed] = 'T', [ProcessedTime] = '" + sDTNow + "' WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NULL";
                //            TM.SQLNonQuery(sTimberlineConnString, sCommText);
                //        }
                //    }
                //    else if (dtFinal.Rows.Count == 0)
                //    {
                //        // If there are no order items within an order that match the codes from the Codes table then flag entire job as processed.

                //        string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                //        sCommText = "UPDATE [Verified] SET [Processed] = 'T', [ProcessedTime] = '" + sDTNow + "' WHERE [RefNum] = '" + sRefNum + "' AND [Processed] IS NULL";
                //        TM.SQLNonQuery(sTimberlineConnString, sCommText);

                //        // Note: Send an email here indicating no matching Codes in DP2.

                //        return;
                //    }
                //}

                #endregion 

                else if (dTbl.Rows.Count == 0)
                {
                    //bNoGatherData = true; //Note: not sure if needed.
                    //return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetOrderItemsToRender(string sRefNum, string sProdNum, string sPkgTag, string sSequence) // Gather OrderItems to render. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Gathering order item data.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);


            DataTable dtToRender1 = new DataTable();

            DataTable dtCodes = new DataTable();
            string sCommText = "SELECT * FROM [Codes]";
            TM.SQLQuery(sTimberlineConnString, sCommText, dtCodes);

            try
            {
                DataTable dTbl = new DataTable();
                DataTable dTbl02 = new DataTable();

                foreach (DataRow drCodes in dtCodes.Rows)
                {
                    dTbl.Clear();

                    bGotCode = false;

                    string sCode = Convert.ToString(drCodes["OrderItemCode"]).Trim();
                    
                    //sCommText = "SELECT [OrderID], [ID], [Sequence], [Quantity], [ProductID] FROM [OrderItems] WHERE [OrderID] = '" + sRefNum + "' AND [ProductID] Like '" + sCode + "%'";
                    sCommText = "SELECT [OrderID], [ID], [Sequence], [Quantity], [ProductID] FROM [OrderItems] WHERE [OrderID] = '" + sRefNum + "' AND [ProductID] = '" + sCode + "'";

                    TM.SQLQuery(sDP2ConnString, sCommText, dTbl);

                    if (dTbl.Rows.Count > 0)
                    {
                        dtToRender1.Clear();
                        dtToRender1 = dTbl.Copy();
                        dTbl.Clear();

                        this.GetOrderItemsToRenderAddInfo(sRefNum, dtToRender1, sCode);

                        bGotCode = true;
                    }
                    if (dTbl.Rows.Count == 0 && bGotCode != true)
                    {
                        DataTable dTbl01 = new DataTable();
                        sCommText = "SELECT Code FROM Codes WHERE Lookupnum = '" + sProdNum + "' AND Sequence = " + sSequence;

                        TM.CDSQuery(sCDSConnString, sCommText, dTbl01);

                        foreach (DataRow dRow in dTbl01.Rows)
                        {
                            dTbl02.Clear();

                            string sCodesCode = Convert.ToString(dRow["Code"]).Trim();
                            
                            sCommText = "SELECT Code FROM Labels WHERE Packagetag = '" + sPkgTag + "' AND Packagecod = '" + sCodesCode + "' AND Code = '" + sCode + "'";

                            TM.CDSQuery(sCDSConnString, sCommText, dTbl02);

                            if (dTbl02.Rows.Count > 0)
                            {
                                DataTable dTbl03 = new DataTable();
                                sCommText = "SELECT [OrderID], [ID], [Sequence], [Quantity], [ProductID] FROM [OrderItems] WHERE [OrderID] = '" + sRefNum + "' AND [ProductID] = '" + sCode + "'";

                                TM.SQLQuery(sDP2ConnString, sCommText, dTbl03);

                                if (dTbl03.Rows.Count > 0)
                                {
                                    dtToRender1.Clear();
                                    dtToRender1 = dTbl.Copy();
                                    dTbl.Clear();

                                    this.GetOrderItemsToRenderAddInfo(sRefNum, dtToRender1, sCode);

                                    bGotCode = true;
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetOrderItemsToRenderAddInfo(string sRefNum, DataTable dtToRender1, string sCode)
        {
            DataTable dTbl01 = new DataTable();
            bool bColumnCreated = false;
            string sID = string.Empty;

            DataTable dTbl02 = new DataTable();

            try
            {
                foreach (DataRow dr in dtToRender1.Rows)
                {
                    dTbl02.Clear();

                    sID = Convert.ToString(dr["ID"]).Trim();                    

                    string sCommText = "SELECT [OrderID], [Roll], [Frame], [ItemID] FROM [OrderItemImages] WHERE [OrderID] = '" + sRefNum + "' AND [ItemID] IN (" + sID + ")";

                    TM.SQLQuery(sDP2ConnString, sCommText, dTbl02);

                    if (dTbl02.Rows.Count > 0)
                    {
                        dTbl01.Clear();
                        dTbl01 = dTbl02.Copy();
                        dTbl02.Clear();

                        if (bColumnCreated == false)
                        {
                            DataColumn dc = new DataColumn("Code", typeof(String));
                            dc.DefaultValue = sCode;
                            dTbl01.Columns.Add(dc);
                            bColumnCreated = true;
                        }

                        dtFinal.Merge(dTbl01);
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void SaveToOrderItems(ref DataTable dtFinal) // Save gathered frame data to database. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Saving frame and order item data.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);


            int iQuantity = 0;
            string sProdID = string.Empty;
            string sPartNum = string.Empty;
            string sImgloc = string.Empty;

            string sCommTxt = "INSERT INTO [Timberline].[dbo].[OrderItems] ([ProdNum], [RefNum], [FrameNum], [ItemID], [ImagePath], [UID], [Quantity], [OrderItemCode], [PartNumber])" +
                " VALUES (@A, @B, @C, @D, @E, @F, @G, @H, @I)";

            if (dtFinal.Rows.Count > 0)
            {
                try
                {
                    using (SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString))
                    {
                        myConn.Open();

                        foreach (DataRow dr in dtFinal.Rows)
                        {
                            try
                            {
                                string sRefNum = Convert.ToString(dr["OrderID"]).Trim();
                                string sProdNum = Convert.ToString(dr["Roll"]).Trim();
                                string sFrameNum = Convert.ToString(dr["Frame"]).Trim();
                                sFrameNum = sFrameNum.PadLeft(4, '0');
                                string sItemID = Convert.ToString(dr["ItemID"]).Trim();
                                string sCode = Convert.ToString(dr["Code"]).Trim();
                                string sUID = sRefNum + sFrameNum + sItemID;
                                
                                this.GetImageLocation(sRefNum, sFrameNum, ref sImgloc);

                                this.GetProductID(sRefNum, sItemID, ref sProdID, ref iQuantity);
                                sProdID = sProdID.Substring(0, 3);
                                this.GetPartNumber(sProdID, ref sPartNum);

                                this.UpdateVerifiedProcessed(ref sRefNum, ref sFrameNum);

                                SqlCommand myCommand = myConn.CreateCommand();
                                myCommand.CommandText = sCommTxt;
                                myCommand.Parameters.AddWithValue("@A", sProdNum);
                                myCommand.Parameters.AddWithValue("@B", sRefNum);
                                myCommand.Parameters.AddWithValue("@C", sFrameNum);
                                myCommand.Parameters.AddWithValue("@D", sItemID);
                                myCommand.Parameters.AddWithValue("@E", sImgloc);
                                myCommand.Parameters.AddWithValue("@F", sUID);
                                myCommand.Parameters.AddWithValue("@G", iQuantity);
                                myCommand.Parameters.AddWithValue("@H", sCode);
                                myCommand.Parameters.AddWithValue("@I", sPartNum);
                                myCommand.ExecuteNonQuery();
                            }
                            catch (System.Data.SqlClient.SqlException)
                            {
                                // Do nothing but fall back into the above foreach loop.
                            }
                        }

                        dtFinal.Clear();
                    }
                }
                catch (Exception ex)
                {
                    TM.SaveExceptionToDB(ex);
                }
            }
        }

        private void GetProductID(string sRefNum, string sItemID, ref string sProdID, ref int iQuantity) // Gather the product id for each order item. 
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT [ProductID], [Quantity] FROM [OrderItems] WHERE [OrderID] = '" + sRefNum + "' AND [Sequence] = '" + sItemID + "'";

                TM.SQLQuery(sDP2ConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    sProdID = Convert.ToString(dTbl.Rows[0]["ProductID"]).Trim();
                    iQuantity = Convert.ToInt32(dTbl.Rows[0]["Quantity"]);
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetPartNumber(string sProdID, ref string sPartNum) // Gather the TimberlineColorado part number. 
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT [PartNumber] FROM [Codes] WHERE [CDSCode] = '" + sProdID + "'";

                TM.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    sPartNum = Convert.ToString(dTbl.Rows[0]["PartNumber"]).Trim();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GetImageLocation(string sRefNum, string sFrameNum, ref string sImgLoc)
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT [Path] FROM [Images] WHERE [OrderID] = '" + sRefNum + "' AND [Frame] = '" + sFrameNum + "'";

                TM.SQLQuery(sDP2ConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    sImgLoc = Convert.ToString(dTbl.Rows[0]["Path"]).Trim();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void UpdateVerifiedProcessed(ref string sRefNum, ref string sFrame) // Update the verified table that the associated record has had frame data saved for it. 
        {
            try
            {
                string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                string sCommText = "UPDATE [Verified] SET [Processed] = 'T', [ProcessedTime] = '" + sDTNow + "' WHERE [RefNum] = '" + sRefNum + "' AND [FrameNum] = '" + sFrame + "'";

                TM.SQLNonQuery(sTimberlineConnString, sCommText);
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

        #region Process [Loop through the OrderItem table, generate AutoGen and XML files.]

        private void LoopThroughOrderItems() // Process frame data records. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Gathering data for AutoGen and XML file creation.";
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);

            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [OrderItems] WHERE ([AutoGenFileCreated] IS NULL OR [AutoGenFileCreated] = 'F') OR (([Rendered] IS NULL OR [Rendered] = 'F')" +
                    " OR ([XmlFileCreated] IS NULL OR [XmlFileCreated] = 'F') AND [XmlErrorFilePath] IS NULL)";

                TM.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    foreach (DataRow dRow in dTbl.Rows)
                    {

                        string sRefNum = Convert.ToString(dRow["RefNum"]).Trim();
                        string sPRodNum = Convert.ToString(dRow["ProdNum"]).Trim();
                        string sFrameNum = Convert.ToString(dRow["FrameNum"]).Trim();
                        string sOrderItemID = Convert.ToString(dRow["ItemID"]).Trim();
                        string sQuantity = Convert.ToString(dRow["Quantity"]).Trim();
                        string sCode = Convert.ToString(dRow["OrderItemCode"]).Trim();
                        string sPartNum = Convert.ToString(dRow["PartNumber"]).Trim();
                        string sImagePath = Convert.ToString(dRow["ImagePath"]).Trim();
                        string sUID = Convert.ToString(dRow["UID"]);

                        this.GenerateAutoGenFile(sRefNum, sOrderItemID, sUID);

                        if (bAutoGenFileError == false)
                        {
                            this.CheckForRenderedImages(sRefNum, sPRodNum, sFrameNum, sCode, sImagePath, sUID, sOrderItemID);
                        }
                        else if (bAutoGenFileError == true)
                        {
                            return; // Fall back into loop.
                        }
                        if (bRenderError == false)
                        {
                            this.GenerateXML(sQuantity, sPRodNum, sRefNum, sFrameNum, sCode, sPartNum, sOrderItemID, sUID);
                        }
                        else if (bRenderError == true)
                        {
                            return; // Fall back into loop.
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GenerateAutoGenFile(string sRefNum, string sOrderItemID, string sUID) // Create the autogen file required for rendering. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Creating AutoGen file for reference number: " + sRefNum;
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);


            string sCommTextValue = string.Empty;
            DataTable dtGenericValue = new DataTable();
            string sValueString = string.Empty;
            string sCTExportTemplate = string.Empty;
            string sColumn = "Variables_Variable";

            sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'AutoGenDirPath'";
            TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
            string sAutoGenPath = Convert.ToString(sValueString);

            sAutoGenPath += @"CTExport-" + sRefNum + "-" + sOrderItemID + ".txt";


            sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'TCStorageDir'";
            dtGenericValue.Clear();
            TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
            string sStoredFile = Convert.ToString(sValueString);

            sStoredFile += sUID + @"\CTExport-" + sRefNum + "-" + sOrderItemID + ".txt";
            string sStoredPath = Path.GetDirectoryName(sStoredFile);

            if (!Directory.Exists(sStoredPath))
            {
                Directory.CreateDirectory(sStoredPath);
            }

            try
            {
                sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'CTExportTemplate'";
                dtGenericValue.Clear();
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                sCTExportTemplate = Convert.ToString(sValueString);

                string sText = File.ReadAllText(sCTExportTemplate);
                sText = sText.Replace("RefNum", sRefNum);
                sText = sText.Replace("OrdItem", sOrderItemID);
                File.WriteAllText(sAutoGenPath, sText);

                if (!File.Exists(sStoredFile))
                {
                    File.WriteAllText(sStoredFile, sText);
                }                

                if (File.Exists(sStoredFile))
                {
                    string sDTNow = DateTime.Now.ToString();
                    string sCommText = "UPDATE [OrderItems] SET [AutoGenFileCreated] = 'T', [AutoGenFileCreatedTime] = '" + sDTNow + "' WHERE [RefNum] = '" +
                        sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                }
                else if (!File.Exists(sStoredFile))
                {
                    // Check autogen job files\errors if files.exists if so flag autogenfilecreated 'E'.
                    sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'AutoGenErrorDir'";
                    dtGenericValue.Clear();
                    TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                    string sErrorDir = Convert.ToString(sValueString);

                    string sErrorFile = sErrorDir + @"\CTExport-" + sRefNum + "-" + sOrderItemID + ".txt";

                    if (File.Exists(sErrorFile))
                    {
                        string sCommText = "UPDATE [OrderItems] SET [AutoGenFileCreated] = 'E - Error Parsing' WHERE [RefNum] = '" + sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                        TM.SQLNonQuery(sTimberlineConnString, sCommText);
                        bAutoGenFileError = true;

                        string sMySubject = "Timberline Error: Error parsing the AutoGen file.";
                        string sMyBody = "Error parsing AutoGen file for Reference #: " + sRefNum + " and Order Item #: " + sOrderItemID + "'";

                        EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);
                    }
                    else if (!File.Exists(sErrorFile))
                    {
                        string sCommText = "UPDATE [OrderItems] SET [AutoGenFileCreated] = 'E - Missing AutoGen File' WHERE [RefNum] = '" + sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                        TM.SQLNonQuery(sTimberlineConnString, sCommText);
                        bAutoGenFileError = true;

                        string sMySubject = "Timberline Error: Missing AutoGen file.";
                        string sMyBody = "Missing AutoGen file for Reference #: " + sRefNum + " and Order Item #: " + sOrderItemID + "'";

                        EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);
                    }
                }

            }
            catch (Exception ex)
            {
                bAutoGenFileError = true;
                TM.SaveExceptionToDB(ex);
            }
        }

        private void CheckForRenderedImages(string sRefNum, string sProdNum, string sFrameNum, string sCode, string sImagePath, string sUID, string sOrderItemID) // Verify images have been rendered. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Verifying image has been rendered for reference number: " + sRefNum;
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);
            
            string sDTNow = string.Empty;
            string sCommText = string.Empty;

            sCode = sCode.Substring(0, 3);

            try
            {
                string sCommTextValue = string.Empty;
                DataTable dtGenericValue = new DataTable();
                string sValueString = string.Empty;
                string sCTExportTemplate = string.Empty;
                string sColumn = "Variables_Variable";

                sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'TCIntWebServerDir'";
                dtGenericValue.Clear();
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                string sWebServerDir = Convert.ToString(sValueString);

                sImagePath = Path.GetDirectoryName(sImagePath);
                string sImageName = sRefNum + "_" + sProdNum + "_" + sFrameNum + ".jpg";
                sWebServerDir += sUID + @"\" + sCode + @"\";
                
                if (!Directory.Exists(sWebServerDir))
                {
                    Directory.CreateDirectory(sWebServerDir);
                }

                sWebServerDir += sImageName;

                // Structure on all rendered images ProdNum\APSIMG\CTEXPORT\*CODE*\*RefNum*_*PNum*_*FrameNum*.jpg.

                sImagePath = sImagePath + @"\APSIMG\CTEXPORT\" + sCode + @"\" + sImageName;
                if (File.Exists(sImagePath) && !File.Exists(sWebServerDir)) // If rendered images exist and are not on the webserver then copy them to it.
                {
                    File.Copy(sImagePath, sWebServerDir);
                    sDTNow = DateTime.Now.ToString();
                    sCommText = "UPDATE [OrderItems] SET [Rendered] = 'T', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                        sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                }
                else if (!File.Exists(sImagePath)) // If rendered images are not located sleep the thread 5 minutes.
                {
                    sDTNow = DateTime.Now.ToString();
                    sCommText = "UPDATE [OrderItems] SET [Rendered] = 'F', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                        sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                    TM.SQLNonQuery(sTimberlineConnString, sCommText);

                    TimeSpan tsInterval = new TimeSpan(0, 5, 0);
                    Thread.Sleep(tsInterval);

                    if (File.Exists(sImagePath) && !File.Exists(sWebServerDir))
                    {
                        File.Copy(sImagePath, sWebServerDir);
                        sDTNow = DateTime.Now.ToString();
                        sCommText = "UPDATE [OrderItems] SET [Rendered] = 'T', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                            sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                        TM.SQLNonQuery(sTimberlineConnString, sCommText);
                    }
                    else if (!File.Exists(sImagePath)) // If rendered images are not located sleep the thread 10 minutes.
                    {
                        sDTNow = DateTime.Now.ToString();
                        sCommText = "UPDATE [OrderItems] SET [Rendered] = 'F', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                            sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                        TM.SQLNonQuery(sTimberlineConnString, sCommText);

                        TimeSpan tsInterval2 = new TimeSpan(0, 10, 0);
                        Thread.Sleep(tsInterval2);

                        if (File.Exists(sImagePath) && !File.Exists(sWebServerDir))
                        {
                            File.Copy(sImagePath, sWebServerDir);
                            sDTNow = DateTime.Now.ToString();
                            sCommText = "UPDATE [OrderItems] SET [Rendered] = 'T', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                                sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                            TM.SQLNonQuery(sTimberlineConnString, sCommText);
                        }
                        else if (!File.Exists(sImagePath)) // If rendered images are not located sleep the thread 15 minutes.
                        {
                            sDTNow = DateTime.Now.ToString();
                            sCommText = "UPDATE [OrderItems] SET [Rendered] = 'F', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                                sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                            TM.SQLNonQuery(sTimberlineConnString, sCommText);

                            TimeSpan tsInterval3 = new TimeSpan(0, 15, 0);
                            Thread.Sleep(tsInterval3);

                            if (File.Exists(sImagePath) && !File.Exists(sWebServerDir))
                            {
                                File.Copy(sImagePath, sWebServerDir);
                                sDTNow = DateTime.Now.ToString();
                                sCommText = "UPDATE [OrderItems] SET [Rendered] = 'T', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                                    sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                                TM.SQLNonQuery(sTimberlineConnString, sCommText);
                            }
                            else if (!File.Exists(sImagePath)) // If rendered images are not located flag as an error and move on.
                            {
                                sDTNow = DateTime.Now.ToString();
                                sCommText = "UPDATE [OrderItems] SET [Rendered] = 'E - 30minNoRender', [LastRenderedCheck] = '" + sDTNow + "' WHERE [RefNum] = '" +
                                    sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                                TM.SQLNonQuery(sTimberlineConnString, sCommText);
                                bRenderError = true;

                                string sMySubject = "Timberline Error: Images not rendered.";
                                string sMyBody = "Images did not rendered after 30 minutes for Reference #: " + sRefNum + " and Order Item #: " + sOrderItemID + "'";

                                EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GenerateXML(string sQuantity, string sProdNum, string sRefNum, string sFrameNum, string sCode, string sPartNum, string sOrderItemID, string sUID) // Generate the XML file. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " Generating XML file for reference number: " + sRefNum;
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);


            DataTable dtShippingData = new DataTable();
            StringBuilder sb = new StringBuilder();
            DataTable dtGenericValue = new DataTable();
            string sCommText = string.Empty;

            string sName = string.Empty;            
            string sAddy1 = string.Empty;
            string sAddy2 = string.Empty;
            string sCity = string.Empty;
            string sState = string.Empty;
            string sZip = string.Empty;
            string sPhone = string.Empty;
            string sShipMethod = string.Empty;

            sCode = sCode.Substring(0, 3);

            try
            {
                string sCommTextValue = string.Empty;
                dtGenericValue.Clear();
                string sValueString = string.Empty;
                string sXML = string.Empty;                
                string sColumn = "Variables_Variable";

                sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'TCStorageDir'";
                dtGenericValue.Clear();
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                string sXMLFinal = Convert.ToString(sValueString);

                sXMLFinal += sUID + @"\" + sUID + ".xml";

                this.GatherXMLData(ref dtShippingData, sProdNum, sFrameNum);

                if (dtShippingData.Rows.Count > 0)
                {

                    sName = Convert.ToString(dtShippingData.Rows[0]["attn"]).Trim();
                    sName = sName.Replace("&", "and");
                    sAddy1 = Convert.ToString(dtShippingData.Rows[0]["address1"]).Trim();
                    sAddy2 = Convert.ToString(dtShippingData.Rows[0]["address2"]).Trim();
                    sCity = Convert.ToString(dtShippingData.Rows[0]["city"]).Trim();
                    sState = Convert.ToString(dtShippingData.Rows[0]["state"]).Trim();
                    sZip = Convert.ToString(dtShippingData.Rows[0]["zip"]).Trim();
                    sPhone = Convert.ToString(dtShippingData.Rows[0]["phone"]).Trim();
                }
                else if (dtShippingData.Rows.Count == 0)
                {
                    // Note:
                    // Flag the OrderItems table record for the order item to XmlFileCreated = 'F'                    
                    // If there is no data in the Directship table then query "ShipTo"

                    DataTable dTbl = new DataTable();
                    sCommText = "SELECT * FROM [ShipTo] WHERE [UID] = '" + sUID + "'";

                    TM.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                    if (dTbl.Rows.Count > 0)
                    {
                        sName = Convert.ToString(dTbl.Rows[0]["Recipient"]).Trim();
                        sAddy1 = Convert.ToString(dTbl.Rows[0]["Address1"]).Trim();
                        sAddy2 = Convert.ToString(dTbl.Rows[0]["Address2"]).Trim();
                        sCity = Convert.ToString(dTbl.Rows[0]["City"]).Trim();
                        sState = Convert.ToString(dTbl.Rows[0]["State"]).Trim();
                        sZip = Convert.ToString(dTbl.Rows[0]["Zip"]).Trim();
                        sPhone = Convert.ToString(dTbl.Rows[0]["Phone"]).Trim();
                        sShipMethod = Convert.ToString(dTbl.Rows[0]["ShipMethod"]).Trim();
                    }
                    else if (dTbl.Rows.Count == 0)
                    {
                        // Send an email to alert that no ship to information is available.
                        string sMysubject = "Timberline Error - No shipping info.";
                        string sMybody = "There is no shipping data in the Directship table for Reference Number: " + sRefNum + " and Order Item: " + sOrderItemID + "." +
                            Environment.NewLine + "UID: " + sUID + "'";

                        EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMysubject, sMybody);

                        string sCommTxt = "UPDATE [OrderItems] SET [XmlFileCreated] = 'E' WHERE [UID] = '" + sUID + "'";
                        TM.SQLNonQuery(sTimberlineConnString, sCommText);
                        bXMLFileError = true;
                        return;
                    }
                }

                sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'TCXMLTemplate'";
                dtGenericValue.Clear();
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                sXML = Convert.ToString(sValueString);

                string sXMLFile = File.ReadAllText(sXML);

                sCommTextValue = string.Empty;
                sColumn = string.Empty;
                dtGenericValue.Clear();
                sValueString = string.Empty;
                string sAffiliateID = string.Empty;
                sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'APSAffiliateID'";
                sColumn = "Variables_Variable";
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                sAffiliateID = Convert.ToString(sValueString);

                sXMLFile = sXMLFile.Replace("APSOrderID", sUID);

                sXMLFile = sXMLFile.Replace("APSAffiliateID", sAffiliateID);
                sXMLFile = sXMLFile.Replace("APSShipToName", sName);
                sXMLFile = sXMLFile.Replace("APSShipToContact", sName);
                sXMLFile = sXMLFile.Replace("APSAddress1", sAddy1);
                sXMLFile = sXMLFile.Replace("APSAddress2", sAddy2);
                sXMLFile = sXMLFile.Replace("APSCity", sCity);
                sXMLFile = sXMLFile.Replace("APSState", sState);
                sXMLFile = sXMLFile.Replace("APSZip", sZip);
                sXMLFile = sXMLFile.Replace("APSEmail", "");
                sXMLFile = sXMLFile.Replace("APSPhone", sPhone);
                sXMLFile = sXMLFile.Replace("APSShipMethod", sShipMethod);
                string sDTNow = DateTime.Now.ToString();
                sXMLFile = sXMLFile.Replace("APSOrderDate", sDTNow);
                sXMLFile = sXMLFile.Replace("APSSoldToName", sName);
                sXMLFile = sXMLFile.Replace("APSSoldToAddress1", sAddy1);
                sXMLFile = sXMLFile.Replace("APSSoldToAddress2", sAddy2);
                sXMLFile = sXMLFile.Replace("APSSoldToCity", sCity);
                sXMLFile = sXMLFile.Replace("APSSoldToCountry", "US");
                sXMLFile = sXMLFile.Replace("APSSoldToState", sState);
                sXMLFile = sXMLFile.Replace("APSSoldToProvince", "");
                sXMLFile = sXMLFile.Replace("APSSoldToZip", sZip);

                File.WriteAllText(sXMLFinal, sXMLFile);

                string sItemCount = string.Empty;
                sItemCount = "01";

                sCommTextValue = string.Empty;
                sColumn = string.Empty;
                dtGenericValue.Clear();
                sValueString = string.Empty;
                string sWebServerDir = string.Empty;
                sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'TCWebServerDir'";
                sColumn = "Variables_Variable";
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                sWebServerDir = Convert.ToString(sValueString);

                sWebServerDir = sWebServerDir + sUID + @"/";
                string sImagePath = string.Empty;
                string sImageName = sRefNum + "_" + sProdNum + "_" + sFrameNum + ".jpg";

                sImagePath = sWebServerDir + sCode + @"/" + sImageName;

                sCommText = "UPDATE [OrderItems] SET [RenderedWebPath] = '" + sImagePath + "' WHERE [UID] = '" + sUID + "'";
                TM.SQLNonQuery(sTimberlineConnString, sCommText);

                sb.Append(Environment.NewLine);
                sb.AppendFormat(" <Merchandise Quantity=\"" + sQuantity + "\" ItemID=\"item_id_" + sItemCount + "\">");
                sb.Append(Environment.NewLine);
                sb.AppendFormat("   <PartNumber>" + sPartNum + "</PartNumber>");
                sb.Append(Environment.NewLine);

                sCommTextValue = string.Empty;
                sColumn = string.Empty;
                string sPrintMode = string.Empty;
                dtGenericValue.Clear();

                if (dtGenericValue.Columns.Count > 0)
                {
                    for (int i = dtGenericValue.Columns.Count - 1; i >= 0; i--)
                    {
                        string sColumnToRemove = Convert.ToString(dtGenericValue.Columns[0].ColumnName.Trim());
                        dtGenericValue.Columns.Remove(sColumnToRemove);
                    }
                }

                sCommTextValue = "SELECT DISTINCT [PrintMode] FROM [Codes] WHERE [CDSCode] = '" + sCode + "'";
                sColumn = "PrintMode";
                TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
                sPrintMode = Convert.ToString(sValueString);

                sb.AppendFormat(@"   <Printmode>" + sPrintMode + "</Printmode>"); 
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"   <ImageSet>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat("     <Image URL=\"" + sImagePath + "\" ViewName=\"front\" />");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"   </ImageSet>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"   <!-- Required, but unused -->");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"   <PackageListOverride>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"     <ItemName></ItemName>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"     <ItemDescription></ItemDescription>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"   </PackageListOverride>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@" </Merchandise>");
                sb.Append(Environment.NewLine);
                sb.AppendFormat(@"</Order>");

                File.AppendAllText(sXMLFinal, sb.ToString());

                TimeSpan ts = new TimeSpan(0, 0, 3);
                Thread.Sleep(ts); // Sleep thread 3 seconds to allow file creation.

                if (File.Exists(sXMLFinal))
                {
                    string sDTNow2 = DateTime.Now.ToString();
                    sCommText = "UPDATE [OrderItems] SET [XmlFileCreated] = 'T', [LocalXMLFilePath] = '" + sXMLFinal + "', [XmlFileCreatedTime] = '" + sDTNow2 +
                        "' WHERE [UID] = '" + sUID + "'";
                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                }
                else if (!File.Exists(sXMLFinal))
                {
                    sCommText = "UPDATE [OrderItems] SET [XmlFileCreated] = 'E' WHERE [RefNum] = '" + sRefNum + "' AND [ItemID] = '" + sOrderItemID + "'";
                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                    bXMLFileError = true;
                }

                // Post current XML.
                if (bXMLFileError != true)
                {
                    string sLocalOrderXML = sXMLFinal;
                    this.POSTXML(sLocalOrderXML, sUID);
                }
                else if (bXMLFileError == true)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void GatherXMLData(ref DataTable dtShippingData, string sProdNum, string sFrameNum) // Task method for above method. 
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT Attn, Address1, Address2, City, State, Zip, Phone FROM Directship WHERE Lookupnum = '" + sProdNum + "' AND Sequence = " + sFrameNum;

                TM.CDSQuery(sCDSConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    dtShippingData.Clear();
                    dtShippingData = dTbl.Copy();
                    dTbl.Clear();
                }
                else if (dTbl.Rows.Count == 0)
                {
                    bReturned = true;
                    return;
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
                bXMLFileError = true;
                return;
            }
        }

        #endregion

        #region Post [Post order XML to TimberlineColorado and gather reply.]

        private void POSTXML(string sLocalOrderXML, string sUID) // Query the orderitems table for xml files to POST. 
        {
            DataTable dtGenericValue = new DataTable();
            string sCommTextValue = string.Empty;
            string sValueString = string.Empty;
            string sColumn = string.Empty;
            string sUrl = string.Empty;
            string sRefNum = string.Empty;
            string sFrame = string.Empty;
            string sItemID = string.Empty;
            string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");

            // Gather HTTP POST URL.
            sCommTextValue = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'XMLPostURL'";
            sColumn = "Variables_Variable";
            TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
            sUrl = Convert.ToString(sValueString);

            // POST XML file.
            this.postXMLData(sUrl, sLocalOrderXML, sUID);
        }

        public string postXMLData(string sUrl, string sLocalXML, string sUID) // Send the order xml via HTTP POST and gather reply. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HH:mm:ss");
            string sStatDTNow = "[" + sDTime1 + "][" + sDTime2 + "]";
            string sStatTxt = sStatDTNow + " POSTing XML file for UID: " + sUID;
            bool bFinished = false;
            sBStatus.AppendFormat(sStatTxt);
            sBStatus.AppendFormat(Environment.NewLine);
            this.SetStatusText(sStatTxt, sStatDTNow, bFinished);


            string sXMLPath = string.Empty;
            sXMLPath = Path.GetDirectoryName(sLocalXML);
            bool bReadyToMove = false;

            // Disregard cert warning on connection.
            ServicePointManager.ServerCertificateValidationCallback = new System.Net.Security.RemoteCertificateValidationCallback
            (
                delegate { return true; }
            );

            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls;
            //ServicePointManager.SecurityProtocol = SecurityProtocolType.Ssl3;

            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sUrl);
                byte[] bytes;
                sLocalXML = File.ReadAllText(sLocalXML);
                bytes = Encoding.UTF8.GetBytes(sLocalXML);
                request.ContentType = "text/xml";
                request.Accept = "text/xml";                
                request.ContentLength = bytes.Length;
                request.Method = "POST";
                request.KeepAlive = false;
                Stream requestStream = request.GetRequestStream();                
                requestStream.Write(bytes, 0, bytes.Length);
                requestStream.Close();
                HttpWebResponse response;
                response = (HttpWebResponse)request.GetResponse();                
                if (response.StatusCode == HttpStatusCode.OK)
                {
                    Stream responseStream = response.GetResponseStream();
                    string sResponseStr = new StreamReader(responseStream).ReadToEnd();

                    // Write the response back out to an xml and store it.
                    string sAcceptedXML = string.Empty;
                    sAcceptedXML = sXMLPath + @"\" + sUID + "-ACCEPTED.xml";
                    if (!File.Exists(sAcceptedXML))
                    {
                        File.WriteAllText(sAcceptedXML, sResponseStr);
                        bReadyToMove = true;

                        // Update record that XML was posted.
                        string sCommText = string.Empty;
                        string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                        sCommText = "UPDATE [OrderItems] SET [XmlFilePosted] = 'T', [XmlFilePostedTime] = '" + sDTNow + "', [XmlFileAccepted] = 'T', [XmlFileAcceptedPath] = '" + sAcceptedXML + "' WHERE [UID] = '" + sUID + "'";
                        TM.SQLNonQuery(sTimberlineConnString, sCommText);
                        
                        string sDestDir = sXMLPath.Replace("Pending", "Posted");
                        if (bReadyToMove != false)
                        {
                            if (!Directory.Exists(sDestDir))
                            {
                                Directory.Move(sXMLPath, sDestDir);

                                sDTNow = DateTime.Now.ToString();
                                DataTable dtGeneric = new DataTable();
                                sCommText = "SELECT * FROM [OrderItems] WHERE [UID] = '" + sUID + "'";
                                TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);

                                string sLocalXMLFilePath = Convert.ToString(dtGeneric.Rows[0]["LocalXMLFilePath"]).Trim();
                                sLocalXMLFilePath = sLocalXMLFilePath.Replace("Pending", "Posted");
                                sAcceptedXML = sAcceptedXML.Replace("Pending", "Posted");

                                sCommText = "UPDATE [OrderItems] SET [LocalXMLFilePath] = '" + sLocalXMLFilePath + "', [XmlFileAcceptedPath] = '" + sAcceptedXML +
                                "' WHERE [UID] = '" + sUID + "'";
                                TM.SQLNonQuery(sTimberlineConnString, sCommText);

                                XmlDocument xmlDoc = new XmlDocument();
                                xmlDoc.Load(sAcceptedXML);
                                XmlNode root = xmlDoc.DocumentElement;
                                XmlNodeList xmlNL = root.SelectNodes("//success/confirmation_number");
                                XmlNodeList xmlNL2 = root.SelectNodes("//success/po");
                                string sConfNum = string.Empty;
                                string sPONum = string.Empty;
                                sConfNum = Convert.ToString(xmlNL[0].InnerText).Trim();
                                sPONum = Convert.ToString(xmlNL2[0].InnerText).Trim();

                                string sCommText2 = "UPDATE [OrderItems] Set [TCConfNum] = '" + sConfNum + "', [TCPONum] = '" + sPONum + "' WHERE [UID] = '" + sUID + "'";
                                TM.SQLNonQuery(sTimberlineConnString, sCommText2);

                                // Note: Email upon successful posting.

                                string sMySubject = string.Empty;
                                string sMyBody = string.Empty;

                                TM.SendNotificationOrderSentToTimberline(sUID, ref sMySubject, ref sMyBody);
                                
                                EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);

                            }
                            else if (Directory.Exists(sDestDir))
                            {
                                // This should never happen.
                            }
                        }
                    }
                    else if (File.Exists(sAcceptedXML))
                    {
                        // This should never happen.
                    }
                }
                else
                {                  
                    // If any errors this will trigger a WebException which is caught below.
                }

                return null;
            }
            catch (WebException ex)
            {
                bReadyToMove = false;

                string sPageContent = new StreamReader(ex.Response.GetResponseStream()).ReadToEnd().ToString();

                string sErrorXML = string.Empty;
                sErrorXML = sXMLPath + @"\" + sUID + "-ERROR.xml";
                if (!File.Exists(sErrorXML))
                {
                    File.WriteAllText(sErrorXML, sPageContent);
                    bReadyToMove = true;

                    // Update record that XML posting returned an error.
                    string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                    string sCommText = "UPDATE [OrderItems] SET [XmlFilePosted] = 'T', [XmlFilePostedTime] = '" + sDTNow + "', [XmlFileAccepted] = 'F', [XmlErrorFilePath] = '" + sErrorXML +
                        "' WHERE [UID] = '" + sUID + "'";
                    TM.SQLNonQuery(sTimberlineConnString, sCommText);

                    string sMySubject = "Timberline Error: XML Post error.";
                    string sMyBody = "XML Post error for UID #: " + sUID + "'" + Environment.NewLine + sPageContent;

                    EM.SendEmail(sEmailServer, sMyEmail, sEmailBCC, sMySubject, sMyBody);

                    string sDestDir = sXMLPath.Replace("Pending", "Error");

                    if (bReadyToMove != false)
                    {
                        if (!Directory.Exists(sDestDir))
                        {
                            Directory.Move(sXMLPath, sDestDir);

                            sDTNow = DateTime.Now.ToString();
                            DataTable dtGeneric = new DataTable();
                            sCommText = "SELECT * FROM [OrderItems] WHERE [UID] = '" + sUID + "'";
                            TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);

                            string sLocalXMLFilePath = Convert.ToString(dtGeneric.Rows[0]["LocalXMLFilePath"]).Trim();
                            string sXmlErrorFilePath = Convert.ToString(dtGeneric.Rows[0]["XmlErrorFilePath"]).Trim();
                            sLocalXMLFilePath = sLocalXMLFilePath.Replace("Pending", "Error");
                            sXmlErrorFilePath = sXmlErrorFilePath.Replace("Pending", "Error");

                            sCommText = "UPDATE [OrderItems] SET [LocalXMLFilePath] = '" + sLocalXMLFilePath + "', [XmlErrorFilePath] = '" + sXmlErrorFilePath +
                            "' WHERE [UID] = '" + sUID + "'";
                            TM.SQLNonQuery(sTimberlineConnString, sCommText);
                        }
                        else if (Directory.Exists(sDestDir))
                        {
                            // This should never happen.
                        }
                    }
                }
                else if (File.Exists(sErrorXML))
                {
                    // This should never happen.
                }

                XmlDocument xmlDoc = new XmlDocument();
                sErrorXML = sErrorXML.Replace("Pending", "Error");
                xmlDoc.Load(sErrorXML);
                XmlNode root = xmlDoc.DocumentElement;
                XmlNodeList xmlNL = root.SelectNodes("//errors/error");
                string sError = string.Empty;
                sError = Convert.ToString(xmlNL[0].InnerText).Trim();

                string sCommText2 = "INSERT INTO [Errors] VALUES ('XML POST Error: " + sError + "', '" + DateTime.Now + "', '0')";
                TM.SQLNonQuery(sTimberlineConnString, sCommText2);

                return sPageContent;
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
                return ex.ToString();
            }
        }

        #endregion        

        #region Misc methods.

        private void Clear() // Clear global variables. 
        {
            bCreated = false;
            dtResults.Clear();
            dtResults.Dispose();
            dtScanned.Clear();
            dtScanned.Dispose();
            dtScannedAddInfo.Clear();
            dtScannedAddInfo.Dispose();
            dtFinal.Clear();
            dtFinal.Dispose();
            bReturned = false;
            bRenderError = false;
            bAutoGenFileError = false;
            sBStatus.Clear();
            rTxtBoxStatus.Clear();
            rTxtBoxStatus.Refresh();
            bBatchTimer = false;
            bXMLFileError = false;
            bNoGatherData = false;
            bGotCode = false;
        }  

        private void SetStatusText(string sStatTxt, string sStatDTNow, bool bFinished) // Set form status text. 
        {
            try
            {
                if (bFinished != true)
                {
                    rTxtBoxStatus.AppendText(sStatTxt + Environment.NewLine);
                }
                else if (bFinished == true)
                {
                    rTxtBoxStatus.AppendText(sStatTxt + Environment.NewLine);
                    this.ExportLogTextToFile();
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        private void ExportLogTextToFile() // Save session log to file. 
        {
            string sDTime1 = DateTime.Now.ToString("MM-dd-yy");
            string sDTime2 = DateTime.Now.ToString("HHmmss");
            string sDTime3 = @"[" + sDTime1 + @"][" + sDTime2 + @"]";

            DataTable dtGenericValue = new DataTable();
            string sCommTextValue = "SELECT * FROM [Variables] WHERE [Variables_Label] = 'TCLooperLogPath'";
            string sColumn = "Variables_Variable";
            string sValueString = string.Empty;

            TM.SQLQueryWithReturnValue(sTimberlineConnString, sCommTextValue, dtGenericValue, ref sValueString, sColumn);
            string sLogPath = sValueString;

            string sLogFile = sLogPath + sDTime3 + @"-TimberlineColorado.txt";
            int iMaxLogFileSize = 10485760; // 10MB
            int iFileSize = 0;

            DataTable dtGeneric = new DataTable();
            string sCommText = string.Empty;
            sCommText = "SELECT * FROM [Variables] WHERE [Variables_Label] = 'TCLooperLogFile'";
            TM.SQLQuery(sTimberlineConnString, sCommText, dtGeneric); ;
            string sCurrentLogFile = Convert.ToString(dtGeneric.Rows[0]["Variables_Variable"]).Trim();
            try
            {
                if (sCurrentLogFile == string.Empty || !File.Exists(sCurrentLogFile))
                {
                    File.WriteAllText(sLogFile, sBStatus.ToString());

                    sCurrentLogFile = sLogFile;

                    sCommText = "UPDATE [Variables] SET [Variables_Variable] = '" + sCurrentLogFile + "' WHERE [Variables_Label] = 'TCLooperLogFile'";

                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                }
                else if (sCurrentLogFile != string.Empty && File.Exists(sCurrentLogFile))
                {
                    iFileSize = sCurrentLogFile.Length;
                    if (iFileSize >= iMaxLogFileSize)
                    {
                        string sAppendingText = Environment.NewLine + "Log file has reached maximum size. Creating a new log file.";
                        File.AppendAllText(sCurrentLogFile, sAppendingText);

                        File.WriteAllText(sLogFile, sBStatus.ToString());
                        sCurrentLogFile = sLogFile;

                        sCommText = "UPDATE [Variables] SET [Variables_Variable] = '" + sCurrentLogFile + "' WHERE [Variables_Label] = 'TCLooperLogFile'";

                        TM.SQLNonQuery(sTimberlineConnString, sCommText);
                    }
                    else if (iFileSize <= iMaxLogFileSize)
                    {
                        File.AppendAllText(sCurrentLogFile, sBStatus.ToString());
                    }
                }
                else
                {
                    DateTime dTimeError = DateTime.Now;

                    string sLogError = string.Empty;

                    sLogError = "Could not generate a log file: Values are sCurrentLogFile = [" + sCurrentLogFile + "], sLogFile = [" + sLogFile +
                        "], iFileSize = [" + iFileSize + "], sbStatus.ToString = [" + sBStatus.ToString() + "]";

                    sCommText = "INSERT INTO [Errors] VALUES ('" + sLogError + "', '" + dTimeError + "', '0')";

                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                }
            }
            catch (Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        } 

        private void CheckForShipped()
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [OrderItems] WHERE [XmlFileShippedPath] IS NOT NULL";

                TM.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    string sUID = Convert.ToString(dTbl.Rows[0]["UID"]).Trim();
                    string sLocalXMLFilePath = Convert.ToString(dTbl.Rows[0]["LocalXMLFilePath"]).Trim();
                    string sLocalXMLFilePathShipped = sLocalXMLFilePath.Replace("Posted", "Shipped").Trim();
                    string sXmlFileProcessedPath = Convert.ToString(dTbl.Rows[0]["XmlFileShippedPath"]).Trim();
                    string sXmlFileShippedPath = sXmlFileProcessedPath.Replace("Processed", "Shipped").Trim();
                    string sXmlFileAcceptedPath = Convert.ToString(dTbl.Rows[0]["XmlFileAcceptedPath"]).Trim();
                    string sNewXmlFileAcceptedPath = sXmlFileAcceptedPath.Replace("Posted", "Shipped").Trim();

                    string sPostedPath = Path.GetDirectoryName(sLocalXMLFilePath).Trim();
                    string sShippedPath = sPostedPath.Replace("Posted", "Shipped").Trim();

                    if (Directory.Exists(sPostedPath))
                    {
                        Directory.Move(sPostedPath, sShippedPath);
                    }
                    else if (!Directory.Exists(sPostedPath))
                    {

                    }

                    if (File.Exists(sXmlFileProcessedPath))
                    {
                        File.Move(sXmlFileProcessedPath, sXmlFileShippedPath);
                    }
                    else if (!File.Exists(sXmlFileProcessedPath))
                    {

                    }

                    sCommText = "UPDATE [OrderItems] SET [LocalXMLFilePath] = '" + sLocalXMLFilePathShipped + "', [XmlFileAcceptedPath] = '" + sNewXmlFileAcceptedPath + "'," +
                        " [XmlFileShippedPath] = '" + sXmlFileShippedPath + "' WHERE [UID] = '" + sUID + "'";

                    TM.SQLNonQuery(sTimberlineConnString, sCommText);
                }
            }
            catch(Exception ex)
            {
                TM.SaveExceptionToDB(ex);
            }
        }

        #endregion

    }
}
