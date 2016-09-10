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

namespace TimberlineLooper
{
    public partial class TimberlineLooper_Main : Form
    {

        bool bCreated = false;
        DataTable dtResults = new DataTable();
        DataTable dtScannedAddInfo = new DataTable();
        DataTable dtScanned = new DataTable();
        
        public TimberlineLooper_Main()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DoWork();
        }

        private void Clear()
        {
            bCreated = false;
            dtResults.Clear();
            dtResults.Dispose();
            dtScanned.Clear();
            dtScanned.Dispose();
            dtScannedAddInfo.Clear();
            dtScannedAddInfo.Dispose();
        }

        private void ExceptionEmail()
        {
            DataTable dt = new DataTable();

            try
            {
                string sExceptionString = "";
                DateTime datetimeException = DateTime.Now;
                string sDateTimeException = "";

                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "SELECT * FROM [OSRetouch].[dbo].[OSR_Errors] WHERE [Errors_Email_Sent] = '0' OR [Errors_Email_Sent] IS NULL";

                myConn.Open();

                myCommand.CommandTimeout = 0;

                SqlDataReader myReader = myCommand.ExecuteReader();

                if (myReader.HasRows)
                {
                    dt.Load(myReader);

                    foreach (DataRow dr in dt.Rows)
                    {
                        sExceptionString = Convert.ToString(dr["Errors_String"]).Trim();
                        datetimeException = Convert.ToDateTime(dr["Errors_DateTime"]);
                        sDateTimeException = Convert.ToString(dr["Errors_DateTime"]);

                        string sMysubject = string.Format("Timberline Error Reporting");
                        string sMybody = string.Format("An exception was recorded in the Errors database at " + datetimeException + ":" + "\n" + "\n" + sExceptionString);

                        string sEmailServer = "email";
                        string sEmailMyBccAdd = "jlett@company.mail";
                        string sEmailMyErrorSendAdd = "thegrump1976@gmail.com";

                        TimberlineLooper.Email.EmailError(sEmailServer, sEmailMyBccAdd, sEmailMyErrorSendAdd, sMysubject, sMybody);

                        this.ExceptionEmailSent(sDateTimeException);

                    }

                }

                myReader.Close();
                myReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();

                dt.Clear();
                dt.Dispose();

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        // Capture string and datetime from above method and update the Errors_Email_Sent field to 1 (Sent)
        private void ExceptionEmailSent(string sDateTimeException)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "UPDATE [OSRetouch].[dbo].[OSR_Errors] SET [Errors_Email_Sent] = '1'";

                myConn.Open();

                myCommand.ExecuteNonQuery();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void SaveExceptionToDB(Exception ex)
        {
            DateTime dtime = DateTime.Now;

            string sException = ex.ToString().Trim();
            sException = sException.Replace(@"'", "");

            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "INSERT INTO [Errors] VALUES ('" + sException + "', '" + dtime + "', '0')";

                myConn.Open();

                myCommand.ExecuteNonQuery();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();

            }
            catch (Exception exception)
            {
                MessageBox.Show(exception.Message);
            }
        }


        // Loop through ID table for customers
        // query Items for work based on said customers - gather ref num
        // query Stamps for trigger points indicating ready work
        // query Labels for products in preconfig or Timberline codes
        // store above in datatable then query Codes for sequence, quantity, lookupnum
        // query DirectShip for shipping information
        // query Sport ????????????????????????? - not sure if needed
        // query Endcust for Email
        // query Group for group data ????????????????? - not sure if needed


        private void DoWork()
        {
            this.Clear();
            this.SearchItems();
            this.ExistsInDP2();
            this.SaveResultsToVerifiedTable();
            this.LoopThroughVerifiedTable();
            this.ExceptionEmail();
        }

        private void SearchItems()
        {
            DataTable dtItems = new DataTable();

            DateTime dTimeNow = DateTime.Now;
            DateTime dTNowDateOnly = dTimeNow.Date;
            DateTime dTimeNowMinusDays = DateTime.Now.AddDays(-60);
            DateTime dTimeNowMinusDaysDateOnly = dTimeNowMinusDays.Date;

            string sDTNowDateOnly = dTNowDateOnly.ToString("MM/dd/yy");
            string sDTNowMinusDaysDateOnly = dTimeNowMinusDaysDateOnly.ToString("MM/dd/yy");

            try
            {
                OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = "SELECT Lookupnum FROM ITEMS WHERE items.d_dueout > CTOD('" + sDTNowMinusDaysDateOnly + "') AND ITEMS.PACKAGETAG IN" +
                " (SELECT PACKAGETAG FROM LABELS WHERE (LABELS.CODE = 'CFB' OR LABELS.CODE = 'CCA' OR LABELS.CODE = 'CPC') AND LABELS.PACKAGETAG <> '    ') ORDER BY items.d_dueout ASC";

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dtItems.Load(CDSreader);

                    int iCount = dtItems.Rows.Count;
                }

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();

                if (dtItems != null)
                {
                    this.ScanForTriggerPoints(dtItems);
                }                

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void ScanForTriggerPoints(DataTable dtItems)
        {            
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
                    string sProdNum = Convert.ToString(dr["lookupnum"]).Trim();

                    OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                    OleDbCommand CDScommand = CDSconn.CreateCommand();

                    CDScommand.CommandText = "SELECT * FROM STAMPS WHERE Lookupnum = '" + sProdNum +
                    "' AND (Action = " + "'DIGI PRINT'" + " OR Action = '" + "PRNT_TRAV" + "' OR Action = " + "'COLOR QC')";

                    CDSconn.Open();

                    CDScommand.CommandTimeout = 0;

                    OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                    if (CDSreader.HasRows)
                    {
                        dt.Load(CDSreader);
                        dRow = dtScanned.NewRow();
                        dRow["ProdNum"] = sProdNum;
                        dtScanned.Rows.Add(dRow);
                    }

                    CDScommand.Dispose();

                    CDSreader.Close();
                    CDSreader.Dispose();

                    CDSconn.Close();
                    CDSconn.Dispose();
                }

                if (dtScanned != null)
                {
                    this.GetAdditionalInfoForDTScanned();
                }

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void GetAdditionalInfoForDTScanned()
        {
            try
            {
                foreach (DataRow dr in dtScanned.Rows)
                {
                    string sProdNum = Convert.ToString(dr["ProdNum"]).Trim();                    

                    OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                    OleDbCommand CDScommand = CDSconn.CreateCommand();

                    CDScommand.CommandText = "SELECT Lookupnum, Order, Packagetag FROM ITEMS WHERE Lookupnum = '" + sProdNum + "'";

                    CDSconn.Open();

                    CDScommand.CommandTimeout = 0;

                    OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                    if (CDSreader.HasRows)
                    {
                        dtScannedAddInfo.Load(CDSreader);
                    }

                    CDScommand.Dispose();

                    CDSreader.Close();
                    CDSreader.Dispose();

                    CDSconn.Close();
                    CDSconn.Dispose();
                }

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void ExistsInDP2()
        {
            DataTable dt = new DataTable();

            try
            {
                foreach (DataRow dr in dtScannedAddInfo.Rows)
                {
                    string sRefNum = Convert.ToString(dr["Order"]).Trim();
                    string sProdNum = Convert.ToString(dr["lookupnum"]).Trim();
                    string sPkgTag = Convert.ToString(dr["Packagetag"]).Trim();                    

                    SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.DP2ConnString);

                    SqlCommand myCommand = myConn.CreateCommand();

                    myCommand.CommandText = "SELECT [ID] FROM [Orders] WHERE [ID] = '" + sRefNum + "'";

                    myConn.Open();

                    SqlDataReader myDataReader = myCommand.ExecuteReader();

                    if (myDataReader.HasRows)
                    {
                        dt.Load(myDataReader);

                        this.CheckForFrames(sRefNum);
                        this.CheckCodes(sProdNum, sPkgTag);

                    }
                    else
                    {
                        // Do nothing if no record is found to continue loop.
                    }

                    myDataReader.Close();
                    myDataReader.Dispose();

                    myCommand.Dispose();

                    myConn.Close();
                    myConn.Dispose();

                }
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void CheckForFrames(string sRefNum)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.DP2ConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "SELECT * FROM [Images] WHERE [OrderID] = '" + sRefNum + "'";

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    // do nothing
                }
                else
                {
                    myDataReader.Close();
                    myDataReader.Dispose();

                    myCommand.Dispose();

                    myConn.Close();
                    myConn.Dispose();

                    return;
                }


                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void CheckCodes(string sProdNum, string sPkgTag)
        {

            try
            {
                OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = "SELECT * FROM Codes WHERE Lookupnum = '" + sProdNum + "' AND ((Codes.Code = 'CFB' OR Codes.Code = 'CCA' OR Codes.Code = 'CPC' AND Package = .F.) OR (Package = .T. AND"
                + " Code IN (SELECT DISTINCT Packagecod FROM Labels WHERE Labels.packagetag = '" + sPkgTag + "' AND (Labels.Code = 'CFB' OR Labels.Code = 'CCA' OR Labels.Code = 'CPC')))) ORDER BY Sequence ASC";

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dtResults.Load(CDSreader);
                }
                else
                {
                    CDScommand.Dispose();

                    CDSreader.Close();
                    CDSreader.Dispose();

                    CDSconn.Close();
                    CDSconn.Dispose();
                }

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void SaveResultsToVerifiedTable()
        {

            DataTable dt = new DataTable();

            int iCount = dtResults.Rows.Count;

            string sCommTxt = "INSERT INTO [Timberline].[dbo].[Verified] ([ProdNum], [RefNum], [FrameNum], [InDP2], [LastCheck], [CustNum], [PkgTag], [ServiceType], [Quantity])" +
                " VALUES (@A, @B, @C, 'T', @D, @E, @F, @G, @H)";

            if (dtResults != null)
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
                                string sQuantity = Convert.ToString(dr["quantity"]).Trim();

                                this.GetRefNumForVerified(sProdNum, ref dt);

                                string sRefNum = Convert.ToString(dt.Rows[0]["order"]).Trim();
                                string sCustNum = Convert.ToString(dt.Rows[0]["customer"]).Trim();
                                string sPkgTag = Convert.ToString(dt.Rows[0]["packagetag"]).Trim();
                                string sServiceType = Convert.ToString(dt.Rows[0]["Sertype"]).Trim();

                                SqlCommand myCommand = myConn.CreateCommand();
                                myCommand.CommandText = sCommTxt;
                                myCommand.Parameters.AddWithValue("@A", sProdNum);
                                myCommand.Parameters.AddWithValue("@B", sRefNum);
                                myCommand.Parameters.AddWithValue("@C", sFrameNum);
                                DateTime dTNow = DateTime.Now;
                                string sDTNow = dTNow.ToString("MM/dd/yy H:mm:ss");
                                myCommand.Parameters.AddWithValue("@D", sDTNow);
                                myCommand.Parameters.AddWithValue("@E", sCustNum);
                                myCommand.Parameters.AddWithValue("@F", sPkgTag);
                                myCommand.Parameters.AddWithValue("@G", sServiceType);
                                myCommand.Parameters.AddWithValue("@H", sQuantity);
                                myCommand.ExecuteNonQuery();
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
                    this.SaveExceptionToDB(ex);
                }
            }
        }

        private void GetRefNumForVerified(string sProdNum, ref DataTable dt)
        {
            try
            {
                OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = "SELECT Order, Customer, Packagetag, Sertype FROM ITEMS WHERE Lookupnum = '" + sProdNum + "'";

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dt.Clear();
                    dt.Load(CDSreader);
                }
                else
                {
                    CDScommand.Dispose();

                    CDSreader.Close();
                    CDSreader.Dispose();

                    CDSconn.Close();
                    CDSconn.Dispose();

                    return;
                }

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void LoopThroughVerifiedTable()
        {
            DataTable dt = new DataTable();

            DataTable dtXML = new DataTable();

            DataTable dtSorted = new DataTable();

            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();
                                
                myCommand.CommandText = "SELECT * FROM [Verified] WHERE [QueuedToRender] IS NULL AND [ItemID] IS NULL";

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dt.Load(myDataReader);

                    string sRefNum = Convert.ToString(dt.Rows[0]["RefNum"]);
                    string sFrame = Convert.ToString(dt.Rows[0]["FrameNum"]);
                    string sFrameNumber = string.Empty;
                    string sProdNum = Convert.ToString(dt.Rows[0]["ProdNum"]);
                    string sQuantity = Convert.ToString(dt.Rows[0]["Quantity"]);
                    string sFrameNum = sFrameNumber.TrimStart(new Char[] { '0' });

                    this.GetOrderItemsToRender2(ref sRefNum, ref dtSorted, ref sFrameNumber, ref sFrame, ref sProdNum, ref sQuantity, ref sFrameNum);

                    #region old code

                    //this.GatherXMLData(sProdNum, sFrameNum, sQuantity, ref dtSorted, sRefNum);

                    //foreach (DataRow dr in dt.Rows) // is this foreach needed or should it just hit the first record and process that
                    //{
                    //    string sRefNum = Convert.ToString(dr["RefNum"]).Trim();
                    //    string sFrameNumber = Convert.ToString(dr["FrameNum"]).Trim();

                    //    this.GetOrderItemsToRender2(sRefNum, ref dtSorted, sFrameNumber);

                    //    string sProdNum = Convert.ToString(dr["ProdNum"]).Trim();
                    //    string sFrameNum = Convert.ToString(dr["FrameNum"]).Trim();
                    //    string sQuantity = Convert.ToString(dr["Quantity"]).Trim();
                    //    sFrameNum = sFrameNum.TrimStart(new Char[] { '0' });

                    //    this.GatherXMLData(sProdNum, sFrameNum, sQuantity, ref dtSorted, sRefNum);
                    //}

                    #endregion

                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void GetOrderItemsToRender2(ref string sRefNum, ref DataTable dtSorted, ref string sFrameNumber, ref string sFrame, ref string sProdNum, ref string sQuantity, ref string sFrameNum)
        {
            DataTable dtToRender1 = new DataTable();

            DataTable dtCodes = new DataTable();

            this.QueryCodes(ref dtCodes);

            try
            {
                foreach (DataRow drCodes in dtCodes.Rows)
                {

                    string sCode = Convert.ToString(drCodes["CDSCode"]).Trim();

                    SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.DP2ConnString);

                    SqlCommand myCommand = myConn.CreateCommand();

                    myCommand.CommandText = "SELECT [OrderID], [ID], [Sequence], [Quantity], [ProductID] FROM [OrderItems] WHERE [OrderID] = '" + sRefNum + "' AND [ProductID] Like '" + sCode + "%'";

                    myConn.Open();

                    SqlDataReader myDataReader = myCommand.ExecuteReader();

                    if (myDataReader.HasRows)
                    {
                        dtToRender1.Clear();
                        dtToRender1.Load(myDataReader);

                        // load another datatable here that collects all orderitems for a single frame if multiple are ordered
                        // either in a datatable that doesn't get cleared until leaving the method or in a string

                        this.EliminateDuplicateOrderItemsForRendering(sRefNum, ref dtToRender1, ref sFrameNumber, ref dtSorted, ref sFrame, ref sCode);

                        foreach (DataRow dr in dtSorted.Rows)
                        {
                            string sSequence = Convert.ToString(dr["ItemID"]).Trim();

                            this.GenerateAutoGenFile(sRefNum, sSequence, sFrameNumber);

                            this.GatherXMLData(sProdNum, sFrameNum, sQuantity, ref dtSorted, sRefNum);
                        } 
                    }

                    myDataReader.Close();
                    myDataReader.Dispose();

                    myCommand.Dispose();

                    myConn.Close();
                    myConn.Dispose();
                }
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void EliminateDuplicateOrderItemsForRendering(string sRefNum, ref DataTable dtToRender1, ref string sFrameNumber, ref DataTable dtSorted, ref string sFrame, ref string sCode)
        {
            DataTable dt = new DataTable();

            string sID = string.Empty;
            sID = Convert.ToString(dtToRender1.Rows[0]["ID"]).Trim();
            string sItemID = sID;

            this.SaveItemIDforVerified(sRefNum, sFrame, ref sItemID);

            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.DP2ConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "SELECT [OrderID], [Roll], [Frame], [ItemOrderID], [ItemID] FROM [OrderItemImages] WHERE [OrderID] = '" + sRefNum + "' AND [Frame] = '" +
                    sFrame + "' AND [ItemID] IN (" + sID + ")";                

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dt.Clear();
                    dt.Load(myDataReader);
                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }

            string sDistinctColumn = "Frame";

            DataView dv = dt.DefaultView;
            dv.Sort = "ItemID DESC";
            dtSorted.Clear();
            dtSorted = dv.ToTable();

            sFrameNumber = Convert.ToString(dtSorted.Rows[0]["Frame"]).Trim();

            this.RemoveDuplicateRows(dtSorted, sDistinctColumn);
        }

        private void SaveItemIDforVerified(string sRefNum, string sFrame, ref string sItemID)
        {
            DataTable dt = new DataTable();

            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "UPDATE [Verified] SET [ItemID] = '" + sItemID + "' WHERE [RefNum] = '" + sRefNum + "' AND [FrameNum] = '" + sFrame + "'";

                myConn.Open();

                myCommand.ExecuteNonQuery();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public DataTable RemoveDuplicateRows(DataTable dtSorted, string DistinctColumn)
        {
            try
            {
                ArrayList alUniqueRecords = new ArrayList();
                ArrayList alDupeRecords = new ArrayList();

                // Check if records is already added to UniqueRecords otherwise,
                // Add the records to DuplicateRecords
                foreach (DataRow dr in dtSorted.Rows)
                {
                    if (alUniqueRecords.Contains(dr[DistinctColumn]))
                        alDupeRecords.Add(dr);
                    else
                        alUniqueRecords.Add(dr[DistinctColumn]);
                }

                // Remove dupliate rows from DataTable added to DuplicateRecords
                foreach (DataRow dr in alDupeRecords)
                {
                    dtSorted.Rows.Remove(dr);
                }

                // Return the clean DataTable which contains unique records.
                return dtSorted;
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
                return null;
            }
        }

        private void QueryCodes(ref DataTable dtCodes)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "SELECT * FROM [Codes]";

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dtCodes.Load(myDataReader);
                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void GenerateAutoGenFile(string sRefNum, string sSequence, string sFrameNumber)
        {
            string sAutoGenPath = @"\\ws1\shared\autogen job files\CTExport-" + sRefNum + "-" + sSequence + ".txt";

            string sOriginalFile = @"\\ws1\vol1\tmp\CTExport-1.txt";

            string sEditedFile = @"\\ws1\vol1\tmp\CTExport-" + sRefNum + "-" + sSequence + ".txt"; // used for testing

            try
            {
                string sText = File.ReadAllText(sOriginalFile);
                sText = sText.Replace("RefNum", sRefNum);
                sText = sText.Replace("Seq", sSequence);
                File.WriteAllText(sAutoGenPath, sText);

                this.UpdateVerifiedRendered(sRefNum, sFrameNumber);

            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void UpdateVerifiedRendered(string sRefNum, string sFrameNumber)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "UPDATE [Verified] SET [QueuedToRender] = 'T' WHERE [RefNum] = '" + sRefNum + "' AND [FrameNum] = '" + sFrameNumber + "'";

                myConn.Open();

                myCommand.ExecuteNonQuery();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void GatherXMLData(string sProdNum, string sFrameNum, string sQuantity, ref DataTable dtSorted, string sRefNum)
        {
            DataTable dt = new DataTable();

            try
            {
                OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = "SELECT Attn, Address1, Address2, City, State, Zip, Phone FROM Directship WHERE Lookupnum = '" + sProdNum + "' AND Sequence = " + sFrameNum;

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dt.Clear();
                    dt.Load(CDSreader);

                    this.XMLWork(ref dt, sQuantity, ref dtSorted, sProdNum, sRefNum);
                }
                else
                {
                    CDScommand.Dispose();

                    CDSreader.Close();
                    CDSreader.Dispose();

                    CDSconn.Close();
                    CDSconn.Dispose();

                    return;
                }

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void XMLWork(ref DataTable dt, string sQuantity, ref DataTable dtSorted, string sProdNum, string sRefNum)
        {

            DateTime DTNow = DateTime.Now;
            string sDTNow = DTNow.ToString();

            DataTable dtOrdItemImages = new DataTable();
            DataTable dtImageLoc = new DataTable();
            DataTable dtProdId = new DataTable();

            StringBuilder sb = new StringBuilder();
            StringBuilder sb2 = new StringBuilder();

            string sXML = @"\\ws1\vol1\tmp\test.xml";

            foreach (DataRow drSorted in dtSorted.Rows)
            {
                string sFrameNum = Convert.ToString(drSorted["Frame"]).Trim();                
                string sItemID = Convert.ToString(drSorted["ItemID"]).Trim();

                string sXMLFinal = @"\\ws1\vol1\tmp\" + sRefNum + "-" + sFrameNum + ".xml";

                string sName = Convert.ToString(dt.Rows[0]["attn"]).Trim();
                sName = sName.Replace("&", "and");
                string sAddy1 = Convert.ToString(dt.Rows[0]["address1"]).Trim();
                string sAddy2 = Convert.ToString(dt.Rows[0]["address2"]).Trim();
                string sCity = Convert.ToString(dt.Rows[0]["city"]).Trim();
                string sState = Convert.ToString(dt.Rows[0]["state"]).Trim();
                string sZip = Convert.ToString(dt.Rows[0]["zip"]).Trim();
                string sPhone = Convert.ToString(dt.Rows[0]["phone"]).Trim();

                try
                {

                    string sXMLFile = File.ReadAllText(sXML);

                    sXMLFile = sXMLFile.Replace("APSOrderID", sRefNum);
                    sXMLFile = sXMLFile.Replace("APSAffiliateID", "CT_0019400");
                    sXMLFile = sXMLFile.Replace("APSShipToName", sName);
                    sXMLFile = sXMLFile.Replace("APSShipToContact", sName);
                    sXMLFile = sXMLFile.Replace("APSAddress1", sAddy1);
                    sXMLFile = sXMLFile.Replace("APSAddress2", sAddy2);
                    sXMLFile = sXMLFile.Replace("APSCity", sCity);
                    sXMLFile = sXMLFile.Replace("APSState", sState);
                    sXMLFile = sXMLFile.Replace("APSZip", sZip);
                    sXMLFile = sXMLFile.Replace("APSEmail", "");
                    sXMLFile = sXMLFile.Replace("APSPhone", sPhone);
                    sXMLFile = sXMLFile.Replace("APSShipMethod", "UPSGROUND");
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


                    int iRowCount = 0;
                    int iItems = 0;
                    string sItemCount = string.Empty;

                    string sLastFrame = string.Empty;

                    foreach (DataRow dr in dtSorted.Rows)
                    {
                        iRowCount = dtSorted.Rows.Count;

                        iItems++;
                        sItemCount = Convert.ToString(iItems).Trim();

                        if (iItems < 10 && iItems >= 1)
                        {
                            sItemCount = sItemCount.PadLeft(sItemCount.Length + 1, '0');
                        }

                        // need to query orderitems table against ref + sequence for ProductID
                        this.GetProductID(sRefNum, sItemID, ref dtProdId);

                        string sProdIDFull = Convert.ToString(dtProdId.Rows[0]["ProductID"]).Trim();
                        string sProdID = sProdIDFull.Substring(0, 3);

                        string sPartNum = string.Empty;

                        this.GetPartNumber(sProdID, ref sPartNum);

                        string sImageName = sRefNum + "_" + sProdNum + "_" + sFrameNum + ".jpg";

                        // query Items based on ref# for image location
                        this.GetImageLocation(sRefNum, ref dtImageLoc);

                        string sImageDir = Convert.ToString(dtImageLoc.Rows[0]["Imgloc"]).Trim();

                        // set up conditions for determining which sub folder to link images to
                        // structure on all rendered images (jpgs only) ProdNum\APSIMG\CTEXPORT\*CODE*\*RefNum*_*PNum*_*FrameNum*.jpg

                        string sImagePath = string.Empty; //Save this to Verified for the record and add a t/f column to verify whether rendering is done???????????????????????

                        if (sProdID == "CFB")
                        {
                            sImagePath = sImageDir + @"APSIMG\CTEXPORT\CFB\" + sImageName;
                        }

                        if (sProdID == "CCA")
                        {
                            sImagePath = sImageDir + @"APSIMG\CTEXPORT\CCA\" + sImageName;
                        }

                        if (sProdID == "CPC")
                        {
                            sImagePath = sImageDir + @"APSIMG\CTEXPORT\CPC\" + sImageName; 
                        }

                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(" <Merchandise Quantity=\"" + sQuantity + "\" ItemID=\"item_id_" + sItemCount + "\">");
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat("   <PartNumber>" + sPartNum + "</PartNumber>");
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"   <Printmode>dark</Printmode>"); // How to determine light or dark????????????????????????????
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"   <ImageSet>");
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat("     <Image URL=\"" + sImagePath + "\" ViewName=\"front\" />"); // Set image URL to current image for this record.
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

                        File.AppendAllText(sXMLFinal, sb.ToString());

                        sb.Clear();
                    }

                    sb2.Append(Environment.NewLine);
                    sb2.AppendFormat(@"</Order>");

                    File.AppendAllText(sXMLFinal, sb2.ToString());

                }
                catch (Exception ex)
                {
                    this.SaveExceptionToDB(ex);
                }
            }
        }

        private void GetPartNumber(string sProdID, ref string sPartNum)
        {
            DataTable dt = new DataTable();

            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.TimberlineConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "SELECT [PartNumber] FROM [Codes] WHERE [CDSCode] = '" + sProdID + "'";

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dt.Load(myDataReader);

                    sPartNum = Convert.ToString(dt.Rows[0]["PartNumber"]).Trim();
                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void GetImageLocation(string sRefNum, ref DataTable dtImageLoc)
        {
            try
            {
                OleDbConnection CDSconn = new OleDbConnection(TimberlineLooper.Properties.Settings.Default.CDSConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = "SELECT Imgloc FROM ITEMS WHERE Order = '" + sRefNum + "'"; 

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dtImageLoc.Clear();
                    dtImageLoc.Load(CDSreader);
                }
                else
                {
                    CDScommand.Dispose();

                    CDSreader.Close();
                    CDSreader.Dispose();

                    CDSconn.Close();
                    CDSconn.Dispose();
                }

                CDScommand.Dispose();

                CDSreader.Close();
                CDSreader.Dispose();

                CDSconn.Close();
                CDSconn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void GetProductID(string sRefNum, string sItemID, ref DataTable dtProdID)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(TimberlineLooper.Properties.Settings.Default.DP2ConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = "SELECT [ProductID] FROM [OrderItems] WHERE [OrderID] = '" + sRefNum + "' AND [Sequence] = '" + sItemID + "'";

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dtProdID.Clear();
                    dtProdID.Load(myDataReader);
                }

                myDataReader.Close();
                myDataReader.Dispose();

                myCommand.Dispose();

                myConn.Close();
                myConn.Dispose();
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }
    }
}
