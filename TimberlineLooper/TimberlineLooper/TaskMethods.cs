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
    class TaskMethods
    {
        string sTimberlineConnString = TimberlineLooper.Properties.Settings.Default.TimberlineConnString.ToString();
        string sDP2ConnString = TimberlineLooper.Properties.Settings.Default.DP2ConnString.ToString();
        string sCDSConnString =  TimberlineLooper.Properties.Settings.Default.CDSConnString.ToString();

        public void SQLNonQuery(string sConnString, string sCommText)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(sConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = sCommText;

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

        public void SQLQuery(string sConnString, string sCommText, DataTable dTbl)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(sConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = sCommText;

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dTbl.Clear();
                    dTbl.Load(myDataReader);
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

        public void SQLQueryWithReturnValue(string sConnString, string sCommTextValue, DataTable dTbl, ref string sValueString, string sColumn)
        {
            try
            {
                SqlConnection myConn = new SqlConnection(sConnString);

                SqlCommand myCommand = myConn.CreateCommand();

                myCommand.CommandText = sCommTextValue;

                myConn.Open();

                SqlDataReader myDataReader = myCommand.ExecuteReader();

                if (myDataReader.HasRows)
                {
                    dTbl.Clear();
                    dTbl.Load(myDataReader);

                    sValueString = string.Empty;
                    sValueString = Convert.ToString(dTbl.Rows[0][sColumn]).Trim();
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

        public void CDSQuery(string sConnString, string sCommText, DataTable dTbl)
        {
            try
            {
                OleDbConnection CDSconn = new OleDbConnection(sConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = sCommText;

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

                if (CDSreader.HasRows)
                {
                    dTbl.Clear();
                    dTbl.Load(CDSreader);
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

        public void CDSNonQuery(string sConnString, string sCommText)
        {
            try
            {
                OleDbConnection CDSconn = new OleDbConnection(sConnString);

                OleDbCommand CDScommand = CDSconn.CreateCommand();

                CDScommand.CommandText = sCommText;

                CDSconn.Open();

                CDScommand.CommandTimeout = 0;

                OleDbDataReader CDSreader = CDScommand.ExecuteReader();

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

        public void SaveExceptionToDB(Exception ex)
        {
            string sException = ex.ToString().Trim();
            sException = sException.Replace(@"'", "");

            string sCommText = "INSERT INTO [Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

            this.SQLNonQuery(sTimberlineConnString, sCommText);
        }

        public void EmailVariables(ref string sEmailServer, ref string sMyEmail, ref string sEmailBCC)
        {
            try
            {
                string sCommText = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'EmailServer'"; // Email server name.
                DataTable dt = new DataTable();

                this.SQLQuery(sTimberlineConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sEmailServer = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                sCommText = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'EmailSendTo'"; // My internal APS email.

                this.SQLQuery(sTimberlineConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sMyEmail = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }

                sCommText = "SELECT [Variables_Variable] FROM [Variables] WHERE [Variables_Label] = 'EmailBCCSend'"; // List of in lab email addys and my gmail for notification emails.

                this.SQLQuery(sTimberlineConnString, sCommText, dt);

                if (dt.Rows.Count > 0)
                {
                    sEmailBCC = Convert.ToString(dt.Rows[0]["Variables_Variable"]).Trim();
                }
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void SendNotificationOrderSentToTimberline(string sUID, ref string sMySubject, ref string sMyBody)
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [OrderItems] WHERE [UID] = '" + sUID + "'";

                this.SQLQuery(sTimberlineConnString, sCommText, dTbl);
                
                StringBuilder sBuilder = new StringBuilder();

                string sXMLProdNum = string.Empty;
                string sXMLRefNum = string.Empty;
                string sXMLFrameNum = string.Empty;
                string sXMLOrderItem = string.Empty;
                string sXMLQuantity = string.Empty;
                string sXMLOrderItemCode = string.Empty;
                string sXMLRenderedImagePath = string.Empty;
                string sXMLSentXMLPath = string.Empty;
                string sXMLPostedDT = string.Empty;
                string sXMLAcceptedXMLPath = string.Empty;
                string sXMLTCConfNum = string.Empty;
                string sXMLTCPONum = string.Empty;

                if (dTbl.Rows.Count > 0)
                {
                    sXMLProdNum = Convert.ToString(dTbl.Rows[0]["ProdNum"]).Trim();
                    sXMLRefNum = Convert.ToString(dTbl.Rows[0]["RefNum"]).Trim();
                    sXMLFrameNum = Convert.ToString(dTbl.Rows[0]["FrameNum"]).Trim();
                    sXMLOrderItem = Convert.ToString(dTbl.Rows[0]["ItemID"]).Trim();
                    sXMLQuantity = Convert.ToString(dTbl.Rows[0]["Quantity"]).Trim();
                    sXMLOrderItemCode = Convert.ToString(dTbl.Rows[0]["OrderItemCode"]).Trim();
                    sXMLRenderedImagePath = Convert.ToString(dTbl.Rows[0]["RenderedWebPath"]).Trim();
                    sXMLSentXMLPath = Convert.ToString(dTbl.Rows[0]["LocalXMLFilePath"]).Trim();
                    sXMLPostedDT = Convert.ToString(dTbl.Rows[0]["XmlFilePostedTime"]).Trim();
                    sXMLAcceptedXMLPath = Convert.ToString(dTbl.Rows[0]["XmlFileAcceptedPath"]).Trim();
                    sXMLTCConfNum = Convert.ToString(dTbl.Rows[0]["TCConfNum"]).Trim();
                    sXMLTCPONum = Convert.ToString(dTbl.Rows[0]["TCPONum"]).Trim();

                    sBuilder.AppendFormat("Production #: " + sXMLProdNum);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Reference #: " + sXMLRefNum);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Frame #: " + sXMLFrameNum);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("DP2 Order Item #: " + sXMLOrderItem);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Quantity: " + sXMLQuantity);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("DP2 Product ID: " + sXMLOrderItemCode);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Rendered Image Path: " + sXMLRenderedImagePath);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Sent XML Path: " + sXMLSentXMLPath);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("XML Posted: " + sXMLPostedDT);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Accepted XML Path: " + sXMLAcceptedXMLPath);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Timberline Confirmation #: " + sXMLTCConfNum);
                    sBuilder.Append(Environment.NewLine);
                    sBuilder.AppendFormat("Timberline PO #: " + sXMLTCPONum);
                    sBuilder.Append(Environment.NewLine);                    

                    sMySubject = "Timberline order transmitted.";
                    sMyBody = sBuilder.ToString();
                }
            }
            catch(Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void CheckForExceptions(ref string sMySubject, ref string sMyBody) // Check for logged errors in database and notify if new errors are found.
        {
            try
            {
                string sExceptionString = string.Empty;
                DateTime datetimeException = DateTime.Now;
                string sDateTimeException = string.Empty;

                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [Errors] WHERE [Errors_Email_Sent] = '0' OR [Errors_Email_Sent] IS NULL";
                this.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    foreach (DataRow dr in dTbl.Rows)
                    {
                        sExceptionString = Convert.ToString(dr["Errors_String"]).Trim();
                        datetimeException = Convert.ToDateTime(dr["Errors_DateTime"]);
                        sDateTimeException = Convert.ToString(dr["Errors_DateTime"]);

                        sMySubject = string.Format("Timberline Error Reporting");
                        sMyBody = string.Format("An exception was recorded in the Errors database at " + datetimeException + ":" + Environment.NewLine + Environment.NewLine + sExceptionString);                      

                        sCommText = "UPDATE [Errors] SET [Errors_Email_Sent] = '1' WHERE [Errors_DateTime] = '" + datetimeException + "'";
                        this.SQLNonQuery(sTimberlineConnString, sCommText);
                    }
                }
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void CheckOrderItemsForErrors(ref string sMySubject, ref string sMyBody) // Check for error flags with the OrderItems table. 
        {
            try
            {
                DataTable dTbl = new DataTable();
                string sCommText = "SELECT * FROM [OrderItems] WHERE ([AutoGenFileCreated] = 'E - Error Parsing'" +
                    " OR [AutoGenFileCreated] = 'E - Missing AutoGen File' OR [Rendered] = 'F' OR [XmlFileCreated] = 'E') AND [ErrorEmailSent] IS NULL";

                this.SQLQuery(sTimberlineConnString, sCommText, dTbl);

                if (dTbl.Rows.Count > 0)
                {
                    string sRefNum = Convert.ToString(dTbl.Rows[0]["RefNum"]).Trim();
                    string sPRodNum = Convert.ToString(dTbl.Rows[0]["ProdNum"]).Trim();
                    string sFrameNum = Convert.ToString(dTbl.Rows[0]["FrameNum"]).Trim();
                    string sOrderItemID = Convert.ToString(dTbl.Rows[0]["ItemID"]).Trim();
                    string sQuantity = Convert.ToString(dTbl.Rows[0]["Quantity"]).Trim();
                    string sCode = Convert.ToString(dTbl.Rows[0]["OrderItemCode"]).Trim();
                    string sPartNum = Convert.ToString(dTbl.Rows[0]["PartNumber"]).Trim();
                    string sImagePath = Convert.ToString(dTbl.Rows[0]["ImagePath"]).Trim();
                    string sUID = Convert.ToString(dTbl.Rows[0]["UID"]);

                    sMySubject = string.Format("Timberline Error Reporting");
                    sMyBody = string.Format("An error has occurred in the OrderItems table for UID = '" + sUID + "'");

                    sCommText = string.Empty;
                    string sDTNow = DateTime.Now.ToString("MM/dd/yy H:mm:ss");
                    sCommText = "UPDATE [OrderItems] SET [ErrorEmailSent] = 'T' WHERE [UID] = '" + sUID + "'";
                    this.SQLNonQuery(sTimberlineConnString, sCommText);
                }
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public void CheckForShippedXML(ref string sMySubject, ref string sMyBody) // Query the database for shipped orders and move files/dirs accordingly. 
        {
            try
            {
                DataTable dtGeneric = new DataTable();
                string sCommText = "SELECT * FROM [OrderItems] WHERE [XmlFileShippedPath] IS NOT NULL AND ([XmlOrderShippedReceived] IS NULL OR [XmlOrderShippedReceived] = 'F')";

                this.SQLQuery(sTimberlineConnString, sCommText, dtGeneric);

                if (dtGeneric.Rows.Count > 0)
                {
                    string sUID = Convert.ToString(dtGeneric.Rows[0]["UID"]).Trim();
                    string sConfNum = Convert.ToString(dtGeneric.Rows[0]["TCConfNum"]).Trim();
                    string sPONum = Convert.ToString(dtGeneric.Rows[0]["TCPONum"]).Trim();
                    string sShippedXmlPath = Convert.ToString(dtGeneric.Rows[0]["XmlFileShippedPath"]).Trim();
                    string sShippedXmlName = Path.GetFileName(sShippedXmlPath);
                    string sAcceptedXML = Convert.ToString(dtGeneric.Rows[0]["XmlFileAcceptedPath"]).Trim();
                    string sAcceptedPath = Path.GetDirectoryName(sAcceptedXML);
                    string sNewAcceptedPath = sAcceptedXML.Replace("Posted", "Shipped").Trim();
                    string sShippedPath = sAcceptedPath.Replace("Posted", "Shipped");
                    string sMovedXml = sAcceptedPath + @"\" + sShippedXmlName;
                    string sProdNum = Convert.ToString(dtGeneric.Rows[0]["ProdNum"]).Trim();
                    string sRefNum = Convert.ToString(dtGeneric.Rows[0]["RefNum"]).Trim();
                    string sFrameNum = Convert.ToString(dtGeneric.Rows[0]["FrameNum"]).Trim();
                    string sOrderItemID = Convert.ToString(dtGeneric.Rows[0]["ItemID"]).Trim();
                    int iQuantity = Convert.ToInt32(dtGeneric.Rows[0]["Quantity"]);
                    string sOrderItemCode = Convert.ToString(dtGeneric.Rows[0]["OrderItemCode"]).Trim();
                    DateTime dtPosted = Convert.ToDateTime(dtGeneric.Rows[0]["XmlFilePostedTime"]);
                    DateTime dtShipped = Convert.ToDateTime(dtGeneric.Rows[0]["XmlShippedDateTime"]);
                    string sLocalXMLFilePath = Convert.ToString(dtGeneric.Rows[0]["LocalXMLFilePath"]).Trim();
                    string sNewLocalXMLFilePath = sLocalXMLFilePath.Replace("Posted", "Shipped").Trim();
                    string sRenderedWebPath = Convert.ToString(dtGeneric.Rows[0]["RenderedWebPath"]).Trim();

                    XmlDocument xmlDoc = new XmlDocument();
                    xmlDoc.Load(sLocalXMLFilePath);
                    XmlNode root = xmlDoc.DocumentElement;
                    XmlNodeList xmlNL01 = root.SelectNodes("//Order/ShipToName");
                    XmlNodeList xmlNL02 = root.SelectNodes("//Order/ShipToContact");
                    XmlNodeList xmlNL03 = root.SelectNodes("//Order/AddressLine1");
                    XmlNodeList xmlNL04 = root.SelectNodes("//Order/City");
                    XmlNodeList xmlNL05 = root.SelectNodes("//Order/State");
                    XmlNodeList xmlNL06 = root.SelectNodes("//Order/Country");
                    XmlNodeList xmlNL07 = root.SelectNodes("//Order/Zip");
                    XmlNodeList xmlNL08 = root.SelectNodes("//Order/Phone");
                    XmlNodeList xmlNL09 = root.SelectNodes("//Order/ShipMethod");

                    string sShipToName = Convert.ToString(xmlNL01[0].InnerText).Trim();
                    string sShipToContact = Convert.ToString(xmlNL02[0].InnerText).Trim();
                    string sAddress1 = Convert.ToString(xmlNL03[0].InnerText).Trim();
                    string sCity = Convert.ToString(xmlNL04[0].InnerText).Trim();
                    string sState = Convert.ToString(xmlNL05[0].InnerText).Trim();
                    string sCountry = Convert.ToString(xmlNL06[0].InnerText).Trim();
                    string sZip = Convert.ToString(xmlNL07[0].InnerText).Trim();
                    string sPhone = Convert.ToString(xmlNL08[0].InnerText).Trim();
                    string sShipMethod = Convert.ToString(xmlNL09[0].InnerText).Trim();

                    if (File.Exists(sShippedXmlPath) && (!File.Exists(sMovedXml)))
                    {
                        File.Move(sShippedXmlPath, sMovedXml);
                    }
                    else if (!File.Exists(sShippedXmlPath) || (File.Exists(sMovedXml)))
                    {
                        //Note: Send an email here.
                        return;
                    }

                    if (Directory.Exists(sAcceptedPath))
                    {
                        Directory.Move(sAcceptedPath, sShippedPath);

                        sMovedXml = sMovedXml.Replace("Posted", "Shipped");

                        sCommText = "UPDATE [OrderItems] SET [XmlFileShippedPath] = '" + sMovedXml + "', [XmlOrderShippedReceived] = 'T', [LocalXMLFilePath] = '" + sNewLocalXMLFilePath + "'" +
                            ", [XmlFileAcceptedPath] = '" + sNewAcceptedPath + "' WHERE [UID] = '" + sUID + "'";
                        this.SQLNonQuery(sTimberlineConnString, sCommText);

                        // Send email indicating order has shipped.                      

                        StringBuilder sb = new StringBuilder();
                        sb.Clear();

                        sb.AppendFormat(@"Order information:");
                        sb.Append(Environment.NewLine);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Production Number: " + sProdNum);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Reference Number: " + sRefNum);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Frame Number: " + sFrameNum);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"DP2 Order Item: " + sOrderItemID);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"UID: " + sUID);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Quantity: " + iQuantity);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"DP2 Product ID: " + sOrderItemCode);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"POSTed date: " + dtPosted);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"TC Confirmation number: " + sConfNum);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"TC PO number: " + sPONum);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Shipped date: " + dtShipped);
                        sb.Append(Environment.NewLine);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Shipping information:");
                        sb.Append(Environment.NewLine);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Ship to name: " + sShipToName);
                        sb.Append(Environment.NewLine);
                        if (sShipToContact != string.Empty || sShipToContact != null)
                        {
                            sb.AppendFormat(@"Ship to contact: " + sShipToContact);
                            sb.Append(Environment.NewLine);
                        }
                        sb.AppendFormat(@"Address: " + sAddress1);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"City: " + sCity);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"State: " + sState);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Country: " + sCountry);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Zip: " + sZip);
                        sb.Append(Environment.NewLine);
                        if (sPhone != string.Empty || sPhone != null)
                        {
                            sb.AppendFormat(@"Phone: " + sPhone);
                            sb.Append(Environment.NewLine);
                        }
                        sb.AppendFormat(@"Ship method: " + sShipMethod);
                        sb.Append(Environment.NewLine);
                        sb.AppendFormat(@"Rendered product: " + sRenderedWebPath);
                        sb.Append(Environment.NewLine);

                        sMySubject = "Timberline order shipped notification received.";
                        sMyBody = sb.ToString();
                    }
                    else if (!Directory.Exists(sAcceptedPath))
                    {
                        //Note: Send an email here.
                        return;
                    }
                }
                else if (dtGeneric.Rows.Count == 0)
                {
                    return;
                }
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        public DataTable RemoveDuplicateRows(DataTable dt, string sColName)
        {
            Hashtable hTable = new Hashtable();
            ArrayList duplicateList = new ArrayList();

            //Add list of all the unique item value to hashtable, which stores combination of key, value pair.
            //And add duplicate item value in arraylist.
            foreach (DataRow drow in dt.Rows)
            {
                if (hTable.Contains(drow[sColName]))
                    duplicateList.Add(drow);
                else
                    hTable.Add(drow[sColName], string.Empty);
            }

            //Removing a list of duplicate items from datatable.
            foreach (DataRow dRow in duplicateList)
                dt.Rows.Remove(dRow);

            //Datatable which contains unique records will be return as output.
            return dt;
        }

        public void RemoveDuplicatesFromDataTable(ref DataTable dtFinal)
        {
            List<string> keyColumns = new List<string>();
            keyColumns.Add("Frame");
            keyColumns.Add("Code");
            Dictionary<string, string> uniquenessDict = new Dictionary<string, string>(dtFinal.Rows.Count);
            StringBuilder stringBuilder = null;
            int rowIndex = 0;
            DataRow row;
            DataRowCollection rows = dtFinal.Rows;

            while (rowIndex < rows.Count)
            {
                row = rows[rowIndex];

                stringBuilder = new StringBuilder();

                foreach (string colname in keyColumns)
                {
                    stringBuilder.Append(((string)row[colname]));
                }
                if (uniquenessDict.ContainsKey(stringBuilder.ToString()))
                {
                    rows.Remove(row);
                }
                else
                {
                    uniquenessDict.Add(stringBuilder.ToString(), string.Empty);
                    rowIndex++;
                }
            }
        }
    }
}

