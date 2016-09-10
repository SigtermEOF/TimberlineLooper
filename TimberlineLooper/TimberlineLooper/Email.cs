using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.Data.Odbc;
using System.Data.OleDb;
using System.Data.Sql;
using System.Data.SqlClient;
using System.Diagnostics;

namespace TimberlineLooper
{
    class Email
    {
        string sTimberlineConnString = TimberlineLooper.Properties.Settings.Default.TimberlineConnString.ToString();
        string sDP2ConnString = TimberlineLooper.Properties.Settings.Default.DP2ConnString.ToString();
        string sCDSConnString = TimberlineLooper.Properties.Settings.Default.CDSConnString.ToString();

        public void SendEmail(string sEmailServer, string sMyEmail, string sEmailBCC, string sMysubject, string sMybody)
        {
            MailAddress from = new MailAddress("APSAUTO@ADVANCEDPHOTO.COM", "APS");
            MailAddress to = new MailAddress(sMyEmail);
            MailMessage message = new MailMessage(from, to);
            message.Subject = sMysubject;
            message.Body = sMybody;
            MailAddress bcc = new MailAddress(sEmailBCC);
            message.Bcc.Add(bcc);
            SmtpClient myclient = new SmtpClient(sEmailServer);
            myclient.Credentials = CredentialCache.DefaultNetworkCredentials;

            try
            {
                myclient.Send(message);
            }
            catch (Exception ex)
            {
                this.SaveExceptionToDB(ex);
            }
        }

        private void SaveExceptionToDB(Exception ex)
        {
            string sException = ex.ToString().Trim();
            sException = sException.Replace(@"'", "");

            string sCommText = "INSERT INTO [Errors] VALUES ('" + sException + "', '" + DateTime.Now.ToString().Trim() + "', '0')";

            this.SQLNonQuery(sTimberlineConnString, sCommText);
        }

        private void SQLNonQuery(string sConnString, string sCommText)
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
    }
}
