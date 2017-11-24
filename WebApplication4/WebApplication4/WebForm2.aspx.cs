using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace WebApplication4
{
    public partial class WebForm2 : System.Web.UI.Page
    {

        protected void Page_Load(object sender, EventArgs e)
        {
            DataSet ds = GetData();
            string htmlString = getHtml(ds);

            SendMail(htmlString);
        }

        public DataSet GetData()
        {
            SqlConnection sqlConnection1 = new SqlConnection(@"Data Source = rnrtestserver.database.windows.net; Initial Catalog = AMD_PROD;Connection Timeout=9000;
            User ID = rnrAdmin; Password = India@123");
            SqlCommand cmd = new SqlCommand();
            DataTable dt = new DataTable();
            cmd.CommandText = "show_msg_forothers";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@StartDate", "2017-11-20");
            cmd.Parameters.AddWithValue("@EndDate", "2017-11-24");
           // cmd.Parameters.AddWithValue("@ViewName", "Semantic");
            cmd.Connection = sqlConnection1;
            sqlConnection1.Open();
            cmd.CommandTimeout = 600;
            GridView2.DataSource = cmd.ExecuteReader();
            GridView2.DataBind();
            // Data is accessible through the DataReader object here.

            sqlConnection1.Close();
            dt.Columns.Add("date", typeof(string));
            dt.Columns.Add("Viewname", typeof(string));
            dt.Columns.Add("ViewOwnerName", typeof(string));
            foreach (GridViewRow row in GridView2.Rows)
            {
                string date = Convert.ToString(row.Cells[0].Text);
                string viewname = row.Cells[1].Text;
                string ViewOwnerName = row.Cells[2].Text;
                dt.Rows.Add(date, viewname,ViewOwnerName);
                foreach (DataRow rows in dt.Rows)
                {
                    foreach (DataColumn column in dt.Columns)
                    {
                        //Console.WriteLine(rows[column]);
                    }
                }
            }

            DataSet ds = new DataSet();
            ds.Tables.Add(dt);
            return ds;
        }

        public void SendMail(string htmlString)
        {
            try
            {
                string fromMail = "v-robelw@microsoft.com";
                string fromPassword = "nem!@#$%12345";
                string[] ToMailIds = new string[] { "rnradhoc@microsoft.com" };
                //string[] ToMailIds = new string[] { "v-robelw@microsoft.com", "v-raman@microsoft.com" };


                NetworkCredential cred = new NetworkCredential(fromMail, fromPassword);
                MailMessage mail = new MailMessage();
                mail.From = new MailAddress(fromMail, "Adhoc Support Team");
                foreach (string mailId in ToMailIds)
                {
                    mail.To.Add(new MailAddress(mailId));
                }

                SmtpClient client = new SmtpClient
                {
                    Port = 587,
                    EnableSsl = true,
                    DeliveryMethod = SmtpDeliveryMethod.Network,
                    Credentials = cred
                };

                mail.IsBodyHtml = true;
                client.Host = "outlook.office365.com";
                mail.Subject = string.Format("Test mail", "notification");
                StringBuilder sb = new StringBuilder();
                sb.Append("Its a test mail for notification.");

                mail.Body = htmlString; //dt.ToString(); // sb.ToString();
                client.Send(mail);
            }
            catch (Exception ex)
            {

            }
        }

        public static string getHtml(DataSet dataSet)
        {
            try
            {
                string messageBody = "<font>The following dates have missing data for the corresponding views </font><br><br>";

                if (dataSet.Tables[0].Rows.Count == 0)
                    return messageBody;
                string htmlTableStart = "<table style=\"border-collapse:collapse; text-align:center;\" >";
                string htmlTableEnd = "</table>";
                string htmlHeaderRowStart = "<tr style =\"background-color:#6FA1D2; color:#ffffff;\">";
                string htmlHeaderRowEnd = "</tr>";
                string htmlTrStart = "<tr style =\"color:#555555;\">";
                string htmlTrEnd = "</tr>";
                string htmlTdStart = "<td style=\" border-color:#5c87b2; border-style:solid; border-width:thin; padding: 5px;\">";
                string htmlTdEnd = "</td>";

                messageBody += htmlTableStart;
                messageBody += htmlHeaderRowStart;
                messageBody += htmlTdStart + "date " + htmlTdEnd;
                messageBody += htmlTdStart + "viewName " + htmlTdEnd;
                messageBody += htmlTdStart + "ViewOwnerName " + htmlTdEnd;

                messageBody += htmlHeaderRowEnd;

                foreach (DataRow Row in dataSet.Tables[0].Rows)
                {
                    messageBody = messageBody + htmlTrStart;
                    messageBody = messageBody + htmlTdStart + Row["date"] + htmlTdEnd;
                    messageBody = messageBody + htmlTdStart + Row["viewname"] + htmlTdEnd;
                    messageBody = messageBody + htmlTdStart + Row["ViewOwnerName"] + htmlTdEnd;


                    messageBody = messageBody + htmlTrEnd;
                }
                messageBody = messageBody + htmlTableEnd;
                return messageBody;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}