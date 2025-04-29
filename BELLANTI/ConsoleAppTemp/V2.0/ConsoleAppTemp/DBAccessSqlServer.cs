using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleAppTemp
{
    internal class DBAccessSqlServer
    {
        private static OleDbConnection objConnection;
        public static string ConnectionString = ConfigurationManager.AppSettings["connectionStringSQLServerMIA"].ToString();



        public static void OpenConnection()
        {
            try
            {
                if (objConnection == null)
                {
                    objConnection = new OleDbConnection(ConnectionString);
                    objConnection.Open();
                }
                else
                {
                    if (objConnection.State != ConnectionState.Open)
                    {
                        objConnection = new OleDbConnection(ConnectionString);
                        objConnection.Open();
                    }
                }
            }
            catch (Exception ex)
            {
                new LogWriter(ex.ToString());
                // Send email notification
                //SendErrorEmail("MIA - Error during query execution SQL Server", ex.ToString());
            }
        }

        public static void CloseConnection()
        {
            try
            {
                if (!(objConnection == null))
                {
                    if (objConnection.State == System.Data.ConnectionState.Open)
                    {
                        objConnection.Close();
                        objConnection.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                new LogWriter(ex.ToString());
                // Send email notification
                //SendErrorEmail("MIA - Error during query execution SQL Server", ex.ToString());
            }
        }

        public static bool ExecuteQuery(string query)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                {
                    using (OleDbCommand cmd = new OleDbCommand(query, connection))
                    {
                        cmd.CommandTimeout = 600;
                        connection.Open();
                        int rowsAffected = cmd.ExecuteNonQuery();
                        connection.Close();
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                new LogWriter("Errore durante esecuzione query: " + query + "\n" + ex.ToString());
                // Send email notification
                //SendErrorEmail("MIA - Error during query execution SQL Server", ex.ToString());
                return false;
            }
        }

        public static bool ReadDataThroughAdapter(string query, DataTable tblName)
        {
            try
            {
                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                {
                    using (OleDbCommand cmd = new OleDbCommand(query, connection))
                    {
                        if (connection.State == ConnectionState.Closed)
                        {
                            connection.Open();
                        }

                        cmd.Connection = connection;
                        cmd.CommandText = query;
                        cmd.CommandTimeout = 600;
                        cmd.CommandType = CommandType.Text;

                        OleDbDataAdapter adapter = new OleDbDataAdapter(cmd);
                        adapter.Fill(tblName);
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                new LogWriter("Eccezione durante esecuzione query: " + query + ". \n" + ex.Message.ToString() + "\n" + ex.ToString());
                // Send email notification
                //SendErrorEmail("MIA - Error during query execution SQL Server", ex.ToString());
                return false;
            }
        }

        public static OleDbDataReader ReadDataThroughReader(string query)
        {
            //DataReader used to sequentially read data from a data source
            OleDbDataReader reader;

            try
            {
                using (OleDbConnection connection = new OleDbConnection(ConnectionString))
                {
                    using (OleDbCommand cmd = new OleDbCommand(query, connection))
                    {
                        if (connection.State == ConnectionState.Closed)
                        {
                            connection.Open();
                        }

                        cmd.Connection = connection;
                        cmd.CommandText = query;
                        cmd.CommandTimeout = 600;
                        cmd.CommandType = CommandType.Text;

                        reader = cmd.ExecuteReader();
                        return reader;
                    }
                }
            }
            catch (Exception ex)
            {
                new LogWriter("Eccezione durante esecuzione query: " + query + ". \n" + ex.Message.ToString() + "\n" + ex.ToString());
                // Send email notification
                //SendErrorEmail("MIA - Error during query execution SQL Server", ex.ToString());
                throw;
            }
        }

        //public static void SendErrorEmail(string subject, string errorMessage)
        //{
        //    try
        //    {
        //        MailMessage msg = new MailMessage();
        //        msg.To.Add(ConfigurationManager.AppSettings["destinatari"].ToString());
        //        string pwd = ConfigurationManager.AppSettings["pwd_smtp"];
        //        msg.From = new MailAddress(ConfigurationManager.AppSettings["mittente"].ToString());
        //        msg.Subject = subject;
        //        msg.Priority = MailPriority.Normal;
        //        msg.IsBodyHtml = true;
        //        msg.Body = errorMessage;

        //        SmtpClient smtpClient = new SmtpClient();
        //        smtpClient.UseDefaultCredentials = false;
        //        smtpClient.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["user_smtp"].ToString(), pwd);
        //        smtpClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
        //        smtpClient.Host = ConfigurationManager.AppSettings["smtp"];
        //        smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

        //        smtpClient.Send(msg);

        //        smtpClient.Dispose();
        //        msg.Dispose();
        //    }
        //    catch (Exception ex)
        //    {
        //        // If sending email fails, log the error
        //        new LogWriter("Error sending error email: \n" + ex.ToString());
        //    }
        //}
    }
}


