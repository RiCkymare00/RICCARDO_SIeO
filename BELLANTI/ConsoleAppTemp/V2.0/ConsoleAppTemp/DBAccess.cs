using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data;
using System.Linq;
using System.Net.Mail;
using System.Reflection.Metadata;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using iAnywhere.Data.SQLAnywhere;

namespace ConsoleAppTemp
{
    public class DBAccess
    {
        private static readonly SAConnection connection = new();
        private static readonly SACommand command = new();
        private static SADataAdapter adapter = new();
        public SATransaction DbTran;

        private static readonly string strConnString = ConfigurationManager.AppSettings["connectionStringSybaseBellanti"].ToString();

        public static void CreateConn()
        {
            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.ConnectionString = strConnString;
                    connection.Open();
                }
            }
            catch (Exception ex)
            {
                _ = new LogWriter("Eccezione durante apertura della connessione. " + ex.Message);
                // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                throw;
            }
        }


        public void CloseConn()
        {
            connection.Close();
        }


        public static int ExecuteDataAdapter(DataTable tblName, string strSelect)
        {
            try
            {
                if (connection.State == 0)
                {
                    CreateConn();
                }

                adapter.SelectCommand.CommandText = strSelect;
                adapter.SelectCommand.CommandType = CommandType.Text;
                SACommandBuilder DbCommandBuilder = new(adapter);

                string insert = DbCommandBuilder.GetInsertCommand().CommandText.ToString();
                string update = DbCommandBuilder.GetUpdateCommand().CommandText.ToString();
                string delete = DbCommandBuilder.GetDeleteCommand().CommandText.ToString();

                return adapter.Update(tblName);
            }
            catch (Exception ex)
            {
                _ = new LogWriter("Eccezione durante esecuzione query: " + strSelect + ". " + ex.Message);
                // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                throw;
            }
        }


        public static void ReadDataThroughAdapter(string query, DataTable tblName, SACommand cmd = null)
        {
            try
            {
                if (connection.State == ConnectionState.Closed)
                {
                    CreateConn();
                }

                command.Connection = connection;
                command.CommandText = query;
                command.CommandType = CommandType.Text;

                adapter = new SADataAdapter(command);
                _ = adapter.Fill(tblName);
            }
            catch (Exception ex)
            {
                if (cmd != null)
                {
                    _ = new LogWriter("Eccezione durante esecuzione query: " + cmd.CommandText + "\n" + ex.Message.ToString());
                    // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                }
                else
                {
                    _ = new LogWriter("Eccezione durante esecuzione query: " + command.CommandText + "\n" + ex.Message.ToString());
                    // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                }
                throw;
            }
        }


        public static SADataReader ReadDataThroughReader(string query)
        {
            //DataReader used to sequentially read data from a data source
            SADataReader reader;

            try
            {
                if (connection.State == ConnectionState.Closed)
                {
                    CreateConn();
                }

                command.Connection = connection;
                command.CommandText = query;
                command.CommandType = CommandType.Text;

                reader = command.ExecuteReader();
                return reader;
            }
            catch (Exception ex)
            {
                _ = new LogWriter("Eccezione durante esecuzione query: " + command.CommandText + ". " + ex.Message);
                // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                throw;
            }
        }

        public static bool ExecuteQuery(string query)
        {
            try
            {
                using (SAConnection connection = new(strConnString))
                {
                    using (SACommand cmd = new(query, connection))
                    {
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
                // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                return false;
            }
        }


        public static int ExecuteQuery(SACommand dbCommand)
        {
            try
            {
                if (connection.State == 0)
                {
                    CreateConn();
                }

                dbCommand.Connection = connection;
                dbCommand.CommandType = CommandType.Text;


                return dbCommand.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                _ = new LogWriter("Eccezione durante esecuzione query: " + dbCommand.CommandText + ". " + ex.Message);
                // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                return 0;
            }
        }

        public static object ExecuteQueryWithScalar(string query)
        {
            try
            {
                using (SAConnection connection = new SAConnection(strConnString))
                {
                    using (SACommand cmd = new SACommand(query, connection))
                    {
                        connection.Open();
                        object result = cmd.ExecuteScalar();
                        connection.Close();
                        return result;
                    }
                }
            }
            catch (Exception ex)
            {
                new LogWriter("Errore durante esecuzione query: " + query + "\n" + ex.ToString());
                // SendErrorEmail("MIA - Error during query execution SQL Sybase bellanti", ex.ToString());
                return null; // Or throw an exception if you prefer
            }
        }

        public static void SendErrorEmail(string subject, string errorMessage)
        {
            try
            {
                MailMessage msg = new MailMessage();
                msg.To.Add(ConfigurationManager.AppSettings["destinatari"].ToString());
                string pwd = ConfigurationManager.AppSettings["pwd_smtp"];
                msg.From = new MailAddress(ConfigurationManager.AppSettings["mittente"].ToString());
                msg.Subject = subject;
                msg.Priority = MailPriority.Normal;
                msg.IsBodyHtml = true;
                msg.Body = errorMessage;

                SmtpClient smtpClient = new SmtpClient();
                smtpClient.UseDefaultCredentials = false;
                smtpClient.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["user_smtp"].ToString(), pwd);
                smtpClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
                smtpClient.Host = ConfigurationManager.AppSettings["smtp"];
                smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;

                smtpClient.Send(msg);

                smtpClient.Dispose();
                msg.Dispose();
            }
            catch (Exception ex)
            {
                // If sending email fails, log the error
                new LogWriter("Error sending error email: \n" + ex.ToString());
            }
        }

    }


}
