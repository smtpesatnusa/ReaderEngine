using MySql.Data.MySqlClient;
using System;
using System.Data;

namespace ReaderEngine
{
    public class ConnectionDB
    {
        MySqlConnection conn;

        static string host = "192.168.192.150";
        static string database = "smt_attendance";
        static string userDB = "smt_developer";
        static string password = "w(v97weP8UGe=bYd";

        public static string strProvider = "server=" + host + ";Database=" + database + ";User ID=" + userDB + ";Password=" + password+ ";SslMode=None";
        public MySqlConnection connection = new MySqlConnection("server=" + host + ";Database=" + database + ";User ID=" + userDB + ";Password=" + password + ";SslMode=None");

        public bool Open()
        {
            try
            {
                strProvider = "server=" + host + ";Database=" + database + ";User ID=" + userDB + ";Password=" + password + ";SslMode=None";
                conn = new MySqlConnection(strProvider);
                conn.Open();
                return true;
            }
            catch (Exception er)
            {
                //MessageBox.Show("Connection Error ! " + er.Message, "Information");
            }
            return false;
        }
        public void Close()
        {
            try
            {
                conn.Close();
                conn.Dispose();
            }
            catch
            {

            }
           
        }

        public DataTable GetDataTable(string sq)
        {
            DataTable dt = new DataTable();
            using(MySqlDataAdapter da = new MySqlDataAdapter(sq, conn))
            {
                da.Fill(dt);
            }

            return dt;
        }

        public DataSet ExecuteDataSet(string sql)
        {
            try
            {
                DataSet ds = new DataSet();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                da.Fill(ds, "result");
                return ds;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return null;
        }
        public MySqlDataReader ExecuteReader(string sql)
        {
            try
            {
                MySqlDataReader reader;
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                reader = cmd.ExecuteReader();
                return reader;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return null;
        }
        public int ExecuteNonQuery(string sql)
        {
            try
            {
                int affected;
                MySqlTransaction mytransaction = conn.BeginTransaction();
                MySqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                affected = cmd.ExecuteNonQuery();
                mytransaction.Commit();
                return affected;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return -1;
        }
    }

    public class ConnectionDBESD
    {
        MySqlConnection conn;

        static string host = "198.168.9.1";
        static string database = "idaesd";
        static string userDB = "ida";
        static string password = "ida6422690";

        public static string strProvider = "server=" + host + ";Database=" + database + ";User ID=" + userDB + ";Password=" + password + ";SslMode=None;Connection Timeout=30";
        public MySqlConnection connection = new MySqlConnection("server=" + host + ";Database=" + database + ";User ID=" + userDB + ";Password=" + password + ";SslMode=None;Connection Timeout=30");

        public bool Open()
        {
            try
            {
                strProvider = "server=" + host + ";Database=" + database + ";User ID=" + userDB + ";Password=" + password + ";SslMode=None";
                conn = new MySqlConnection(strProvider);
                conn.Open();
                return true;
            }
            catch (Exception er)
            {
                //MessageBox.Show("Connection Error ! " + er.Message, "Information");
            }
            return false;
        }
        public void Close()
        {
            try
            {
                conn.Close();
                conn.Dispose();
            }
            catch
            {

            }

        }

        public DataTable GetDataTable(string sq)
        {
            DataTable dt = new DataTable();
            using (MySqlDataAdapter da = new MySqlDataAdapter(sq, conn))
            {
                da.Fill(dt);
            }

            return dt;
        }

        public DataSet ExecuteDataSet(string sql)
        {
            try
            {
                DataSet ds = new DataSet();
                MySqlDataAdapter da = new MySqlDataAdapter(sql, conn);
                da.Fill(ds, "result");
                return ds;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return null;
        }
        public MySqlDataReader ExecuteReader(string sql)
        {
            try
            {
                MySqlDataReader reader;
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                reader = cmd.ExecuteReader();
                return reader;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return null;
        }
        public int ExecuteNonQuery(string sql)
        {
            try
            {
                int affected;
                MySqlTransaction mytransaction = conn.BeginTransaction();
                MySqlCommand cmd = conn.CreateCommand();
                cmd.CommandText = sql;
                affected = cmd.ExecuteNonQuery();
                mytransaction.Commit();
                return affected;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return -1;
        }
    }

}
