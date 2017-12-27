using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Data;
using System.Data.OracleClient;
using System.Data.Odbc;
using System.Windows.Forms;

namespace LG_ZGQ
{
    class DBDataZGQ
    {
        /// <summary>
        /// SQL
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="cmdText"></param>
        /// <param name="Exceptionmessage"></param>
        /// <returns></returns>
        public int SQL_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            int x = 0;
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand com = null;
            try
            {
                com = new SqlCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch(Exception  ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int SQL_ExecuteNonQuery(string connectionString, string cmdText, SqlParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand com = null;
            try
            {
                com = new SqlCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (SqlParameter OdbcParameter in Parameters)
                        com.Parameters.Add(OdbcParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int SQL_ExecuteNonQuery(string connectionString, string cmdText, SqlParameter[] Parameters,CommandType comType, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con = new SqlConnection(connectionString);
            SqlCommand com = null;
            try
            {
                com = new SqlCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;

                if (Parameters.Length > 0)
                {                    
                    foreach (SqlParameter SqlParameter in Parameters)
                        com.Parameters.Add(SqlParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
               com.Dispose();
     
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public SqlDataReader SQL_ExecuteReader(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con  = new SqlConnection(connectionString);
            SqlCommand com = null;
            SqlDataReader sdr =null;
            try
            {
                com = new SqlCommand(cmdText, con);
                con.Open();
               sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch (Exception ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                com.Dispose();
                return sdr;
            }
        }
        public DataTable SQL_DataAdapter(string  connectionString,string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                return dt;
            }

        }
        public DataTable SQL_DataAdapter(string connectionString, string cmdText, SqlParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                SqlCommand com = new SqlCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (SqlParameter OdbcParameter in Parameters)
                        com.Parameters.Add(OdbcParameter);
                }
                SqlDataAdapter da = new SqlDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                return dt;
            }

        }
        public DataSet SQL_DataAdapter_DataSet(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection(connectionString);
            DataSet ds = new DataSet();
            try
            {
                SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch (Exception ee)
            {
                Exceptionmessage = "异常：" + ee.Message;
                con.Close();
                return ds;
            }

        }
        /// <summary>
        /// ORACLE
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="cmdText"></param>
        /// <returns></returns>
        public int Oracle_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OracleConnection con = new OracleConnection(connectionString);
            OracleCommand com = null;
            try
            {
                com = new OracleCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int Oracle_ExecuteNonQuery(string connectionString, string cmdText, OracleParameter[] Parameters, ref string Exceptionmessage)
        {
         
            int x = 0;
            OracleConnection con = new OracleConnection(connectionString);
            OracleCommand com = null;
            try
            {
                com = new OracleCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OracleParameter OdbcParameter in Parameters)
                        com.Parameters.Add(OdbcParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int Oracle_ExecuteNonQuery(string connectionString, string cmdText, OracleParameter[] Parameters, CommandType comType, ref string Exceptionmessage)
        {
           
            int x = 0;
            OracleConnection con = new OracleConnection(connectionString);
            OracleCommand com = null;
            try
            {
                com = new OracleCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;

                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public OracleDataReader Oracle_ExecuteReader(string connectionString, string cmdText)
        {
           
            int x = 0;
            OracleConnection con = new OracleConnection(connectionString);
            OracleCommand com = null;
            OracleDataReader sdr = null;
            try
            {
                com = new OracleCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return sdr;
            }
        }
        public DataTable Oracle_DataAdapter(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            
            OracleConnection con = new OracleConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                OracleDataAdapter da = new OracleDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch
            {
                con.Close();
                return dt;
            }

        }
        public DataTable Oracle_DataAdapter(string connectionString, string cmdText, OracleParameter[] Parameters)
        {
          
            OracleConnection con = new OracleConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                OracleCommand com = new OracleCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                OracleDataAdapter da = new OracleDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch
            {
                con.Close();
                return dt;
            }

        }
        public DataSet Oracle_DataAdapter_DataSet(string connectionString, string cmdText)
        {
       
            OracleConnection con = new OracleConnection(connectionString);
            DataSet ds = new DataSet();
            try
            {
                OracleDataAdapter da = new OracleDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close(); 
                da.Dispose();
                return ds;
            }
            catch
            {
                con.Close();
                return ds;
            }

        }

        /// <summary>
        /// OLEDB
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="cmdText"></param>
        /// <returns></returns>
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText)
        {
           
            int x = 0;
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand com = null;
            try
            {
                com = new OleDbCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText,ref  OleDbParameter[] Parameters)
        {
           
            int x = 0;
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand com = null;
            try
            {
                com = new OleDbCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText,ref  OleDbParameter[] Parameters, CommandType comType)
        {
           
            int x = 0;
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand com = null;
            try
            {
                com = new OleDbCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;

                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public OleDbDataReader OleDb_ExecuteReader(string connectionString, string cmdText)
        {
           
            int x = 0;
            OleDbConnection con = new OleDbConnection(connectionString);
            OleDbCommand com = null;
            OleDbDataReader sdr = null;
            try
            {
                com = new OleDbCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return sdr;
            }
        }
        public DataTable OleDb_DataAdapter(string connectionString, string cmdText)
        {
           
            OleDbConnection con = new OleDbConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch
            {
                con.Close();
                return dt;
            }

        }
        public DataTable OleDb_DataAdapter(string connectionString, string cmdText,ref  OleDbParameter[] Parameters)
        {
          
            OleDbConnection con = new OleDbConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                OleDbCommand com = new OleDbCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                OleDbDataAdapter da = new OleDbDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close(); 
                da.Dispose();
                return dt;
            }
            catch
            {
                con.Close();
                return dt;
            }

        }
        public DataSet OleDb_DataAdapter_DataSet(string connectionString, string cmdText)
        {
        
            OleDbConnection con = new OleDbConnection(connectionString);
            DataSet ds = new DataSet();
            try
            {
                OleDbDataAdapter da = new OleDbDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch
            {
                con.Close();
                return ds;
            }

        }
        
        /// <summary>
        /// ODB C
        /// </summary>
        /// <param name="connectionString"></param>
        /// <param name="cmdText"></param>
        /// <returns></returns>
        public int Odbc_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection(connectionString);
            OdbcCommand com = null;
            try
            {
                com = new OdbcCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch(Exception  e)
            {
                con.Close();
                com.Dispose();
                Exceptionmessage = e.ToString();
                return -1;
            }
        }
        public int Odbc_ExecuteNonQuery(string connectionString, string cmdText,ref  OdbcParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection(connectionString);
            OdbcCommand com = null;
            try
            {
                com = new OdbcCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public int Odbc_ExecuteNonQuery(string connectionString, string cmdText,ref  OdbcParameter[] Parameters, CommandType comType, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection(connectionString);
            OdbcCommand com = null;
            try
            {
                com = new OdbcCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;

                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return -1;
            }
        }
        public OdbcDataReader Odbc_ExecuteReader(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection(connectionString);
            OdbcCommand com = null;
            OdbcDataReader sdr = null;
            try
            {
                com = new OdbcCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch
            {
                con.Close();
                com.Dispose();
                return sdr;
            }
        }
        public DataTable Odbc_DataAdapter(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                OdbcDataAdapter da = new OdbcDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch
            {
                con.Close();
                return dt;
            }

        }
        public DataTable Odbc_DataAdapter(string connectionString, string cmdText,ref  OdbcParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection(connectionString);
            DataTable dt = new DataTable();
            try
            {
                OdbcCommand com = new OdbcCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                    com.Parameters.AddRange(Parameters);
                OdbcDataAdapter da = new OdbcDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch
            {
                con.Close();
                return dt;
            }

        }
        public DataSet Odbc_DataAdapter_DataSet(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection(connectionString);
            DataSet ds = new DataSet();
            try
            {
                OdbcDataAdapter da = new OdbcDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch
            {
                con.Close();
                return ds;
            }

        }
    }

    class SqlDB_ZGQ
    {
        //Data Source=172.18.100.20;Initial Catalog=tj_db;User Id=sa;Password=sql;
        /// <summary>
        /// 判断SQL数据库
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public bool Sql_isOpenConnection(string connectionString,ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection();
            try
            {
                con.ConnectionString = connectionString;
                con.Open();
                con.Close();
                return  true;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return false;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行SQL的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int Sql_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con = new SqlConnection();
            SqlCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new SqlCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage =ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if(con.State==ConnectionState.Open)
                con.Close();
            }
        }

        /// <summary>
        ///  执行SQL的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Parameters">SqlParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns> 
        public int Sql_ExecuteNonQuery(string connectionString, string cmdText,ref  SqlParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con = new SqlConnection();
            SqlCommand com = null;
            try
            {
                con.ConnectionString = connectionString;

                com = new SqlCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (SqlParameter SqlParameter in Parameters)
                        com.Parameters.Add(SqlParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        ///  执行SQL的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="cmdText">SqlParameter[]参数</param>
        /// <param name="cmdText">CommandType</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int Sql_ExecuteNonQuery(string connectionString, string cmdText, ref SqlParameter[] Parameters, CommandType comType, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con = new SqlConnection();
            SqlCommand com = null;
            try
            {
                try
                {
                    con.ConnectionString = connectionString;
                }
                catch
                {
                    Exceptionmessage = "连接字符串异常";
                    return -1;
                }
                com = new SqlCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                  
                    foreach (SqlParameter SqlParameter in Parameters)
                        com.Parameters.Add(SqlParameter);
                }
               
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();

                if (x == -1)
                    x = 0;
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行 SQL的ExecuteReader()，返回SqlDataReader
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public SqlDataReader Sql_ExecuteReader(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            SqlConnection con = new SqlConnection();
            SqlCommand com = null;
            SqlDataReader sdr = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new SqlCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return sdr;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// SqlDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable Sql_DataAdapter(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage =ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// SqlDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        ///  <param name="cmdText">SqlParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable Sql_DataAdapter(string connectionString, string cmdText,ref SqlParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                SqlCommand com = new SqlCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (SqlParameter SqlParameter in Parameters)
                        com.Parameters.Add(SqlParameter);
                }
                SqlDataAdapter da = new SqlDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// SqlDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataSet Sql_DataAdapter_DataSet(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            SqlConnection con = new SqlConnection();
            DataSet ds = new DataSet();
            try
            {
                con.ConnectionString = connectionString;
                SqlDataAdapter da = new SqlDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return ds;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }
      
    }

    class OdbcDB_ZGQ
    {

        /// <summary>
        /// 判断Odbc数据库
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public bool Odbc_isOpenConnection(string connectionString, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection();
            try
            {
                con.ConnectionString = connectionString;
                con.Open();
                con.Close();
                return true;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return false;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行Odbc的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int Odbc_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection();
            OdbcCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OdbcCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        ///  执行Odbc的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Parameters">OdbcParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns> 
        public int Odbc_ExecuteNonQuery(string connectionString, string cmdText,ref OdbcParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection();
            OdbcCommand com = null;
            try
            {
                con.ConnectionString = connectionString;

                com = new OdbcCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OdbcParameter OdbcParameter in Parameters)
                        com.Parameters.Add(OdbcParameter);
                }

                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        ///  执行Odbc的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="cmdText">OdbcParameter[]参数</param>
        /// <param name="cmdText">CommandType</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int Odbc_ExecuteNonQuery(string connectionString, string cmdText,ref OdbcParameter[] Parameters, CommandType comType, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection();
            OdbcCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OdbcCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;

                if (Parameters.Length > 0)
                {
                    foreach (OdbcParameter OdbcParameter in Parameters)
                    com.Parameters.Add(OdbcParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();

                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行 Odbc的ExecuteReader()，返回OdbcDataReader
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public OdbcDataReader Odbc_ExecuteReader(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OdbcConnection con = new OdbcConnection();
            OdbcCommand com = null;
            OdbcDataReader sdr = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OdbcCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return sdr;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// OdbcDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable Odbc_DataAdapter(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                OdbcDataAdapter da = new OdbcDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// OdbcDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        ///  <param name="cmdText">OdbcParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable Odbc_DataAdapter(string connectionString, string cmdText,ref OdbcParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                OdbcCommand com = new OdbcCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OdbcParameter OdbcParameter in Parameters)
                        com.Parameters.Add(OdbcParameter);
                  
                }
                OdbcDataAdapter da = new OdbcDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// OdbcDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataSet Odbc_DataAdapter_DataSet(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OdbcConnection con = new OdbcConnection();
            DataSet ds = new DataSet();
            try
            {
                con.ConnectionString = connectionString;
                OdbcDataAdapter da = new OdbcDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return ds;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

    }

    class OleDbDB_ZGQ
    {
       // string ConnectionString = "Provider=MSDAORA;" + "Data Source=lisdb ;" + "User id=blinfo;" + "Password=blinfo;";
        /// <summary>
        /// 判断OleDb数据库
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public bool OleDb_isOpenConnection(string connectionString, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OleDbConnection con = new OleDbConnection();
            try
            {
                con.ConnectionString = connectionString;
                con.Open();
                con.Close();
                return true;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return false;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行OleDb的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OleDbConnection con = new OleDbConnection();
            OleDbCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                
                com = new OleDbCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }


        /// <summary>
        ///  执行OleDb的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="cmdText">OleDbParameter[]参数</param>
        /// <param name="cmdText">CommandType</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText, CommandType comType, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OleDbConnection con = new OleDbConnection();
            OleDbCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OleDbCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();

                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }
        /// <summary>
        ///  执行OleDb的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Parameters">OleDbParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns> 
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText,ref OleDbParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OleDbConnection con = new OleDbConnection();
            OleDbCommand com = null;
            try
            {
                con.ConnectionString = connectionString;

                com = new OleDbCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OleDbParameter OledbParameter in Parameters)
                        com.Parameters.Add(OledbParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        ///  执行OleDb的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="cmdText">OleDbParameter[]参数</param>
        /// <param name="cmdText">CommandType</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int OleDb_ExecuteNonQuery(string connectionString, string cmdText,ref OleDbParameter[] Parameters, CommandType comType, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OleDbConnection con = new OleDbConnection();
            OleDbCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OleDbCommand();
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;

                if (Parameters.Length > 0)
                {
                    foreach (OleDbParameter OledbParameter in Parameters)
                        com.Parameters.Add(OledbParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();

                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行 OleDb的ExecuteReader()，返回OleDbDataReader
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public OleDbDataReader OleDb_ExecuteReader(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OleDbConnection con = new OleDbConnection();
            OleDbCommand com = null;
            OleDbDataReader sdr = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OleDbCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return sdr;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// OleDbDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable OleDb_DataAdapter(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OleDbConnection con = new OleDbConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// OleDbDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        ///  <param name="cmdText">OleDbParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable OleDb_DataAdapter(string connectionString, string cmdText,ref  OleDbParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OleDbConnection con = new OleDbConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                OleDbCommand com = new OleDbCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OleDbParameter OledbParameter in Parameters)
                        com.Parameters.Add(OledbParameter);
                }
                OleDbDataAdapter da = new OleDbDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// OleDbDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataSet OleDb_DataAdapter_DataSet(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OleDbConnection con = new OleDbConnection();
            DataSet ds = new DataSet();
            try
            {
                con.ConnectionString = connectionString;
                OleDbDataAdapter da = new OleDbDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return ds;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

    }

    class OracleDB_ZGQ
    {
        //Data Source=ORCL_100;User ID=jk_ycxd;Password=123;
        /// <summary>
        /// 判断Oracle数据库
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public bool Oracle_isOpenConnection(string connectionString, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OracleConnection con = new OracleConnection();
            try
            {
                con.ConnectionString = connectionString;
                con.Open();
                con.Close();
                return true;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return false;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行Oracle的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int Oracle_ExecuteNonQuery(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OracleConnection con = new OracleConnection();
            OracleCommand com = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OracleCommand(cmdText, con);
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        ///  执行Oracle的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Parameters">OracleParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns> 
        public int Oracle_ExecuteNonQuery(string connectionString, string cmdText,ref OracleParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OracleConnection con = new OracleConnection();
            OracleCommand com = null;
            try
            {
                con.ConnectionString = connectionString;

                com = new OracleCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OracleParameter OracleParameter in Parameters)
                        com.Parameters.Add(OracleParameter);
                }
                con.Open();
                x = com.ExecuteNonQuery();
                con.Close();
                com.Dispose();
                return x;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        ///  执行Oracle的ExecuteNonQuery()，返回影响行数
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="cmdText">OracleParameter[]参数</param>
        /// <param name="cmdText">CommandType</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public int Oracle_ExecuteNonQuery(string connectionString, string cmdText,ref OracleParameter[] Parameters, CommandType comType, ref string Exceptionmessage)
        {

            MessageBox.Show("1");
            Exceptionmessage = "";
            int x = 0;
            OracleConnection con = new OracleConnection();
            OracleCommand com = null;
            MessageBox.Show("2");
            try
            {
                con.ConnectionString = connectionString;
                com = new OracleCommand();
                MessageBox.Show("3");
                com.Connection = con;
                com.CommandType = comType;
                com.CommandText = cmdText;
                MessageBox.Show("4");
                if (Parameters.Length > 0)
                {
                    foreach (OracleParameter OracleParameter in Parameters)
                        com.Parameters.Add(OracleParameter);
                }
                MessageBox.Show("5");
                con.Open();
               
                x = com.ExecuteNonQuery();
               
                con.Close();
                com.Dispose();
                MessageBox.Show("6");
                return x;
            }
            catch (Exception ee)
            {
                MessageBox.Show("8");
                Exceptionmessage = ee.Message;
                con.Close();
                com.Dispose();
                return -1;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// 执行 Oracle的ExecuteReader()，返回OracleDataReader
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public OracleDataReader Oracle_ExecuteReader(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            int x = 0;
            OracleConnection con = new OracleConnection();
            OracleCommand com = null;
            OracleDataReader sdr = null;
            try
            {
                con.ConnectionString = connectionString;
                com = new OracleCommand(cmdText, con);
                con.Open();
                sdr = com.ExecuteReader();
                con.Close();
                com.Dispose();
                return sdr;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return sdr;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }
        }

        /// <summary>
        /// OracleDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable Oracle_DataAdapter(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OracleConnection con = new OracleConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                OracleDataAdapter da = new OracleDataAdapter(cmdText, con);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// OracleDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        ///  <param name="cmdText">OracleParameter[]参数</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataTable Oracle_DataAdapter(string connectionString, string cmdText,ref OracleParameter[] Parameters, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OracleConnection con = new OracleConnection();
            DataTable dt = new DataTable();
            try
            {
                con.ConnectionString = connectionString;
                OracleCommand com = new OracleCommand();
                com.Connection = con;
                com.CommandText = cmdText;
                if (Parameters.Length > 0)
                {
                    foreach (OracleParameter OracleParameter in Parameters)
                        com.Parameters.Add(OracleParameter);
                }
                OracleDataAdapter da = new OracleDataAdapter(com);
                con.Open();
                da.Fill(dt);
                con.Close();
                da.Dispose();
                return dt;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return dt;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

        /// <summary>
        /// OracleDataAdapter
        /// </summary>
        /// <param name="connectionString">连接字符串</param>
        /// <param name="cmdText">执行语句</param>
        /// <param name="Exceptionmessage">异常消息</param>
        /// <returns></returns>
        public DataSet Oracle_DataAdapter_DataSet(string connectionString, string cmdText, ref string Exceptionmessage)
        {
            Exceptionmessage = "";
            OracleConnection con = new OracleConnection();
            DataSet ds = new DataSet();
            try
            {
                con.ConnectionString = connectionString;
                OracleDataAdapter da = new OracleDataAdapter(cmdText, con);
                con.Open();
                da.Fill(ds);
                con.Close();
                da.Dispose();
                return ds;
            }
            catch (Exception ee)
            {
                Exceptionmessage = ee.Message;
                con.Close();
                return ds;
            }
            finally
            {
                if (con.State == ConnectionState.Open)
                    con.Close();
            }

        }

    }

}
