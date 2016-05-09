using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;
using Oracle.ManagedDataAccess.Types;
namespace FG.PDMReader
{
    /// <summary>
    /// Desciption: Oracle数据库访问类 (注：针对 odp.net ). 
    /// Author    : siupan
    /// Date      : 2013-09-21
    /// odp.net manged
    /// </summary>
    public static class OracleHelper
    {
        #region [ 连接对象 ]
        /// <summary>
        /// 连接对象 字段
        /// </summary>
        private static OracleConnection conn = null;
        /// <summary>
        /// 连接串 字段
        /// </summary>

        //private static string connstr = @"Data Source=(DESCRIPTION=(ADDRESS_LIST=(ADDRESS=(PROTOCOL=TCP)(HOST=localhost)(PORT=1521)))(CONNECT_DATA=(SERVER=DEDICATED)(SERVICE_NAME=ORCL)));User Id=leaf;Password=leaf;";
        private static string connstr = @"Data Source=172.16.130.68:1521/ORCL;Persist Security Info=True;User ID=FGMDM;Password=FGMDM;";

        /// <summary>
        /// 取得连接串
        /// </summary>
        public static string GetConnectionString
        {
            get
            {
                return connstr;
            }
        }

        public static string SetConnectionString
        {
            set
            {
                connstr = value;
            }
        }

        /// <summary>
        /// 取得连接对象, 没有打开
        /// </summary>
        public static OracleConnection GetOracleConnection
        {
            get
            {
                return new OracleConnection(GetConnectionString);
            }
        }

        /// <summary>
        /// 取得连接对象， 并打开
        /// </summary>
        public static OracleConnection GetOracleConnectionAndOpen
        {
            get
            {
                OracleConnection conn = GetOracleConnection;
                if (conn.State == ConnectionState.Closed)
                {
                    try
                    {
                        conn.Open();
                    }
                    catch (Exception ex)
                    {
                        return null;
                    }
                }
                
                return conn;
            }
        }

        /// <summary>
        /// 彻底关闭并释放 OracleConnection 对象，再置为null.
        /// </summary>
        /// <param name="conn">OracleConnection</param>
        public static void CloseOracleConnection(OracleConnection conn)
        {
            if (conn == null)
                return;
            conn.Close();
            conn.Dispose();
            conn = null;
        }
        #endregion

        #region [ ExecuteNonQuery ]
        /// <summary>
        /// 普通SQL语句执行增删改
        /// </summary>
        /// <param name="cmdText">SQL语句</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns>受影响行数</returns>
        public static int ExecuteNonQuery(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteNonQuery(cmdText, CommandType.Text, commandParameters);
        }
        /// <summary>
        /// 存储过程执行增删改
        /// </summary>
        /// <param name="cmdText">存储过程</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns>受影响行数</returns>
        public static int ExecuteNonQueryByProc(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteNonQuery(cmdText, CommandType.StoredProcedure, commandParameters);
        }
        /// <summary>
        /// 执行增删改
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns>受影响行数</returns>
        public static int ExecuteNonQuery(string cmdText, CommandType cmdType, params OracleParameter[] commandParameters)
        {
            int result = 0;
            OracleConnection conn = null;

            try
            {
                conn = GetOracleConnectionAndOpen;
                OracleCommand command = new OracleCommand();
                PrepareCommand(command, conn, cmdType, cmdText, commandParameters);
                result = command.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                result = -1;
            }
            finally
            {
                if (conn != null)
                    CloseOracleConnection(conn);
            }

            return result;
        }
        #endregion

        #region [ ExecuteReader ]
        /// <summary>
        /// SQL语句得到 OracleDataReader 对象
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns>OracleDataReader 对象</returns>
        public static OracleDataReader ExecuteReader(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteReader(cmdText, CommandType.Text, commandParameters);
        }
        /// <summary>
        /// 存储过程得到 OracleDataReader 对象
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns>OracleDataReader 对象</returns>
        public static OracleDataReader ExecuteReaderByProc(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteReader(cmdText, CommandType.StoredProcedure, commandParameters);
        }
        /// <summary>
        /// 得到 OracleDataReader 对象
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns>OracleDataReader 对象</returns>
        public static OracleDataReader ExecuteReader(string cmdText, CommandType cmdType, params OracleParameter[] commandParameters)
        {
            OracleDataReader result = null;
            OracleConnection conn = null;

            try
            {
                conn = GetOracleConnectionAndOpen;
                OracleCommand command = new OracleCommand();
                PrepareCommand(command, conn, cmdType, cmdText, commandParameters);
                result = command.ExecuteReader(CommandBehavior.CloseConnection);
            }
            catch (Exception ex)
            {
                result = null;
            }
            finally
            {
                if (conn != null)
                    CloseOracleConnection(conn);
            }

            return result;
        }
        #endregion

        #region [ ExecuteScalar ]
        /// <summary>
        /// 执行SQL语句, 返回Object
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> Object </returns>
        public static Object ExecuteScalar(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteScalar(cmdText, CommandType.Text, commandParameters);
        }

        /// <summary>
        /// 执行存储过程, 返回Object
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> Object </returns>
        public static Object ExecuteScalarByProc(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteScalar(cmdText, CommandType.StoredProcedure, commandParameters);
        }

        /// <summary>
        /// 返回Object
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> Object </returns>
        public static Object ExecuteScalar(string cmdText, CommandType cmdType, params OracleParameter[] commandParameters)
        {
            Object result = null;
            OracleConnection conn = null;

            try
            {
                conn = GetOracleConnectionAndOpen;
                OracleCommand command = new OracleCommand();

                PrepareCommand(command, conn, cmdType, cmdText, commandParameters);
                result = command.ExecuteScalar();
            }
            catch (Exception ex)
            {
                result = null;
            }
            finally
            {
                if (conn != null)
                    CloseOracleConnection(conn);
            }

            return result;
        }
        #endregion

        #region [ ExecuteDataSet ]
        /// <summary>
        /// 执行SQL语句, 返回DataSet
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> DataSet </returns>
        public static DataSet ExecuteDataSet(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteDataSet(cmdText, CommandType.Text, commandParameters);
        }

        /// <summary>
        /// 执行存储过程, 返回DataSet
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> DataSet </returns>
        public static DataSet ExecuteDataSetByProc(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteDataSet(cmdText, CommandType.StoredProcedure, commandParameters);
        }

        /// <summary>
        /// 返回DataSet
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> DataSet </returns>
        public static DataSet ExecuteDataSet(string cmdText, CommandType cmdType, params OracleParameter[] commandParameters)
        {
            DataSet result = null;
            OracleConnection conn = null;

            try
            {
                conn = GetOracleConnectionAndOpen;
                OracleCommand command = new OracleCommand();

                PrepareCommand(command, conn, cmdType, cmdText, commandParameters);
                OracleDataAdapter adapter = new OracleDataAdapter();
                adapter.SelectCommand = command;
                result = new DataSet();
                adapter.Fill(result);
            }
            catch (Exception ex)
            {
                result = null;
            }
            finally
            {
                if (conn != null)
                    CloseOracleConnection(conn);
            }

            return result;
        }
        #endregion

        #region [ ExecuteDataTable ]
        /// <summary>
        /// 执行SQL语句, 返回DataTable
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> DataTable </returns>
        public static DataTable ExecuteDataTable(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteDataTable(cmdText, CommandType.Text, commandParameters);
        }

        /// <summary>
        /// 执行存储过程, 返回DataTable
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> DataTable </returns>
        public static DataTable ExecuteDataTableByProc(string cmdText, params OracleParameter[] commandParameters)
        {
            return ExecuteDataTable(cmdText, CommandType.StoredProcedure, commandParameters);
        }

        /// <summary>
        /// 返回DataTable
        /// </summary>
        /// <param name="cmdText">命令字符串</param>
        /// <param name="cmdType">命令类型</param>
        /// <param name="commandParameters">可变参数</param>
        /// <returns> DataTable </returns>
        public static DataTable ExecuteDataTable(string cmdText, CommandType cmdType, params OracleParameter[] commandParameters)
        {
            DataTable dtResult = null;
            DataSet ds = ExecuteDataSet(cmdText, cmdType, commandParameters);

            if (ds != null && ds.Tables.Count > 0)
            {
                dtResult = ds.Tables[0];
            }
            return dtResult;
        }
        #endregion

        #region [ PrepareCommand ]
        /// <summary>
        /// Command对象执行前预处理
        /// </summary>
        /// <param name="command"></param>
        /// <param name="connection"></param>
        /// <param name="trans"></param>
        /// <param name="cmdType"></param>
        /// <param name="cmdText"></param>
        /// <param name="commandParameters"></param>
        private static void PrepareCommand(OracleCommand command, OracleConnection connection, CommandType cmdType, string cmdText, OracleParameter[] commandParameters)
        {
            try
            {
                if (connection.State != ConnectionState.Open) connection.Open();

                command.Connection = connection;
                command.CommandText = cmdText;
                command.CommandType = cmdType;

                //if (trans != null) command.Transaction = trans;

                if (commandParameters != null)
                {
                    foreach (OracleParameter parm in commandParameters)
                        command.Parameters.Add(parm);
                }
            }
            catch
            {

            }
        }
        #endregion
    }//end of class
}//end of namespace