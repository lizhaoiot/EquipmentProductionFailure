/*
 * 类说明：根据对象执行SQL语句
 */
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.SqlClient;
using System.Reflection;

namespace SQL
{
    /// <summary>
    /// SQL帮助类
    /// </summary>
    public class SqlHelper
    {
        public static string connectionStr = null;
        public static DataTable GetUserList(string connectionStr, string sql, SqlParameter[] prams)
        {
            DataTable result = null;
            SqlConnection sqlConnection = new SqlConnection(connectionStr);
            sqlConnection.Open();
            using (SqlCommand sqlCommand = new SqlCommand(sql, sqlConnection))
            {
                sqlCommand.Parameters.AddRange(prams);
                DataSet dataSet = new DataSet();
                DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                dbDataAdapter.SelectCommand = sqlCommand;
                dbDataAdapter.Fill(dataSet);
                result = dataSet.Tables[0];
            }
            sqlConnection.Close();
            return result;
        }
        public static SqlConnection GetConnection()
        {
            if (connectionStr == null)
            {
                return GetConnection("packet size=4096;user id=sa;pwd=swumcf&653#78*SLK2123;data source=10.13.10.96;persist security info=False;initial catalog=POMS");
            }
            return GetConnection(connectionStr);
        }
        public static SqlConnection GetConnection(string connstr)
        {
            SqlConnection sqlConnection = null;
            sqlConnection = new SqlConnection(connstr);
            sqlConnection.Open();
            return sqlConnection;
        }
        public static void BulkCopy(DataTable src, string dest)
        {
            using (SqlConnection connection = GetConnection())
            {
                SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(connection);
                sqlBulkCopy.DestinationTableName = dest;
                sqlBulkCopy.WriteToServer(src);
            }
        }
        public static void BulkCopy(DataTable src, string dest, SqlConnection conn, SqlTransaction t)
        {
            SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(conn, SqlBulkCopyOptions.Default, t);
            sqlBulkCopy.DestinationTableName = dest;
            sqlBulkCopy.WriteToServer(src);
        }
        public static void ExecCommand(string commtext)
        {
            using (SqlConnection connection = GetConnection())
            {
                SqlCommand sqlCommand = new SqlCommand(commtext, connection);
                sqlCommand.CommandType = CommandType.Text;
                sqlCommand.ExecuteNonQuery();
                sqlCommand.Dispose();
            }
        }
        public static void ExecCommand(string[] commtext)
        {
            using (SqlConnection connection = GetConnection())
            {
                for (int i = 0; i < commtext.Length; i++)
                {
                    SqlCommand sqlCommand = new SqlCommand(commtext[i], connection);
                    sqlCommand.CommandType = CommandType.Text;
                    sqlCommand.ExecuteNonQuery();
                    sqlCommand.Dispose();
                }
            }
        }
        public static int ExecCommand(string commtext, SqlParameter[] prams)
        {
            using (SqlConnection connection = GetConnection())
            {
                int num = -1;
                SqlCommand sqlCommand = new SqlCommand(commtext, connection);
                sqlCommand.CommandType = CommandType.Text;
                if (prams != null)
                {
                    sqlCommand.Parameters.AddRange(prams);
                }
                num = sqlCommand.ExecuteNonQuery();
                sqlCommand.Dispose();
                return num;
            }
        }
        public static void ExecCommand(string commtext, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
        {
            SqlCommand sqlCommand = new SqlCommand(commtext, conn);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Transaction = t;
            if (prams != null)
            {
                sqlCommand.Parameters.AddRange(prams);
            }
            sqlCommand.ExecuteNonQuery();
        }
        public static void ExecCommand(string commtext, SqlConnection conn, SqlTransaction t)
        {
            SqlCommand sqlCommand = new SqlCommand(commtext, conn);
            sqlCommand.CommandType = CommandType.Text;
            sqlCommand.Transaction = t;
            sqlCommand.ExecuteNonQuery();
        }
        public static int ExecStoredProcedure(string procedureName, SqlParameter[] prams)
        {
            using (SqlConnection connection = GetConnection())
            {
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.CommandTimeout = 600;
                sqlCommand.Connection = connection;
                sqlCommand.CommandType = CommandType.StoredProcedure;
                sqlCommand.CommandText = procedureName;
                if (prams != null)
                {
                    sqlCommand.Parameters.AddRange(prams);
                }
                int result = sqlCommand.ExecuteNonQuery();
                sqlCommand.Dispose();
                sqlCommand.Parameters.Clear();
                return result;
            }
        }
        public static DataTable ExecStoredProcedureDataTable(string procedureName, SqlParameter[] prams)
        {
            using (SqlConnection connection = GetConnection())
            {
                DataTable dataTable = null;
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.CommandTimeout = 600;
                sqlCommand.Connection = connection;
                sqlCommand.CommandType = CommandType.StoredProcedure;
                sqlCommand.CommandText = procedureName;
                if (prams != null)
                {
                    sqlCommand.Parameters.AddRange(prams);
                }
                DataSet dataSet = new DataSet();
                DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                dbDataAdapter.SelectCommand = sqlCommand;
                dbDataAdapter.Fill(dataSet);
                if (dataSet.Tables.Count == 0)
                {
                    return null;
                }
                dataTable = dataSet.Tables[0];
                sqlCommand.Dispose();
                sqlCommand.Parameters.Clear();
                return dataTable;
            }
        }
        public static DataSet ExecStoredProcedureDataSet(string procedureName, SqlParameter[] prams)
        {
            using (SqlConnection connection = GetConnection())
            {
                DataSet dataSet = new DataSet();
                SqlCommand sqlCommand = new SqlCommand();
                sqlCommand.CommandTimeout = 600;
                sqlCommand.Connection = connection;
                sqlCommand.CommandType = CommandType.StoredProcedure;
                sqlCommand.CommandText = procedureName;
                if (prams != null)
                {
                    sqlCommand.Parameters.AddRange(prams);
                }
                DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                dbDataAdapter.SelectCommand = sqlCommand;
                dbDataAdapter.Fill(dataSet);
                sqlCommand.Dispose();
                return dataSet;
            }
        }
        public static DataSet ExecStoredProcedureDataSet(string procedureName, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
        {
            DataSet dataSet = new DataSet();
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.CommandTimeout = 600;
            sqlCommand.Connection = conn;
            sqlCommand.Transaction = t;
            sqlCommand.CommandType = CommandType.StoredProcedure;
            sqlCommand.CommandText = procedureName;
            if (prams != null)
            {
                sqlCommand.Parameters.AddRange(prams);
            }
            DbDataAdapter dbDataAdapter = new SqlDataAdapter();
            dbDataAdapter.SelectCommand = sqlCommand;
            dbDataAdapter.Fill(dataSet);
            return dataSet;
        }
        public static DataTable ExecStoredProcedureDataTable(string procedureName, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
        {
            DataTable dataTable = null;
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.CommandTimeout = 600;
            sqlCommand.Connection = conn;
            sqlCommand.Transaction = t;
            sqlCommand.CommandType = CommandType.StoredProcedure;
            sqlCommand.CommandText = procedureName;
            if (prams != null)
            {
                sqlCommand.Parameters.AddRange(prams);
            }
            DataSet dataSet = new DataSet();
            DbDataAdapter dbDataAdapter = new SqlDataAdapter();
            dbDataAdapter.SelectCommand = sqlCommand;
            dbDataAdapter.Fill(dataSet);
            return dataSet.Tables[0];
        }
        public static int ExecStoredProcedure(string procedureName, SqlParameter[] prams, SqlConnection conn, SqlTransaction t)
        {
            SqlCommand sqlCommand = new SqlCommand();
            sqlCommand.Connection = conn;
            sqlCommand.Transaction = t;
            sqlCommand.CommandType = CommandType.StoredProcedure;
            sqlCommand.CommandText = procedureName;
            if (prams != null)
            {
                sqlCommand.Parameters.AddRange(prams);
            }
            return sqlCommand.ExecuteNonQuery();
        }
        public static int ExecScalar(string CommText)
        {
            using (SqlConnection connection = GetConnection())
            {
                SqlCommand sqlCommand = new SqlCommand(CommText, connection);
                sqlCommand.CommandType = CommandType.Text;
                return (sqlCommand.ExecuteScalar() != DBNull.Value) ? ((int)sqlCommand.ExecuteScalar()) : 0;
            }
        }
        public static int ExecScalar(string CommText, SqlConnection TConn)
        {
            SqlCommand sqlCommand = new SqlCommand(CommText, TConn);
            sqlCommand.CommandType = CommandType.Text;
            return (sqlCommand.ExecuteScalar() != DBNull.Value) ? ((int)sqlCommand.ExecuteScalar()) : 0;
        }
        public static int ExecScalar(string CommText, SqlTransaction t)
        {
            using (SqlConnection connection = GetConnection())
            {
                SqlCommand sqlCommand = new SqlCommand(CommText, connection);
                sqlCommand.Transaction = t;
                sqlCommand.CommandType = CommandType.Text;
                return (sqlCommand.ExecuteScalar() != DBNull.Value) ? Convert.ToInt32(sqlCommand.ExecuteScalar()) : 0;
            }
        }
        public static int ExecScalar(string CommText, SqlConnection TConn, SqlTransaction t)
        {
            SqlCommand sqlCommand = new SqlCommand(CommText, TConn);
            sqlCommand.Transaction = t;
            sqlCommand.CommandType = CommandType.Text;
            return (sqlCommand.ExecuteScalar() != DBNull.Value) ? Convert.ToInt32(sqlCommand.ExecuteScalar()) : 0;
        }
        public static int ExecScalar(string CommText, SqlParameter[] prams)
        {
            using (SqlConnection connection = GetConnection())
            {
                SqlCommand sqlCommand = new SqlCommand(CommText, connection);
                sqlCommand.CommandType = CommandType.Text;
                if (prams != null)
                {
                    sqlCommand.Parameters.AddRange(prams);
                }
                return (sqlCommand.ExecuteScalar() != DBNull.Value) ? Convert.ToInt32(sqlCommand.ExecuteScalar()) : 0;
            }
        }
        public static int ExecScalar(string CommText, SqlParameter[] prams, SqlConnection TConn)
        {
            SqlCommand sqlCommand = new SqlCommand(CommText, TConn);
            sqlCommand.CommandType = CommandType.Text;
            if (prams != null)
            {
                foreach (SqlParameter value in prams)
                {
                    sqlCommand.Parameters.Add(value);
                }
            }
            return (sqlCommand.ExecuteScalar() != DBNull.Value) ? Convert.ToInt32(sqlCommand.ExecuteScalar()) : 0;
        }
        public static DataTable ExecuteDataTable(string sql)
        {
            using (SqlConnection connection = GetConnection())
            {
                DataTable result = null;
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    DataSet dataSet = new DataSet();
                    DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                    dbDataAdapter.SelectCommand = sqlCommand;
                    dbDataAdapter.Fill(dataSet);
                    result = dataSet.Tables[0];
                    sqlCommand.Parameters.Clear();
                }
                return result;
            }
        }
        public static DataTable ExecuteDataTable(string sql, SqlConnection conn, SqlTransaction t)
        {
            DataTable result = null;
            using (SqlCommand sqlCommand = new SqlCommand(sql, conn))
            {
                sqlCommand.Transaction = t;
                DataSet dataSet = new DataSet();
                DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                dbDataAdapter.SelectCommand = sqlCommand;
                dbDataAdapter.Fill(dataSet);
                result = dataSet.Tables[0];
                sqlCommand.Parameters.Clear();
            }
            return result;
        }
        public static DataTable ExecuteDataTable(string sql, DbParameter[] parameters)
        {
            DataTable result = null;
            using (SqlConnection connection = GetConnection())
            {
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    if (parameters != null)
                    {
                        sqlCommand.Parameters.AddRange(parameters);
                    }
                    sqlCommand.CommandTimeout = 600;
                    DataSet dataSet = new DataSet();
                    DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                    dbDataAdapter.SelectCommand = sqlCommand;
                    dbDataAdapter.Fill(dataSet);
                    result = dataSet.Tables[0];
                    sqlCommand.Parameters.Clear();
                }
                return result;
            }
        }
        public static DataTable ExecuteDataTable(string sql, DbParameter[] parameters, string connectionString)
        {
            DataTable result = null;
            using (SqlConnection connection = GetConnection(connectionString))
            {
                using (SqlCommand sqlCommand = new SqlCommand(sql, connection))
                {
                    sqlCommand.Parameters.AddRange(parameters);
                    sqlCommand.CommandTimeout = 600;
                    DataSet dataSet = new DataSet();
                    DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                    dbDataAdapter.SelectCommand = sqlCommand;
                    dbDataAdapter.Fill(dataSet);
                    result = dataSet.Tables[0];
                    sqlCommand.Parameters.Clear();
                }
                return result;
            }
        }
        public static DataTable ExecuteDataTable(string sql, DbParameter[] parameters, SqlConnection conn, SqlTransaction t)
        {
            DataTable result = null;
            using (SqlCommand sqlCommand = new SqlCommand(sql, conn))
            {
                sqlCommand.Transaction = t;
                sqlCommand.Parameters.AddRange(parameters);
                DataSet dataSet = new DataSet();
                DbDataAdapter dbDataAdapter = new SqlDataAdapter();
                dbDataAdapter.SelectCommand = sqlCommand;
                dbDataAdapter.Fill(dataSet);
                result = dataSet.Tables[0];
                sqlCommand.Parameters.Clear();
            }
            return result;
        }
        public static void Exec4DS(string CmdText, string TableName, out DataSet DS)
        {
            using (SqlConnection selectConnection = GetConnection())
            {
                DS = new DataSet();
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(CmdText, selectConnection);
                sqlDataAdapter.Fill(DS, TableName);
            }
        }
        public static DataSet ExecuteDataSet(string sql, SqlParameter[] prams)
        {
            using (SqlConnection connection = GetConnection())
            {
                DataSet dataSet = new DataSet();
                SqlCommand sqlCommand = new SqlCommand(sql, connection);
                sqlCommand.CommandType = CommandType.Text;
                if (prams != null)
                {
                    sqlCommand.Parameters.AddRange(prams);
                }
                SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                sqlDataAdapter.Fill(dataSet);
                sqlCommand.Parameters.Clear();
                return dataSet;
            }
        }
        public static void ExecXDS(string[] sqlCmds, string[] TableNames, out DataSet DS)
        {
            using (SqlConnection selectConnection = GetConnection())
            {
                DS = new DataSet();
                for (int i = 0; i < sqlCmds.Length; i++)
                {
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmds[i], selectConnection);
                    sqlDataAdapter.Fill(DS, TableNames[i]);
                }
            }
        }
        public static void ExecXDS(string[] sqlCmds, string[] TableNames, out DataSet DS, string connectionString)
        {
            using (SqlConnection selectConnection = GetConnection(connectionString))
            {
                DS = new DataSet();
                for (int i = 0; i < sqlCmds.Length; i++)
                {
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCmds[i], selectConnection);
                    sqlDataAdapter.Fill(DS, TableNames[i]);
                }
            }
        }
        public static void ExecXDS(string[] sqlCmds, string[] TableNames, SqlParameter[] prams, out DataSet DS)
        {
            using (SqlConnection connection = GetConnection())
            {
                DS = new DataSet();
                for (int i = 0; i < sqlCmds.Length; i++)
                {
                    SqlCommand sqlCommand = new SqlCommand(sqlCmds[i]);
                    sqlCommand.Parameters.Add(prams[i]);
                    sqlCommand.Connection = connection;
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                    sqlDataAdapter.Fill(DS, TableNames[i]);
                }
            }
        }
        public static void ExecXDS(string[] sqlCmds, string[] TableNames, List<SqlParameter[]> prams, out DataSet DS)
        {
            using (SqlConnection connection = GetConnection())
            {
                DS = new DataSet();
                for (int i = 0; i < sqlCmds.Length; i++)
                {
                    SqlCommand sqlCommand = new SqlCommand(sqlCmds[i]);
                    SqlParameter[] values = prams[i];
                    sqlCommand.Parameters.AddRange(values);
                    sqlCommand.Connection = connection;
                    SqlDataAdapter sqlDataAdapter = new SqlDataAdapter(sqlCommand);
                    sqlDataAdapter.Fill(DS, TableNames[i]);
                }
            }
        }
        public static DateTime GetDateTimeFromSQL(SqlConnection conn, SqlTransaction t)
        {
            string cmdText = "select getdate()";
            SqlCommand sqlCommand = new SqlCommand(cmdText, conn);
            sqlCommand.Transaction = t;
            return (DateTime)sqlCommand.ExecuteScalar();
        }
        public static DateTime GetDateTimeFromSQL()
        {
            using (SqlConnection connection = GetConnection())
            {
                string cmdText = "select getdate()";
                SqlCommand sqlCommand = new SqlCommand(cmdText, connection);
                return (DateTime)sqlCommand.ExecuteScalar();
            }
        }
        public static DataTable Join(DataTable left, DataTable right, DataColumn[] leftCols, DataColumn[] rightCols, bool includeLeftJoin, bool includeRightJoin)
        {
            DataTable dataTable = new DataTable("JoinResult");
            using (DataSet dataSet = new DataSet())
            {
                dataSet.Tables.AddRange(new DataTable[2]
                {
                left.Copy(),
                right.Copy()
                });
                DataColumn[] array = new DataColumn[leftCols.Length];
                for (int i = 0; i < leftCols.Length; i++)
                {
                    array[i] = dataSet.Tables[0].Columns[leftCols[i].ColumnName];
                }
                DataColumn[] array2 = new DataColumn[rightCols.Length];
                for (int i = 0; i < rightCols.Length; i++)
                {
                    array2[i] = dataSet.Tables[1].Columns[rightCols[i].ColumnName];
                }
                for (int i = 0; i < left.Columns.Count; i++)
                {
                    dataTable.Columns.Add(left.Columns[i].ColumnName, left.Columns[i].DataType);
                }
                for (int i = 0; i < right.Columns.Count; i++)
                {
                    string text = right.Columns[i].ColumnName;
                    while (dataTable.Columns.Contains(text))
                    {
                        text += "_2";
                    }
                    dataTable.Columns.Add(text, right.Columns[i].DataType);
                }
                DataRelation relation = new DataRelation("rLeft", array, array2, false);
                dataSet.Relations.Add(relation);
                dataTable.BeginLoadData();
                foreach (DataRow row in dataSet.Tables[0].Rows)
                {
                    DataRow[] childRows = row.GetChildRows(relation);
                    if (childRows != null && childRows.Length > 0)
                    {
                        object[] itemArray = row.ItemArray;
                        DataRow[] array3 = childRows;
                        foreach (DataRow dataRow2 in array3)
                        {
                            object[] itemArray2 = dataRow2.ItemArray;
                            object[] array4 = new object[itemArray.Length + itemArray2.Length];
                            Array.Copy(itemArray, 0, array4, 0, itemArray.Length);
                            Array.Copy(itemArray2, 0, array4, itemArray.Length, itemArray2.Length);
                            dataTable.LoadDataRow(array4, true);
                        }
                    }
                    else if (includeLeftJoin)
                    {
                        object[] itemArray = row.ItemArray;
                        object[] array4 = new object[itemArray.Length];
                        Array.Copy(itemArray, 0, array4, 0, itemArray.Length);
                        dataTable.LoadDataRow(array4, true);
                    }
                }
                if (includeRightJoin)
                {
                    DataRelation relation2 = new DataRelation("rRight", array2, array, false);
                    dataSet.Relations.Add(relation2);
                    foreach (DataRow row2 in dataSet.Tables[1].Rows)
                    {
                        DataRow[] childRows = row2.GetChildRows(relation2);
                        if (childRows == null || childRows.Length == 0)
                        {
                            object[] itemArray = row2.ItemArray;
                            object[] array4 = new object[dataTable.Columns.Count];
                            Array.Copy(itemArray, 0, array4, array4.Length - itemArray.Length, itemArray.Length);
                            dataTable.LoadDataRow(array4, true);
                        }
                    }
                }
                dataTable.EndLoadData();
            }
            return dataTable;
        }
    }

   


}
