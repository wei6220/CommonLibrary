using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Collections;
using System.IO;

namespace CommonLibrary.DBUtility
{
    /// <summary>
    /// Dao元件
    /// </summary>
    public class DbHelper : IDbHelper
    {
        private string _connectionString;
        /// <summary>
        /// init DbHelper
        /// 預設DB連線字串 => 設定檔section:DATABASE
        /// </summary>
        public DbHelper()
        {
            _connectionString = GetConnectionString("DATABASE", IniHelper.GetFilePath());
        }
        /// <summary>
        /// init DbHelper
        /// </summary>
        /// <param name="connectionString">自訂DB連線字串</param>
        public DbHelper(string connectionString)
        {
            _connectionString = connectionString;
        }

        /// <summary>
        /// 執行sql command
        /// </summary>
        /// <param name="SQLString">sql字串</param>
        /// <param name="SqlParams">參數</param>
        /// <returns>影響筆數</returns>
        public int ExecuteSql(string SQLString, List<IDbDataParameter> SqlParams)
        {
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    try
                    {
                        int cnt = 0;
                        using (SqlCommand cmd = new SqlCommand(SQLString, connection, tran))
                        {
                            if (SqlParams != null && SqlParams.Count() > 0)
                            {
                                foreach (SqlParameter para in SqlParams)
                                {
                                    //if (para.SqlDbType == SqlDbType.NVarChar)
                                    //{
                                    //    para.Value = HttpContext.Current.Server.HtmlDecode(para.Value.ToString());
                                    //}
                                    cmd.Parameters.Add(para);
                                }
                            }

                            int rows = cmd.ExecuteNonQuery();
                            cnt += rows;
                        }

                        tran.Commit();
                        return cnt;
                    }
                    catch (SqlException e)
                    {
                        tran.Rollback();
                        throw e;
                    }
                }
            }
        }

        /// <summary>
        /// 執行sql command(多筆)
        /// </summary>
        /// <param name="executeList">sql執行model</param>
        /// <returns>影響筆數</returns>
        public int ExecuteSql(List<ExecuteModel> executeList)
        {
            if (executeList == null || executeList.Count == 0)
                return 0;

            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                using (SqlTransaction tran = connection.BeginTransaction())
                {
                    try
                    {
                        int cnt = 0;

                        foreach (var exe in executeList)
                        {
                            using (SqlCommand cmd = new SqlCommand(exe.SqlString, connection, tran))
                            {
                                if (exe.SqlParams != null && exe.SqlParams.Count() > 0)
                                {
                                    foreach (SqlParameter para in exe.SqlParams)
                                    {
                                        //if (para.SqlDbType == SqlDbType.NVarChar)
                                        //{
                                        //    para.Value = HttpContext.Current.Server.HtmlDecode(para.Value.ToString());
                                        //}
                                        cmd.Parameters.Add(para);
                                    }
                                }

                                int rows = cmd.ExecuteNonQuery();
                                cnt += rows;
                            }
                        }

                        tran.Commit();
                        return cnt;
                    }
                    catch (SqlException e)
                    {
                        tran.Rollback();
                        throw e;
                    }
                }
            }
        }

        /// <summary>
        /// 查詢sql command
        /// </summary>
        /// <param name="SqlString">sql字串</param>
        /// <param name="SqlParams">參數</param>
        /// <returns> 回傳List{Dictionary(string,object)} </returns>
        public IList QuerySql(string SqlString, List<IDbDataParameter> SqlParams)
        {
            return QuerySql(SqlString, SqlParams, null);
        }

        /// <summary>
        /// 查詢sql command
        /// </summary>
        /// <param name="SqlString">sql字串</param>
        /// <param name="SqlParams">參數</param>
        /// <param name="modelType">資料模型</param>
        /// <returns> List(資料模型) </returns>
        public IList QuerySql(string SqlString, List<IDbDataParameter> SqlParams, Type modelType)
        {
            IList data = new List<dynamic>();
            using (SqlConnection connection = new SqlConnection(_connectionString))
            {
                try
                {
                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
                        LogHelper.Write(string.Format("Open Connection :{0}", _connectionString));
                    }

                    SqlCommand cmd = new SqlCommand(SqlString, connection);
                    if (SqlParams != null && SqlParams.Count() > 0)
                    {
                        foreach (SqlParameter para in SqlParams)
                        {
                            //if (para.SqlDbType == SqlDbType.NVarChar)
                            //{
                            //    para.Value = HttpContext.Current.Server.HtmlDecode(para.Value.ToString());
                            //}
                            cmd.Parameters.Add(para);
                        }
                    }

                    using (SqlDataReader dr = cmd.ExecuteReader())
                    {
                        dynamic model;
                        while ((dr.Read()))
                        {
                            if (modelType != null)
                            {
                                //自訂類別
                                model = Activator.CreateInstance(modelType);
                                System.Reflection.PropertyInfo[] Propertys = modelType.GetProperties();
                                foreach (System.Reflection.PropertyInfo pi in Propertys)
                                {
                                    //檢視db是否該欄位
                                    for (int i = 0; i < dr.FieldCount; i++)
                                    {
                                        if (string.Compare(dr.GetName(i), pi.Name, true) == 0)
                                            pi.SetValue(model, dr[pi.Name]);
                                    }
                                }
                            }
                            else
                            {
                                model = new Dictionary<string, object>();
                                for (int i = 0; i < dr.FieldCount; i++)
                                    model[dr.GetName(i)] = dr[i];
                            }

                            data.Add(model);
                        }
                    }
                    return data;

                }
                catch (Exception ex)
                {
                    LogHelper.Write("[DbHelper] QuerySql error!!");
                    LogHelper.Write(ex.Message);
                    throw new Exception(ex.Message);
                }

            }
        }

        /// <summary>
        /// 取得連線字串
        /// </summary>
        /// <param name="section">設定檔section</param>
        /// <param name="iniPath">設定檔位置</param>
        /// <returns></returns>
        private string GetConnectionString(string section, string iniPath)
        {
            try
            {
                //IniHelper.IniType type = (string.Compare(Path.GetExtension(iniPath), ".xml", true) == 0) ? IniHelper.IniType.xml : IniHelper.IniType.ini;
                return @"Data Source =" + IniHelper.GetValue(section, "servername")
                         + "; Initial Catalog = " + IniHelper.GetValue(section, "datasource")
                         + "; Integrated Security = false"
                         + "; User ID = " + IniHelper.GetValue(section, "userid")
                         + "; Password = '" + IniHelper.GetValue(section, "Password") + "';";
            }
            catch (Exception ex)
            {
                LogHelper.Write("[DbHelper] GetConnectionString error!!");
                throw new Exception(ex.Message);
            }
        }

    }
}
