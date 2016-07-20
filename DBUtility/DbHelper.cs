using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Collections;

namespace CommonLibrary.DBUtility
{
    public class DbHelper
    {
        public DbHelper()
        {
            _ConnectionString = IniHelper.getConnectionString(IniHelper.getIniFileFullName(), "database");
            //_ConnectionString = @"Data Source=.\SQLEXPRESS;Initial Catalog=InsydeBIOSConfigurationDB;Integrated Security=True";
        }
        public DbHelper(string connectionString)
        {
            _ConnectionString = connectionString;
        }

        private string _ConnectionString;
        public string getConnectionString()
        {
            return _ConnectionString;
        }

        public int ExecuteSql(string SQLString, List<IDbDataParameter> SqlParams)
        {
            using (SqlConnection connection = new SqlConnection(_ConnectionString))
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }

                using (SqlCommand cmd = new SqlCommand(SQLString, connection))
                {
                    try
                    {
                        int rows = cmd.ExecuteNonQuery();
                        return rows;
                    }
                    catch (System.Data.SqlClient.SqlException e)
                    {
                        connection.Close();
                        throw e;
                    }
                }
            }
        }

        public IList QuerySql(string SqlString, List<IDbDataParameter> SqlParams)
        {
            return QuerySql(SqlString, SqlParams, null);
        }
        public IList QuerySql(string SqlString, List<IDbDataParameter> SqlParams,Type type)
        {
            IList data = new List<dynamic>();
            using (SqlConnection connection = new SqlConnection(_ConnectionString))
            {
                try
                {
                    if (connection.State == ConnectionState.Closed)
                    {
                        connection.Open();
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
                            if (type != null)
                            {
                                //自訂類別
                                model = Activator.CreateInstance(type);
                                System.Reflection.PropertyInfo[] Propertys = type.GetProperties();
                                foreach (System.Reflection.PropertyInfo pi in Propertys)
                                {
                                    //檢視db是否該欄位
                                    for (int i = 0; i < dr.FieldCount; i++)
                                    {
                                        if (string.Compare(dr.GetName(i),pi.Name,true)==0)
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
                    throw new Exception(ex.Message);
                }

            }
        }
    }
}
