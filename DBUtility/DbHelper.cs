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
        private string _ConnectionString;
        public DbHelper()
        {
            _ConnectionString = getConnectionString("DATABASE");
        }
        public DbHelper(string connectionString)
        {
            _ConnectionString = connectionString;
        }

        public static string getConnectionString(string section)
        {
            try
            {
                // Data Source =.\SQLEXPRESS; Initial Catalog = InsydeBIOSConfigurationDB; Integrated Security = True

                //by .ini
                //string server = IniHelper.getIniValue(section, "servername", IniHelper.getIniFilePath());
                //string db = IniHelper.getIniValue(section, "datasource", IniHelper.getIniFilePath());
                //string userid = IniHelper.getIniValue(section, "userid", IniHelper.getIniFilePath());
                //string password = IniHelper.getIniValue(section, "Password", IniHelper.getIniFilePath());

                //by .xml
                XmlHelper xmlHelper = new XmlHelper();
                IList list = xmlHelper.Read(IniHelper.getIniFilePath(), "SETTINGS/" + section);
                string server = xmlHelper.Find(list, "servername").ToString();
                string db = xmlHelper.Find(list, "datasource").ToString();
                string userid = xmlHelper.Find(list, "userid").ToString();
                string password = xmlHelper.Find(list, "Password").ToString();

                return @"Data Source =" + server
                    + "; Initial Catalog = " + db
                    + "; Integrated Security = false;User ID = " + userid
                    + "; Password = '" + password + "';";

            }
            catch (Exception ex)
            {
                LogHelper.Write("[DbHelper] getConnectionString error!!");
                throw new Exception(ex.Message);
            }
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
                    LogHelper.Write("[DbHelper] QuerySql error!!");
                    throw new Exception(ex.Message);
                }

            }
        }

    }
}
