using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;

namespace CommonLibrary.DBUtility
{
    /// <summary>
    /// Dao interface
    /// </summary>
    public interface IDbHelper
    {
        /// <summary>
        /// 執行sql command
        /// </summary>
        /// <param name="SQLString">sql字串</param>
        /// <param name="SqlParams">參數</param>
        /// <returns>影響筆數</returns>
        int ExecuteSql(string SQLString, List<IDbDataParameter> SqlParams);

        /// <summary>
        /// 執行sql command(多筆)
        /// </summary>
        /// <param name="executeList">sql執行model</param>
        /// <returns>影響筆數</returns>
        int ExecuteSql(List<ExecuteModel> executeList);

        /// <summary>
        /// 查詢sql command
        /// </summary>
        /// <param name="SqlString">sql字串</param>
        /// <param name="SqlParams">參數</param>
        /// <returns> 回傳List{Dictionary(string,object)} </returns>
        IList QuerySql(string SqlString, List<IDbDataParameter> SqlParams);
        
        /// <summary>
        /// 查詢sql command
        /// </summary>
        /// <param name="SqlString">sql字串</param>
        /// <param name="SqlParams">參數</param>
        /// <param name="type">資料模型</param>
        /// <returns> List(資料模型) </returns>
        IList QuerySql(string SqlString, List<IDbDataParameter> SqlParams, Type type);
    }
}