using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary.DBUtility
{
    /// <summary>
    /// sql execute model
    /// </summary>
    public class ExecuteModel
    {
        /// <summary>
        /// 執行字串
        /// </summary>
        public string SqlString { get; set; }
        /// <summary>
        /// 執行參數
        /// </summary>
        public List<IDbDataParameter> SqlParams { get; set; }
    }
}
