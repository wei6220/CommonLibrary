using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace CommonLibrary.Extensions
{
    /// <summary>
    /// DataTable 擴充方法
    /// </summary>
    public static class DataTableExtensions
    {
        /// <summary>
        /// (擴充方法) DataTable -> model List
        /// </summary>
        /// <typeparam name="T">db model</typeparam>
        /// <param name="table">資料源</param>
        /// <returns>model list</returns>
        public static IList<T> ToList<T>(this DataTable table) where T : new()
        {
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            IList<T> result = new List<T>();

            foreach (var row in table.Rows)
            {
                var item = CreateItemFromRow<T>((DataRow)row, properties);
                result.Add(item);
            }
            return result;
        }
        /// <summary>
        /// (擴充方法) DataTable -> model List
        /// 若db columns與model properties不同,可自訂map表
        /// </summary>
        /// <typeparam name="T">db model</typeparam>
        /// <param name="table">資料源</param>
        /// <param name="mappings">自訂map表,format{model property,db column}</param>
        /// <returns>model list</returns>
        public static IList<T> ToList<T>(this DataTable table, Dictionary<string, string> mappings) where T : new()
        {
            IList<PropertyInfo> properties = typeof(T).GetProperties().ToList();
            IList<T> result = new List<T>();

            foreach (var row in table.Rows)
            {
                var item = CreateItemFromRow<T>((DataRow)row, properties, mappings);
                result.Add(item);
            }
            return result;
        }

        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties) where T : new()
        {
            T item = new T();
            foreach (var property in properties)
            {
                if (row[property.Name] == System.DBNull.Value)
                    property.SetValue(item, null, null);
                else
                    property.SetValue(item, row[property.Name], null);
            }
            return item;
        }

        private static T CreateItemFromRow<T>(DataRow row, IList<PropertyInfo> properties, Dictionary<string, string> mappings) where T : new()
        {
            T item = new T();
            foreach (var property in properties)
            {
                if (mappings.ContainsKey(property.Name))
                {
                    if (row[mappings[property.Name]] == System.DBNull.Value)
                        property.SetValue(item, null, null);
                    else
                    {
                        if (property.PropertyType == typeof(int))
                            property.SetValue(item, Convert.ToInt32(row[mappings[property.Name]]), null);
                        else if (property.PropertyType == typeof(System.Nullable<int>))
                            property.SetValue(item, Convert.ToInt32(row[mappings[property.Name]]), null);
                        else
                            property.SetValue(item, row[mappings[property.Name]], null);
                    }
                }
            }
            return item;
        }
    }
}
