using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;

namespace ExcelToolkitAddIn
{
    public class Sqlite
    {
        public SQLiteCommand cmd;
        public SQLiteConnection conn;
        public string dbPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Excel_Toolkit.db");
        public SQLiteDataReader reader;
        public Sqlite()
        {
            conn = new SQLiteConnection($"Data Source={dbPath};Version=3;");
            conn.Open();
        }
        public void Close()
        {
            conn.Close();
            conn.Dispose();
        }

        #region 获取数据
        public List<Dictionary<string, object>> GetData(string sql)
        {
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            cmd = new SQLiteCommand(sql, conn);
            reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                var row = new Dictionary<string, object>();
                for (var i = 0; i < reader.FieldCount; i++)
                {
                    var columnName = reader.GetName(i);
                    var value = reader.GetValue(i);
                    row[columnName] = value;
                    Console.WriteLine(value);
                }
                list.Add(row);
            }
            reader.Close();
            cmd.Dispose();
            return list;
        }

        public List<Dictionary<string, object>> GetTableAll(string tableName)
        {
            var sql = $"SELECT * FROM {tableName}";
            return GetData(sql);
        }

        public int GetMaxId(string tableName)
        {
            var sql = $"SELECT MAX(id) AS id FROM {tableName}";
            object result = GetData(sql)[0]["id"];
            return result == null || result == DBNull.Value ? 0 : Convert.ToInt32(result);
        }
        #endregion

        #region 写入数据
        /// <summary>
        /// 插入数据到表
        /// </summary>
        /// <param name="table">表名</param>
        /// <param name="columns">涉及到的列</param>
        /// <param name="data">二维数组，内层数组为对应行</param>
        /// <returns>插入的行数，未成功则返回0</returns>
        public int Insert(string table, string[] columns, params string[][] data)
        {
            if (data == null || data.Length == 0 || data[0].Length != columns.Length) return 0;
            StringBuilder builder = new StringBuilder(data.Length * 64);
            builder.Append($"INSERT INTO {table} (");
            builder.Append(string.Join(",", columns));
            builder.Append(") VALUES ");
            int end = data.Length;
            foreach (string[] row in data)
            {
                builder.Append("('");
                builder.Append(string.Join("','", row));
                builder.Append(end == 1 ? "')" : "'),");
                end--;
            }
            cmd = new SQLiteCommand(builder.ToString(), conn);
            return cmd.ExecuteNonQuery();
        }
        #endregion

        #region 修改数据
        public int UpDate(string table, Dictionary<string, string> set, string where)
        {
            if (set == null || set.Count == 0) return 0;
            if (string.IsNullOrEmpty(where)) return 0;
            StringBuilder builder = new StringBuilder(128);
            builder.Append($"UPDATE {table} SET ");
            var setPairs = set.Select(kvp => $"{kvp.Key}='{kvp.Value}'");
            builder.Append(string.Join(",", setPairs));
            builder.Append($" WHERE {where}");
            string sql = builder.ToString();
            //System.Diagnostics.Debug.WriteLine(sql);
            cmd = new SQLiteCommand(sql, conn);
            return cmd.ExecuteNonQuery();
        }
        public int UpDate(string table, string set, string where)
        {
            if (set == null ) return 0;
            if (string.IsNullOrEmpty(where)) return 0;
            string sql = $"UPDATE {table} SET {set}  WHERE {where}";
            //System.Diagnostics.Debug.WriteLine(sql);
            cmd = new SQLiteCommand(sql, conn);
            return cmd.ExecuteNonQuery();
        }

        #endregion

        #region 删除数据
        public int Remove(string table, params string[] ids)
        {
            if (ids.Length == 0) return 0;
            string keys = string.Join("','", ids);
            cmd = new SQLiteCommand($"DELETE FROM {table} WHERE id IN ('{keys}')", conn);
            return cmd.ExecuteNonQuery();
        }

        public int Remove(string table, string where)
        {
            if (where == "") return 0;
            string sqlCode = $"DELETE FROM {table} WHERE {where}";
            cmd = new SQLiteCommand(sqlCode, conn);
            return cmd.ExecuteNonQuery();
        }
        #endregion

        #region 转换数据
        public string DataToJson(object data)
        {
            JsonSerializerOptions options = new JsonSerializerOptions
            {
                WriteIndented = true,
                PropertyNamingPolicy = null,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            return JsonSerializer.Serialize(data, options);
        }
        #endregion
    }
}