using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace BandR
{
    public class SQLHelper
    {

        /// <summary>
        /// </summary>
        public static bool ExecuteCmd(string dbConnStr, string sql, out string msg)
        {
            msg = "";

            try
            {
                using (var conn = new SqlConnection(dbConnStr))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        int rowCount = cmd.ExecuteNonQuery();
                    }
                }

            }
            catch (Exception ex)
            {
                msg = ex.ToString();
            }

            return msg == "";
        }

        /// <summary>
        /// </summary>
        public static bool ExecuteQueryDt(string dbConnStr, string sql, out DataTable dt, out string msg)
        {
            dt = new DataTable();
            msg = "";

            try
            {
                using (var conn = new SqlConnection(dbConnStr))
                {
                    conn.Open();
                    using (var da = new SqlDataAdapter(sql, conn))
                    {
                        da.Fill(dt);
                    }
                }

            }
            catch (Exception ex)
            {
                dt = new DataTable();
                msg = ex.ToString();
            }

            return msg == "";
        }

        /// <summary>
        /// </summary>
        public static bool ExecuteQueryList(string dbConnStr, string sql, out List<string> lst, out string msg)
        {
            lst = new List<string>();
            msg = "";

            try
            {
                using (var conn = new SqlConnection(dbConnStr))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        using (var rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                lst.Add(GenUtil.SafeTrim(rdr[0]));
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                lst = new List<string>();
                msg = ex.ToString();
            }

            return msg == "";
        }

        /// <summary>
        /// </summary>
        public static bool ExecuteQueryDict(string dbConnStr, string sql, out Dictionary<string, string> dict, out string msg)
        {
            dict = new Dictionary<string, string>();
            msg = "";

            try
            {
                using (var conn = new SqlConnection(dbConnStr))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        using (var rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                dict.Add(GenUtil.SafeTrim(rdr[0]), GenUtil.SafeTrim(rdr[1]));
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                dict = new Dictionary<string, string>();
                msg = ex.ToString();
            }

            return msg == "";
        }

        /// <summary>
        /// </summary>
        public static bool ExecuteQueryObj(string dbConnStr, string sql, out string obj, out string msg)
        {
            obj = "";
            msg = "";

            try
            {
                using (var conn = new SqlConnection(dbConnStr))
                {
                    conn.Open();
                    using (var cmd = new SqlCommand(sql, conn))
                    {
                        using (var rdr = cmd.ExecuteReader())
                        {
                            while (rdr.Read())
                            {
                                obj = GenUtil.SafeTrim(rdr[0]);
                                break;
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                obj = "";
                msg = ex.ToString();
            }

            return msg == "";
        }

        /// <summary>
        /// </summary>
        public static string MakeSafe(object o)
        {
            return GenUtil.SafeTrim(o).Replace("'", "''");
        }

    }
}
