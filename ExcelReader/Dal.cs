using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelReader
{
    public class Dal
    {
        private static SqlConnection getConnection()
        {
            var myCon = new SqlConnection();
            myCon.ConnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;

            return myCon;
        }
        public static int ExecuteNonQuery(SqlCommand cmd)
        {
            var sql = cmd.CommandText;
            cmd.Connection = getConnection();
            int res;
            try
            {
                cmd.Connection.Open();
                cmd.CommandText = sql;
                res = cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Connection.Close();
            }
            return res;
        }
        public static string getString(SqlCommand ObjCmd, string ColumnName)
        {
            DataTable dt = getDataTable(ObjCmd);
            return (string)dt.Rows[0][ColumnName];
        }
        public static List<int> getIntList(SqlCommand ObjCmd, string ColumnName)
        {
            DataTable dt = getDataTable(ObjCmd);
            List<int> myList = new List<int>();
            foreach (DataRow dr in dt.Rows)
            {
                myList.Add((int)dr[ColumnName]) ;
            }
            return myList;
        }
        public static Dictionary<int, string> getIntDictionary(SqlCommand ObjCmd, string ColumnName)
        {
            DataTable dt = getDataTable(ObjCmd);
            Dictionary<int, string> myDictionary = new Dictionary<int, string>();
            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    myDictionary.Add((int)dr[ColumnName], (string)dr[1]);
                }
                catch (Exception ex)
                {}

            }
            return myDictionary;
        }


        public static DataTable getDataTable(SqlCommand cmd)
        {
            cmd.Connection = getConnection();
            var objDt = new DataTable();
            try
            {
                cmd.Connection.Open();
                SqlDataAdapter objDa = new SqlDataAdapter(cmd);
                objDa.Fill(objDt);
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Connection.Close();
            }
            return objDt;
        }
        public virtual string ExecuteScalar(SqlCommand cmd)
        {
            string res = "";
            string sql = cmd.CommandText;
            cmd.Connection = getConnection();
            try
            {
                cmd.Connection.Open();
                cmd.CommandText = sql;
                object _res = cmd.ExecuteScalar();
                if (_res != null)
                {
                    res = _res.ToString();
                }
                else
                {
                    res = null;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cmd.Connection.Close();
            }
            return res;
        }
    }
}
