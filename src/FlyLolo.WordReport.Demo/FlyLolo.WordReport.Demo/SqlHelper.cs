using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FlyLolo.WordReport.Demo
{
    public static class SqlHelper
    {
        public static string GetConnectionString(string connName)
        {
            return System.Configuration.ConfigurationManager.ConnectionStrings[connName].ToString();
        }
        public static DataSet ExecuteDataSet(string connName, string procName, params IDataParameter[] procParams)
        {
            SqlConnection conn = null;
            DataSet ds = new DataSet();
            SqlDataAdapter da = new SqlDataAdapter();
            SqlCommand cmd = null;

            try
            {
                //Setup command object
                cmd = new SqlCommand(procName);
                cmd.CommandType = CommandType.Text;
                if (procParams != null)
                {
                    for (int index = 0; index < procParams.Length; index++)
                    {
                        if (procParams[index] != null)
                        {
                            cmd.Parameters.Add(procParams[index]);
                        }
                        else
                        {
                            break;
                        }

                    }
                }
                da.SelectCommand = (SqlCommand)cmd;
                string connstr = GetConnectionString(connName);
                conn = new SqlConnection(connstr);
                cmd.Connection = conn;
                conn.Open();

                //Fill the dataset
                da.Fill(ds);
            }
            catch
            {
                throw;
            }
            finally
            {
                if (da != null) da.Dispose();
                if (cmd != null) cmd.Dispose();
                conn.Dispose(); //Implicitly calls cnx.Close()
            }
            return ds;
        }
    }
}
