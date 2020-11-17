using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace PersonInfoWebAPIWPF.Models
{
    public class DataAccessLayer
    {
        SqlConnection sqlConnection;

        public DataAccessLayer()
        {
            var configuration = GetConfiguration();
            sqlConnection = new SqlConnection(configuration.GetSection("ConnectionStrings").GetSection("dbConnection").Value);
        }

        IConfigurationRoot GetConfiguration()
        {
            var builder = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true);
            return builder.Build();
        }

        public DataTable GetData(string pSp, string pParm)
        {
            using (SqlConnection con = sqlConnection)
            {
                DataSet ds = new DataSet();
                try
                {
                    string parm = pParm.Replace("&", "&amp;");
                    SqlParameter param;
                    SqlDataAdapter adapter;

                    con.Open();
                    SqlCommand command = new SqlCommand(pSp, con);
                    command.CommandType = CommandType.StoredProcedure;

                    if (parm.Length != 0)
                    {
                        param = new SqlParameter("@xmlparm", parm);
                        param.Direction = ParameterDirection.Input;
                        param.DbType = DbType.String;
                        command.Parameters.Add(param);
                    }
                    adapter = new SqlDataAdapter(command);
                    adapter.Fill(ds, pSp);
                }
                catch(Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    con.Close();
                    con.Dispose();
                }

                return ds.Tables[0];
            }
        }

        public void SetData(string pSp, string pParm)
        {
            using (SqlConnection con = sqlConnection)
            {
                DataSet ds = new DataSet();
                try
                {
                    string parm = pParm.Replace("&", "&amp;");
                    string sp = pSp;
                    SqlParameter param;

                    con.Open();
                    SqlCommand command = new SqlCommand(sp, con);
                    command.CommandType = CommandType.StoredProcedure;

                    if (parm.Length != 0)
                    {
                        param = new SqlParameter("@xmlparm", parm);
                        param.Direction = ParameterDirection.Input;
                        param.DbType = DbType.String;
                        command.Parameters.Add(param);
                    }
                    command.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }

        public int SetDataWithReturn(string pSp, string pParm)
        {
            using (SqlConnection con = sqlConnection)
            {
                DataSet ds = new DataSet();
                try
                {
                    string parm = pParm.Replace("&", "&amp;");
                    string sp = pSp;
                    SqlParameter param;

                    con.Open();
                    SqlCommand command = new SqlCommand(pSp, con);
                    command.CommandType = CommandType.StoredProcedure;

                    if (parm.Length != 0)
                    {
                        param = new SqlParameter("@xmlparm", parm);
                        param.Direction = ParameterDirection.Input;
                        param.DbType = DbType.String;
                        command.Parameters.Add(param);
                    }
                    int retval = Convert.ToInt32(command.ExecuteScalar());
                    return retval;
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    con.Close();
                    con.Dispose();
                }
            }
        }
    }
}
