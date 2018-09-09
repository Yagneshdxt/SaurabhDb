using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;

namespace DataAccess
{
    public class DBAccess
    {
        public DBAccess(string ConnectionString) {
            if (String.IsNullOrEmpty(ConnectionString)) throw new Exception("Connection String Not found");
            this.ConnString = ConnectionString;
            SDbconObj = new SqlConnection(ConnectionString);
        }

        protected DBAccess() { 
        
        }

        private string ConnString { get; set; }

        private SqlConnection SDbconObj;// = new SqlConnection(ConfigurationManager.ConnectionStrings["saurabhDbcon"].ConnectionString);

        public SqlConnection conObj
        {
            get
            {
                if (SDbconObj == null)
                {
                    SDbconObj = new SqlConnection(ConnString);
                }
                return SDbconObj;
            }
        }

        public void openCon(SqlConnection con)
        {
            if (con.State == System.Data.ConnectionState.Closed)
                con.Open();
        }

        public void CloseCon(SqlConnection con)
        {
            if (con.State == System.Data.ConnectionState.Open)
                con.Close();
        }

        //Get Data set
        public DataSet getDataSet(string sqlQuery, SqlParameter[] sqlParms, CommandType cmTyp = CommandType.Text)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = conObj.CreateCommand();
                cmd.CommandType = cmTyp;
                cmd.CommandText = sqlQuery;
                if (sqlParms != null)
                    cmd.Parameters.AddRange(sqlParms.ToArray());
                SqlDataAdapter dA = new SqlDataAdapter(cmd);
                openCon(conObj);
                dA.Fill(ds);
            }
            catch (Exception)
            {
                ds = null;
                CloseCon(conObj);
                throw;
            }
            finally
            {
                CloseCon(conObj);
            }
            return ds;
        }

        //Get Data set with output parameter.
        public DataSet GetDataSetWithOutParameter(string sqlQuery, SqlParameter[] sqlParms, ref Dictionary<String, String> OparDic, CommandType cmTyp = CommandType.Text)
        {
            DataSet ds = new DataSet();
            try
            {
                SqlCommand cmd = conObj.CreateCommand();
                cmd.CommandType = cmTyp;
                cmd.CommandText = sqlQuery;
                if (sqlParms != null)
                {
                    cmd.Parameters.AddRange(sqlParms.ToArray());
                    foreach (var item in OparDic)
                    {
                        cmd.Parameters[item.Key].Direction = ParameterDirection.Output;
                    }
                }
                openCon(conObj);
                SqlDataAdapter dA = new SqlDataAdapter(cmd);
                dA.Fill(ds);
                List<String> keyColl = new List<string>(OparDic.Keys);
                if (sqlParms != null)
                {
                    foreach (var item in keyColl)
                    {
                        OparDic[item] = cmd.Parameters[item].Value.ToString();
                    }
                }

            }
            catch (Exception)
            {
                ds = null;
                CloseCon(conObj);
                throw;
            }
            finally
            {
                CloseCon(conObj);
            }
            return ds;
        }

        //Execute non query 
        public int ExecNonQuery(string SqlQuery, SqlParameter[] sqlParms, CommandType cmdTyp = CommandType.Text, SqlTransaction Sqltrn = null)
        {
            int retNo = 0;
            try
            {
                SqlCommand cmd = conObj.CreateCommand();
                cmd.CommandType = cmdTyp;
                cmd.CommandText = SqlQuery;
                if (sqlParms != null)
                    cmd.Parameters.AddRange(sqlParms.ToArray());
                if (Sqltrn == null)
                    openCon(conObj);
                else
                    cmd.Transaction = Sqltrn;
                retNo = cmd.ExecuteNonQuery();
            }
            catch (Exception)
            {
                retNo = 0;
                if (Sqltrn == null)
                    CloseCon(conObj);
                throw;
            }
            finally
            {
                if (Sqltrn == null)
                    CloseCon(conObj);
            }
            return retNo;
        }

        //Execute non query with out put parameters
        public int ExecNonQueryWithOutParameter(string SqlQuery, SqlParameter[] sqlParms, ref Dictionary<String, String> OparDic, CommandType cmdTyp = CommandType.Text, SqlTransaction Sqltrn = null)
        {
            int retNo = 0;
            try
            {
                SqlCommand cmd = conObj.CreateCommand();
                cmd.CommandType = cmdTyp;
                cmd.CommandText = SqlQuery;
                if (sqlParms != null)
                {
                    cmd.Parameters.AddRange(sqlParms.ToArray());
                    foreach (var item in OparDic)
                    {
                        cmd.Parameters[item.Key].Direction = ParameterDirection.Output;
                    }
                }
                if (Sqltrn == null)
                    openCon(conObj);
                else
                    cmd.Transaction = Sqltrn;
                retNo = cmd.ExecuteNonQuery();
                List<String> keyColl = new List<string>(OparDic.Keys);
                if (sqlParms != null)
                {
                    foreach (var item in keyColl)
                    {
                        OparDic[item] = cmd.Parameters[item].Value.ToString();
                    }
                }
            }
            catch (Exception)
            {
                retNo = 0;
                if (Sqltrn == null)
                    CloseCon(conObj);
                throw;
            }
            finally
            {
                if (Sqltrn == null)
                    CloseCon(conObj);
            }
            return retNo;
        }

        //Execute scalar 
        public object SelectScalar(string sqlQuery, SqlParameter[] sqlParms = null, CommandType cmdTyp = CommandType.Text)
        {
            object chkint = null;
            try
            {
                SqlCommand cmd = conObj.CreateCommand();
                cmd.CommandType = cmdTyp;
                cmd.CommandText = sqlQuery;
                if (sqlParms != null)
                    cmd.Parameters.AddRange(sqlParms.ToArray());
                openCon(conObj);
                chkint = cmd.ExecuteScalar();
            }
            catch (Exception)
            {
                chkint = null;
                CloseCon(conObj);
                throw;
            }
            finally
            {
                CloseCon(conObj);
            }
            return chkint;
        }

        //Execute scalar with out put parameter
        public object SelectScalarWithOutputParameter(string sqlQuery, ref Dictionary<String, String> OparDic, SqlParameter[] sqlParms = null, CommandType cmdTyp = CommandType.Text)
        {
            object chkint = null;
            try
            {
                SqlCommand cmd = conObj.CreateCommand();
                cmd.CommandType = cmdTyp;
                cmd.CommandText = sqlQuery;
                if (sqlParms != null)
                {
                    cmd.Parameters.AddRange(sqlParms.ToArray());
                    foreach (var item in OparDic)
                    {
                        cmd.Parameters[item.Key].Direction = ParameterDirection.Output;
                    }
                }
                openCon(conObj);
                chkint = cmd.ExecuteScalar();
                List<String> keyColl = new List<string>(OparDic.Keys);
                if (sqlParms != null)
                {
                    foreach (var item in keyColl)
                    {
                        OparDic[item] = cmd.Parameters[item].Value.ToString();
                    }
                }
            }
            catch (Exception)
            {
                chkint = null;
                CloseCon(conObj);
                throw;
            }
            finally
            {
                CloseCon(conObj);
            }
            return chkint;
        }
    }
}
