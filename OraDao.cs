using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using Oracle.ManagedDataAccess.Client;


public class OraDao
{
    //Log4net
    protected static readonly log4net.ILog mLog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

    string oracleConnStr;

    public Oracle.ManagedDataAccess.Client.OracleConnection oracleConn = new Oracle.ManagedDataAccess.Client.OracleConnection();
    public Oracle.ManagedDataAccess.Client.OracleCommand oracleCommand = new Oracle.ManagedDataAccess.Client.OracleCommand();
    public Oracle.ManagedDataAccess.Client.OracleTransaction oracleTrans;

    public OraDao(string ConnStr)
    {
        string strConn = System.Configuration.ConfigurationManager.ConnectionStrings[ConnStr].ConnectionString;
        oracleConn.ConnectionString = strConn;
        //Select使用
        oracleConnStr = strConn;
    }

    public OraDao()
    {
        string strConn = System.Configuration.ConfigurationManager.ConnectionStrings["T100DBConn"].ConnectionString;
        oracleConn.ConnectionString = strConn;
        //Select使用
        oracleConnStr = strConn;
    }

    public void BeginTransaction()
    {
        oracleTrans = oracleConn.BeginTransaction();
        oracleCommand.Transaction = oracleTrans;
    }

    public void Commit()
    {
        oracleTrans.Commit();
    }

    public void Rollback()
    {
        oracleTrans.Rollback();
    }

    public System.Data.DataTable OracelSelect(string OracelQuery)
    {
        System.Data.DataTable dtDBData = new System.Data.DataTable();
        using (Oracle.ManagedDataAccess.Client.OracleConnection oracleSelectConn = new Oracle.ManagedDataAccess.Client.OracleConnection(oracleConnStr))
        {
            oracleSelectConn.Open();
            try
            {
                oracleCommand.CommandText = OracelQuery;
                oracleCommand.Connection = oracleSelectConn;
                using (Oracle.ManagedDataAccess.Client.OracleDataReader reader = oracleCommand.ExecuteReader())
                {
                    dtDBData.Load(reader);
                }
            }
            catch (Exception ex)
            {
                mLog.Error(ex.Message);
            }
            finally
            {
                oracleSelectConn.Close();
            }
        }
        return dtDBData;
    }

    public string OracleSelectToString(string OracelQuery)
    {
        string QueryStr = string.Empty;
        using (Oracle.ManagedDataAccess.Client.OracleConnection oracleSelectConn = new Oracle.ManagedDataAccess.Client.OracleConnection(oracleConnStr))
        {
            oracleSelectConn.Open();
            try
            {
                Oracle.ManagedDataAccess.Client.OracleCommand oracelCommand = new Oracle.ManagedDataAccess.Client.OracleCommand(OracelQuery, oracleSelectConn);
                using (Oracle.ManagedDataAccess.Client.OracleDataReader reader = oracelCommand.ExecuteReader())
                {
                    System.Data.DataTable dtDBData = new System.Data.DataTable();
                    dtDBData.Load(reader);
                    if (dtDBData != null && dtDBData.Rows.Count == 1)
                    {
                        QueryStr = dtDBData.Rows[0][0].ToString();
                    }
                }
            }
            catch (Exception ex)
            {
                mLog.Error(ex.Message);
            }
            finally
            {
                oracleSelectConn.Close();
            }
        }
        return QueryStr;
    }

    public bool OracleInsert(string sqlInsert)
    {
        bool isSuccess = false;
        try
        {
            oracleCommand.CommandText = sqlInsert;
            oracleCommand.Connection = oracleConn;
            oracleCommand.ExecuteNonQuery();
            isSuccess = true;
        }
        catch (Exception ex)
        {
            mLog.Error(ex.Message);
        }
        return isSuccess;
    }

    public bool OracleUpdate(string sqlUpdate)
    {
        bool isSuccess = false;
        try
        {
            oracleCommand.CommandText = sqlUpdate;
            oracleCommand.Connection = oracleConn;
            oracleCommand.ExecuteNonQuery();
            isSuccess = true;
        }
        catch (Exception ex)
        {
            mLog.Error(ex.Message);
        }
        return isSuccess;
    }

    public bool OracleUpdate(string sqlUpdate, ref System.Text.StringBuilder sbMsg)
    {
        bool isSuccess = false;
        try
        {
            oracleCommand.CommandText = sqlUpdate;
            oracleCommand.Connection = oracleConn;
            oracleCommand.ExecuteNonQuery();
            isSuccess = true;
        }
        catch (Exception ex)
        {
            sbMsg.Append(ex.Message + "\\n");
            mLog.Error(ex.Message);
        }
        return isSuccess;
    }

    public System.Data.DataTable getERPEmpInfo(string loginName)
    {
        string sql = string.Format("SELECT E.NAM_EMP AS UserName, E.PS_6 AS LoginID, E.COD_EMP AS ID, C.CONTENT AS Title, C1.CONTENT AS Dept, E.COD_DEPT AS DeptID FROM TE.EEMPB E LEFT JOIN TE.CODD C ON C.CODE_ID = 'CODPOS' AND C.CODE = E.COD_POS LEFT JOIN TE.CODD C1 ON C1.CODE_ID = 'CODDPT' AND C1.CODE = E.COD_DEPT WHERE E.DAT_OUT IS NULL AND E.PS_6 IS NOT NULL AND UPPER(E.PS_6) = '{0}'", loginName.ToUpper());

        System.Data.DataTable dtLoginInfo = OracelSelect(sql);
        return dtLoginInfo;
    }

}
