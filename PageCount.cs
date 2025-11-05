using Org.BouncyCastle.Utilities.Net;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Web;

/// <summary>
/// 統計各頁面的使用率
/// </summary>
public class PageCount
{
    public PageCount()
    {
        
    }


    protected static string GetIPAddress()
    {
        string ip = "";
        if (System.Web.HttpContext.Current.Request.ServerVariables["HTTP_VIA"] == null)
        {
            ip = System.Web.HttpContext.Current.Request.ServerVariables["REMOTE_ADDR"].ToString();
        }
        else
        {
            ip = System.Web.HttpContext.Current.Request.ServerVariables["HTTP_X_FORWARDED_FOR"].ToString();
        }
        ip = ip.Replace("::1", "127.0.0.1");
        return ip;
    }

    protected static string GetMachineName(string IPAddresses)
    {
        string MachineName = "";
        try
        {
            if (IPAddresses!=null)
            {
                if (IPAddresses == "127.0.0.1" || IPAddresses == "localhost")
                {
                    // 使用本地計算機名稱
                    MachineName = Environment.MachineName;
                }
                else
                {
                    // 取得主機名稱
                    IPHostEntry hostEntry = Dns.GetHostEntry(IPAddresses);
                    MachineName = hostEntry.HostName;
                }
            }
        }
        catch (Exception)
        {
        }
        return MachineName;
    }


    public static bool Log(System.Web.UI.Page page, string EmpNo, ref System.Text.StringBuilder sbMsg)
    {
        bool isSuccess = true;
        

        /*統計程式使用次數*/
        string PageNo = System.IO.Path.GetFileName(page.Request.PhysicalPath);

        string IPAddresses = GetIPAddress();
        string hostName = GetMachineName(IPAddresses); 
 
        System.Collections.ArrayList alPageNo = new System.Collections.ArrayList();
        ERP.DBDao dao = new ERP.DBDao();

        string sqlQuery = @"
            SELECT FilterPageNo 
            FROM dbo.PageFilter WITH(NOLOCK)
            WHERE FilterPageNo = @FilterPageNo";
        dao.SqlCommand.Parameters.AddWithValue("@FilterPageNo", PageNo);
        System.Data.DataTable dtFilter = dao.SqlSelect(sqlQuery);

        bool isContinue = (dtFilter.Rows.Count == 0);
        if (isContinue)
        {
            dao.SqlConn.Open();
            try
            {
                if (page.Session["LoginPageNo"] != null)
                {
                    alPageNo = (System.Collections.ArrayList)page.Session["LoginPageNo"];
                    if (alPageNo.IndexOf(PageNo) > -1)
                        isContinue = false;
                }

                if (isContinue)
                {
                    string RealUser = EmpNo;
                    if (page.Session["FirstLoginUser"] != null)
                        RealUser = page.Session["FirstLoginUser"].ToString();

                    string sqlInsert = @"
                    INSERT INTO dbo.PageCount (Date, PageNo, CreateDate, CreateUser, RealCreateUser, MachineName, IPAddresses)
                    VALUES (@Date, @PageNo, GETDATE(), @User, @RealUser, @MachineName, @IPAddresses)";
                    dao.SqlCommand.Parameters.Clear();
                    dao.SqlCommand.Parameters.AddWithValue("@Date", DateTime.Now.ToString("yyyy/MM/dd"));
                    dao.SqlCommand.Parameters.AddWithValue("@PageNo", PageNo);
                    dao.SqlCommand.Parameters.AddWithValue("@MachineName", hostName);
                    dao.SqlCommand.Parameters.AddWithValue("@IPAddresses", IPAddresses);
                    dao.SqlCommand.Parameters.AddWithValue("@User", EmpNo);
                    dao.SqlCommand.Parameters.AddWithValue("@RealUser", RealUser);
                    isSuccess = dao.SqlInsert(sqlInsert, ref sbMsg);

                    alPageNo.Add(PageNo);
                    page.Session["LoginPageNo"] = alPageNo;
                }
            }
            catch (Exception ex)
            {
                isSuccess = false;
                sbMsg.Append(ex.Message);
            }
            finally
            {
                dao.SqlConn.Close();
            }
        }
        return isSuccess;
    }

    public static void InitPageSession(System.Web.UI.Page page)
    {
        page.Session["LoginPageNo"] = null;
    }
}