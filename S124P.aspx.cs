using CoreExtension.Linq;
using NPOI.SS.UserModel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using System;
using System.Activities.Expressions;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO; //檔案控制
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Threading;
using System.Web.UI;
using System.Web.UI.HtmlControls;
using System.Web.UI.WebControls;

public partial class S124P : System.Web.UI.Page
{
    protected override void OnPreInit(EventArgs e)
    {
        //判斷是否由EEP簽核頁面開啟ERP頁面
        if (Request.QueryString["EEP"] != null && Request.QueryString["EEP"].ToString() == "Y")
        {
            //不需Banner
            Page.MasterPageFile = "~/MasterPage_Blank.master";
        }
        else
        {
            //Banner
            Page.MasterPageFile = "~/MasterPage.master";
        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (this.Session["UserEmployeeNo"] != null)
        {
            if (!IsPostBack)
            {
                GetAuh();
                BindLsb(lsbSalesID);
                BindLsb(lsbObject);

                txtStartDate.Text = DateTime.Now.ToString("yyyy-MM") + "-01";
                txtEndDate.Text = DateTime.Now.ToString("yyyy-MM-dd");

                ddlYear.SelectedValue = DateTime.Now.Year.ToString();
            }
        }
        else
            this.Response.Redirect("~/Default.aspx?SSOKey=&RedirectUrl=S1/S124P.aspx");
    }

    protected void BindLsb(ListBox lsb)
    {
        ERP.DBDao daoSelect = new ERP.DBDao();
        string sql = "";

        switch (lsb.ID)
        {
            case "lsbSalesID":
                sql = @"
SELECT EmpNo AS Value, EmpNo + '/' + EmpName AS Text
FROM VW_SA_BelongSection_ByEmp
WHERE Remark = 'SalesID' ";
                break;

            case "lsbObject":
                sql = @"
SELECT ObjectNo AS Value, Object AS Text 
FROM OBJECT (NOLOCK)
WHERE CustomerID = 1";
                break;
        }

        lsb.Items.Clear();
        lsb.DataValueField = "Value";
        lsb.DataTextField = "Text";
        lsb.DataSource = daoSelect.SqlSelect(sql);
        lsb.DataBind();

        // 以下為在Bind後加入第一筆資料、其他選項
        lsb.Items.Insert(0, new ListItem("請選擇", ""));
        lsb.SelectedIndex = 0;

    }

    protected void GetAuh()
    {
        ERP.DBDao daoSelect = new ERP.DBDao();

        string sql = @"
SELECT '全部' AS DeptNo, '全部' AS DeptName
UNION ALL
--SELECT DISTINCT DeptNo, DeptName FROM DeptReportLine WHERE (ParentDeptName = '業務處' AND OrgLevel = 6) OR DeptName = '大陸事業部'
SELECT DISTINCT DeptNo, DeptName FROM DeptReportLine 
WHERE DeptNo IN (SELECT DISTINCT DeptNo FROM VW_SA_BelongSection_ByEmp)";
        DataTable dt = daoSelect.SqlSelect(sql);
        rblBelongSection.DataSource = dt;
        rblBelongSection.DataTextField = "DeptName";
        rblBelongSection.DataValueField = "DeptNo";
        rblBelongSection.DataBind();

        sql = @"SELECT DeptNo, DeptName FROM VW_SA_BelongSection_ByEmp WHERE EmpNo = @User";
        daoSelect.SqlCommand.Parameters.AddWithValue("@User", Session["UserEmployeeNo"].ToString());
        dt = daoSelect.SqlSelect(sql);
        if (dt.Rows.Count == 0)
        {
            JavaScriptFunction.ShowJsMSG(this, "系統沒有您的查詢權限");
        }
        else
        {
            bool isPowerAdmin = (dt.Select(" DeptName = '最大權限名單' ").Length > 0) ? true : false;
            if (isPowerAdmin || Session["UserEmployeeNo"].ToString() == "140310")
                rblBelongSection.Items.FindByValue("全部").Selected = true;
            else
            {
                foreach (ListItem list in rblBelongSection.Items)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        list.Enabled = false;
                        if (list.Value == dr["DeptNo"].ToString())
                        {
                            list.Enabled = true;
                            break;
                        }
                    }
                }
            }
        }
    }

    protected void btnClear_Click(object sender, EventArgs e)
    {
        txtStartDate.Text = "";
        txtEndDate.Text = "";
        lsbObject.SelectedValue = "";
        lsbSalesID.SelectedValue = "";
    }

    protected void btnQuery_Click(object sender, EventArgs e)
    {
        string strAlertMsg = "", sql = "", sqlWhr = "";

        if (string.IsNullOrWhiteSpace(txtStartDate.Text) || string.IsNullOrWhiteSpace(txtEndDate.Text))
            strAlertMsg += "起訖區間不可空白";

        if (string.IsNullOrWhiteSpace(rblBelongSection.SelectedValue))
            strAlertMsg += "事業部 不可空白";

        if (!string.IsNullOrWhiteSpace(strAlertMsg))
            JavaScriptFunction.ShowJsMSG(this, strAlertMsg);
        else
        {
            ERP.DBDao daoSelect = new ERP.DBDao();

            #region Parameters設置
            sqlWhr += (string.IsNullOrWhiteSpace(ddlYear.SelectedValue)) ? "" : " AND 報價日期 LIKE @Year";
            sqlWhr += (string.IsNullOrWhiteSpace(lsbObject.SelectedValue)) ? "" : " AND 客戶代號 = @ObjectNo ";
            sqlWhr += (string.IsNullOrWhiteSpace(lsbSalesID.SelectedValue)) ? "" : " AND 業務工號 = @SalesID ";

            if (rblCostSiteName.SelectedValue != "全部")
                sqlWhr += (string.IsNullOrWhiteSpace(rblCostSiteName.SelectedValue)) ? "" : " AND 成本據點 = @CostSiteName ";

            if (rblBelongSection.SelectedValue != "全部")
                sqlWhr += (string.IsNullOrWhiteSpace(rblBelongSection.SelectedValue)) ? "" : " AND 事業部代號 = @BelongSection ";

            string strSortTypeA = "";
            foreach (ListItem list in cblSortTypeA.Items)
            {
                if (string.IsNullOrWhiteSpace(strSortTypeA))
                    strSortTypeA += (list.Selected) ? string.Format("'{0}'", list.Text) : "";
                else
                    strSortTypeA += (list.Selected) ? string.Format(",'{0}'", list.Text) : "";
            }
            sqlWhr += (string.IsNullOrWhiteSpace(strSortTypeA)) ? "" : string.Format(" AND 產品屬性 IN ({0}) ", strSortTypeA);

            daoSelect.SqlCommand.Parameters.Clear();
            daoSelect.SqlCommand.Parameters.AddWithValue("@Year", ddlYear.SelectedValue + "%");
            daoSelect.SqlCommand.Parameters.AddWithValue("@ObjectNo", lsbObject.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@SalesID", lsbSalesID.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@BelongSection", rblBelongSection.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@CostSiteName", rblCostSiteName.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@1M", ddlYear.SelectedValue + "/01");
            daoSelect.SqlCommand.Parameters.AddWithValue("@2M", ddlYear.SelectedValue + "/02");
            daoSelect.SqlCommand.Parameters.AddWithValue("@3M", ddlYear.SelectedValue + "/03");
            daoSelect.SqlCommand.Parameters.AddWithValue("@4M", ddlYear.SelectedValue + "/04");
            daoSelect.SqlCommand.Parameters.AddWithValue("@5M", ddlYear.SelectedValue + "/05");
            daoSelect.SqlCommand.Parameters.AddWithValue("@6M", ddlYear.SelectedValue + "/06");
            daoSelect.SqlCommand.Parameters.AddWithValue("@7M", ddlYear.SelectedValue + "/07");
            daoSelect.SqlCommand.Parameters.AddWithValue("@8M", ddlYear.SelectedValue + "/08");
            daoSelect.SqlCommand.Parameters.AddWithValue("@9M", ddlYear.SelectedValue + "/09");
            daoSelect.SqlCommand.Parameters.AddWithValue("@10M", ddlYear.SelectedValue + "/10");
            daoSelect.SqlCommand.Parameters.AddWithValue("@11M", ddlYear.SelectedValue + "/11");
            daoSelect.SqlCommand.Parameters.AddWithValue("@12M", ddlYear.SelectedValue + "/12");
            #endregion

            #region 取得 DetailListData (匯出總明細 使用)
            sql = @"
WITH S124P AS
(
	SELECT * FROM RPS_S124P (NOLOCK)
	WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' --AND 短品名 = 'TMRF62X10_TFM'
	{0}
    --AND 報價日期 LIKE '2023%'
)

SELECT  計入件數, 計入金額, 狀態, 業務, 客戶簡稱, 產品屬性, 短品名, SUM(數量) AS 數量, FORMAT(SUM(總價), 'N0') AS 總價, 報價失敗原因, 狀態備註紀錄, 報價單號, 報價月份, 內部訂單, 訂單月份, 事業部
FROM 
(
		SELECT  計入件數, 計入金額, 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 數量, 總價, 報價失敗原因, 狀態備註紀錄, 
				(
					SELECT (
						SELECT  CAST(報價單號 AS NVARCHAR) + ';'
						FROM S124P R0 (NOLOCK)
						WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
						FOR XML PATH('')
					)
				) AS 報價單號, 
				(
					SELECT TOP(1) SUBSTRING(報價日期, 1, 7)
					FROM S124P R0 (NOLOCK)
					WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
					ORDER BY 報價單號
				) AS 報價月份,
				(
					SELECT (
						SELECT  CAST(內部訂單 AS NVARCHAR) + ';'
						FROM S124P R0 (NOLOCK)
						WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
						FOR XML PATH('')
					)
				) AS 內部訂單, 訂單月份, 
				事業部代號, 事業部, 成本據點  
		FROM S124P R (NOLOCK) 
		WHERE R.產品屬性 = '新設計案'
		UNION ALL
		SELECT  計入件數, 計入金額, 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 數量, 總價, 報價失敗原因, 狀態備註紀錄, 
		       報價單號, 
			   SUBSTRING(報價日期, 1, 7) AS 報價月份, 
		       內部訂單,訂單月份,
			   事業部代號, 事業部, 成本據點 
		FROM S124P R (NOLOCK) 
		WHERE R.產品屬性 != '新設計案' 
) t
GROUP BY  計入件數, 計入金額, 狀態, 業務, 客戶簡稱, 產品屬性, 短品名, 報價失敗原因, 狀態備註紀錄, 報價單號, 報價月份, 內部訂單, 訂單月份, 事業部 
ORDER BY 業務, 短品名, 計入件數, 狀態, 數量, 總價
";
            sql = string.Format(sql, sqlWhr);

            DataTable dtDetail = daoSelect.SqlSelect(sql);
            ViewState["S124P_DetailListData"] = dtDetail;
            #endregion

            #region 取得 TotalListData  (為了取短品名使用)
            sql = @"
  SELECT 事業部, R.客戶簡稱 , 狀態, SUBSTRING(報價日期, 1, 7) AS YM, 
  (R.短品名 +  IIF(狀態 = 'Success', '【Success】', '')) AS 短品名
  FROM RPS_S124P R (NOLOCK)
  WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' AND 計入件數 = 'Y' --AND 短品名 = 'TMMP08X8_WSRF_MLO'
  {0}
  ORDER BY 事業部, SUBSTRING(報價日期, 1, 7), 客戶簡稱, IIF(狀態 = 'Success', 0, 1) 
";
            sql = string.Format(sql, sqlWhr);

            DataTable dtTotalListData = daoSelect.SqlSelect(sql);
            ViewState["S124P_TotalListData"] = dtTotalListData;
            #endregion

            #region 取得 SumData
            sql = @"
WITH BasicData AS
(
 SELECT * FROM 
 (
  SELECT 事業部, 狀態, SUBSTRING(報價日期, 1, 7) AS YM, COUNT(*) AS QTY 
  FROM RPS_S124P (NOLOCK)
  WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' AND 計入件數 = 'Y' --AND 短品名 = 'TMMP08X8_WSRF_MLO'
  {0}
  --AND 報價日期 LIKE '2023%'
  GROUP BY 事業部, 狀態, SUBSTRING(報價日期, 1, 7)
 ) t
),
BasicSchema AS 
(
 SELECT * 
 FROM 
 (
  SELECT '全部' AS 事業部, 0 AS SeqDept
  UNION ALL
  SELECT DISTINCT 事業部, 1 AS SeqDept
  FROM BasicData
 ) t,
 (  
  SELECT '報價件數' AS 狀態, 1 AS SeqStatus
  UNION 
  SELECT 'Success' AS 狀態, 2 AS SeqStatus
  UNION 
  SELECT 'Waiting' AS 狀態, 3 AS SeqStatus
  UNION 
  SELECT 'Failed' AS 狀態, 4 AS SeqStatus
  UNION 
  SELECT '報價成功率(%)' AS 狀態, 5 AS SeqStatus
 ) t1
 --ORDER BY SeqDept, 事業部, SeqStatus
 ),
 S124P AS
(
 /*(全部) 報價件數*/
 SELECT '全部' AS 事業部, '報價件數' AS 狀態, YM, SUM(QTY) AS QTY
 FROM BasicData 
 GROUP BY YM

 /*(全部) 依 狀態 的統計件數*/
 UNION ALL
 SELECT '全部' AS 事業部, 狀態, YM,  SUM(QTY) AS QTY
 FROM BasicData 
 GROUP BY 狀態, YM

 /*(全部) 報價成功率*/
 UNION ALL
 SELECT  '全部' AS 事業部, '報價成功率(%)' AS 狀態, YM,  
 ROUND(CAST(SUM(QTY) AS FLOAT) / (SELECT CAST(SUM(QTY) AS FLOAT) FROM BasicData B0 WHERE B0.YM = BasicData.YM), 2) * 100 AS Rate
 FROM BasicData
 WHERE 狀態 = 'Success'
 GROUP BY YM

 /*依 事業部 的報價件數*/
 UNION ALL
 SELECT 事業部, '報價件數' AS 狀態, YM, SUM(QTY) AS QTY
 FROM BasicData 
 GROUP BY 事業部, YM

 /*依 事業部, 狀態 的統計件數*/
 UNION ALL
 SELECT 事業部, 狀態, YM, SUM(QTY) AS QTY
 FROM BasicData 
 GROUP BY 事業部, 狀態, YM

 /*依 事業部 的報價成功率*/
 UNION ALL
 SELECT 事業部, '報價成功率(%)' AS 狀態, YM,  
 ROUND(CAST(SUM(QTY) AS FLOAT) / (SELECT CAST(SUM(QTY) AS FLOAT) FROM BasicData B0 WHERE B0.事業部 = BasicData.事業部 AND B0.YM = BasicData.YM), 2) * 100 AS Rate
 FROM BasicData
 WHERE 狀態 = 'Success'
 GROUP BY 事業部, YM
)

SELECT R.事業部, R.狀態, 
ISNULL(R1.QTY, 0) AS '1月',  
ISNULL(R2.QTY, 0) AS '2月', 
ISNULL(R3.QTY, 0) AS '3月',  
ISNULL(R4.QTY, 0) AS '4月',  
ISNULL(R5.QTY, 0) AS '5月',  
ISNULL(R6.QTY, 0) AS '6月', 
ISNULL(R7.QTY, 0) AS '7月',  
ISNULL(R8.QTY, 0) AS '8月', 
ISNULL(R9.QTY, 0) AS '9月',  
ISNULL(R10.QTY, 0) AS '10月',  
ISNULL(R11.QTY, 0) AS '11月',  
ISNULL(R12.QTY, 0) AS '12月'  
FROM BasicSchema R
LEFT JOIN S124P R1 ON R.事業部 = R1.事業部 AND R.狀態 = R1.狀態 AND R1.YM = @1M
LEFT JOIN S124P R2 ON R.事業部 = R2.事業部 AND R.狀態 = R2.狀態 AND R2.YM = @2M
LEFT JOIN S124P R3 ON R.事業部 = R3.事業部 AND R.狀態 = R3.狀態 AND R3.YM = @3M
LEFT JOIN S124P R4 ON R.事業部 = R4.事業部 AND R.狀態 = R4.狀態 AND R4.YM = @4M
LEFT JOIN S124P R5 ON R.事業部 = R5.事業部 AND R.狀態 = R5.狀態 AND R5.YM = @5M
LEFT JOIN S124P R6 ON R.事業部 = R6.事業部 AND R.狀態 = R6.狀態 AND R6.YM = @6M
LEFT JOIN S124P R7 ON R.事業部 = R7.事業部 AND R.狀態 = R7.狀態 AND R7.YM = @7M
LEFT JOIN S124P R8 ON R.事業部 = R8.事業部 AND R.狀態 = R8.狀態 AND R8.YM = @8M
LEFT JOIN S124P R9 ON R.事業部 = R9.事業部 AND R.狀態 = R9.狀態 AND R9.YM = @9M
LEFT JOIN S124P R10 ON R.事業部 = R10.事業部 AND R.狀態 = R10.狀態 AND R10.YM = @10M
LEFT JOIN S124P R11 ON R.事業部 = R11.事業部 AND R.狀態 = R11.狀態 AND R11.YM = @11M
LEFT JOIN S124P R12 ON R.事業部 = R12.事業部 AND R.狀態 = R12.狀態 AND R12.YM = @12M
ORDER BY R.SeqDept, R.事業部, R.SeqStatus
";
            sql = string.Format(sql, sqlWhr);

            DataTable dtSumData = daoSelect.SqlSelect(sql);
            ViewState["S124P_SumData"] = dtSumData;
            gvSumData.DataSource = dtSumData;
            gvSumData.DataBind();
            #endregion

            #region 取得 FailedData
            sql = @"
WITH BasicData AS
(
 SELECT * FROM 
 (
  SELECT 事業部, ISNULL(R.報價失敗原因, '其他') AS 報價失敗原因, SUBSTRING(報價日期, 1, 7) AS YM, COUNT(*) AS QTY 
  FROM RPS_S124P R (NOLOCK)
  WHERE 狀態 = 'Failed' AND 客戶代號 != 'CHPT-008' AND 計入件數 = 'Y' --AND 短品名 = 'TMMP08X8_WSRF_MLO'
  {0}
  --AND 報價日期 LIKE '2023%'
  GROUP BY 事業部, 報價失敗原因, SUBSTRING(報價日期, 1, 7)
 ) t
),
BasicSchema AS 
(
 SELECT * 
 FROM 
 (
  SELECT '全部' AS 事業部, 0 AS SeqDept
  UNION ALL
  SELECT DISTINCT 事業部, 1 AS SeqDept
  FROM BasicData
 ) t,
 (  
  SELECT ParamValue AS 報價失敗原因, ParamNo AS  SeqStatus
  FROM SA0BP1A
 ) t1
 --ORDER BY SeqDept, 事業部, SeqStatus
 ),
 S124P AS
(
 /*(全部) 報價失敗原因*/
 SELECT '全部' AS 事業部, 報價失敗原因, YM, SUM(QTY) AS QTY
 FROM BasicData 
 GROUP BY 報價失敗原因, YM

 /*依 事業部 統計 報價失敗原因*/
 UNION ALL
 SELECT 事業部, 報價失敗原因, YM, SUM(QTY) AS QTY
 FROM BasicData 
 GROUP BY 事業部, 報價失敗原因, YM
)

SELECT R.事業部, R.報價失敗原因, 
ISNULL(R1.QTY, 0) AS '1月',  
ISNULL(R2.QTY, 0) AS '2月', 
ISNULL(R3.QTY, 0) AS '3月',  
ISNULL(R4.QTY, 0) AS '4月',  
ISNULL(R5.QTY, 0) AS '5月',  
ISNULL(R6.QTY, 0) AS '6月', 
ISNULL(R7.QTY, 0) AS '7月',  
ISNULL(R8.QTY, 0) AS '8月', 
ISNULL(R9.QTY, 0) AS '9月',  
ISNULL(R10.QTY, 0) AS '10月',  
ISNULL(R11.QTY, 0) AS '11月',  
ISNULL(R12.QTY, 0) AS '12月'  
FROM BasicSchema R
LEFT JOIN S124P R1 ON R.事業部 = R1.事業部 AND R.報價失敗原因 = R1.報價失敗原因 AND R1.YM = @1M
LEFT JOIN S124P R2 ON R.事業部 = R2.事業部 AND R.報價失敗原因 = R2.報價失敗原因 AND R2.YM = @2M
LEFT JOIN S124P R3 ON R.事業部 = R3.事業部 AND R.報價失敗原因 = R3.報價失敗原因 AND R3.YM = @3M
LEFT JOIN S124P R4 ON R.事業部 = R4.事業部 AND R.報價失敗原因 = R4.報價失敗原因 AND R4.YM = @4M
LEFT JOIN S124P R5 ON R.事業部 = R5.事業部 AND R.報價失敗原因 = R5.報價失敗原因 AND R5.YM = @5M
LEFT JOIN S124P R6 ON R.事業部 = R6.事業部 AND R.報價失敗原因 = R6.報價失敗原因 AND R6.YM = @6M
LEFT JOIN S124P R7 ON R.事業部 = R7.事業部 AND R.報價失敗原因 = R7.報價失敗原因 AND R7.YM = @7M
LEFT JOIN S124P R8 ON R.事業部 = R8.事業部 AND R.報價失敗原因 = R8.報價失敗原因 AND R8.YM = @8M
LEFT JOIN S124P R9 ON R.事業部 = R9.事業部 AND R.報價失敗原因 = R9.報價失敗原因 AND R9.YM = @9M
LEFT JOIN S124P R10 ON R.事業部 = R10.事業部 AND R.報價失敗原因 = R10.報價失敗原因 AND R10.YM = @10M
LEFT JOIN S124P R11 ON R.事業部 = R11.事業部 AND R.報價失敗原因 = R11.報價失敗原因 AND R11.YM = @11M
LEFT JOIN S124P R12 ON R.事業部 = R12.事業部 AND R.報價失敗原因 = R12.報價失敗原因 AND R12.YM = @12M
WHERE 
ISNULL(R1.QTY, 0) != 0
OR ISNULL(R2.QTY, 0) != 0
OR ISNULL(R3.QTY, 0) != 0
OR ISNULL(R4.QTY, 0) != 0
OR ISNULL(R5.QTY, 0) != 0
OR ISNULL(R6.QTY, 0) != 0
OR ISNULL(R7.QTY, 0) != 0
OR ISNULL(R8.QTY, 0) != 0
OR ISNULL(R9.QTY, 0) != 0
OR ISNULL(R10.QTY, 0) != 0
OR ISNULL(R11.QTY, 0) != 0
OR ISNULL(R12.QTY, 0) != 0
ORDER BY R.SeqDept, R.事業部, R.SeqStatus
";
            sql = string.Format(sql, sqlWhr);

            DataTable dtFailedData = daoSelect.SqlSelect(sql);
            ViewState["S124P_FailedData"] = dtFailedData;
            gvFailedData.DataSource = dtFailedData;
            gvFailedData.DataBind();
            #endregion

            #region 取得 DetailData
            sql = @"
WITH BasicData AS
(
 SELECT * FROM 
 (
  SELECT 事業部, R.客戶簡稱 , 狀態, SUBSTRING(報價日期, 1, 7) AS YM, COUNT(*) AS QTY 
  FROM RPS_S124P R (NOLOCK)
  WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' AND 計入件數 = 'Y' --AND 短品名 = 'TMMP08X8_WSRF_MLO'
  {0}
  --AND 報價日期 LIKE '2023%'
  GROUP BY 事業部, 客戶簡稱, 狀態, SUBSTRING(報價日期, 1, 7)
 ) t
),
BasicSchema AS 
(
  SELECT DISTINCT 事業部, 客戶簡稱
  FROM BasicData 
 ),
 S124P AS
(
 /*依 事業部 統計 報價失敗原因*/
 SELECT 事業部, 客戶簡稱, YM, 
 SUM(QTY) AS TotalQTY,
 (SELECT SUM(B0.QTY) FROM BasicData B0 WHERE B0.狀態 = 'Success' AND B0.事業部 = BasicData.事業部 AND B0.客戶簡稱 = BasicData.客戶簡稱 AND B0.YM = BasicData.YM )  AS SuccessQTY
 FROM BasicData 
 GROUP BY 事業部, 客戶簡稱, YM
)

SELECT R.事業部, R.客戶簡稱, 
CAST(ISNULL(R1.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R1.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R1.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '1月',
CAST(ISNULL(R2.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R2.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R2.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '2月',
CAST(ISNULL(R3.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R3.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R3.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '3月',
CAST(ISNULL(R4.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R4.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R4.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '4月',
CAST(ISNULL(R5.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R5.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R5.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '5月',
CAST(ISNULL(R6.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R6.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R6.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '6月',
CAST(ISNULL(R7.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R7.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R7.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '7月',
CAST(ISNULL(R8.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R8.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R8.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '8月',
CAST(ISNULL(R9.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R9.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R9.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '9月',
CAST(ISNULL(R10.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R10.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R10.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '10月',
CAST(ISNULL(R11.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R11.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R11.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '11月',
CAST(ISNULL(R12.TotalQTY, 0) AS NVARCHAR) + IIF(ISNULL(R12.TotalQTY, 0) = 0, '',  '【' + CAST(ISNULL(R12.SuccessQTY, 0) AS NVARCHAR) + '】' ) AS '12月'
FROM BasicSchema R
LEFT JOIN S124P R1 ON R.事業部 = R1.事業部 AND R.客戶簡稱 = R1.客戶簡稱 AND R1.YM = @1M
LEFT JOIN S124P R2 ON R.事業部 = R2.事業部 AND R.客戶簡稱 = R2.客戶簡稱 AND R2.YM = @2M
LEFT JOIN S124P R3 ON R.事業部 = R3.事業部 AND R.客戶簡稱 = R3.客戶簡稱 AND R3.YM = @3M
LEFT JOIN S124P R4 ON R.事業部 = R4.事業部 AND R.客戶簡稱 = R4.客戶簡稱 AND R4.YM = @4M
LEFT JOIN S124P R5 ON R.事業部 = R5.事業部 AND R.客戶簡稱 = R5.客戶簡稱 AND R5.YM = @5M
LEFT JOIN S124P R6 ON R.事業部 = R6.事業部 AND R.客戶簡稱 = R6.客戶簡稱 AND R6.YM = @6M
LEFT JOIN S124P R7 ON R.事業部 = R7.事業部 AND R.客戶簡稱 = R7.客戶簡稱 AND R7.YM = @7M
LEFT JOIN S124P R8 ON R.事業部 = R8.事業部 AND R.客戶簡稱 = R8.客戶簡稱 AND R8.YM = @8M
LEFT JOIN S124P R9 ON R.事業部 = R9.事業部 AND R.客戶簡稱 = R9.客戶簡稱 AND R9.YM = @9M
LEFT JOIN S124P R10 ON R.事業部 = R10.事業部 AND R.客戶簡稱 = R10.客戶簡稱 AND R10.YM = @10M
LEFT JOIN S124P R11 ON R.事業部 = R11.事業部 AND R.客戶簡稱 = R11.客戶簡稱 AND R11.YM = @11M
LEFT JOIN S124P R12 ON R.事業部 = R12.事業部 AND R.客戶簡稱 = R12.客戶簡稱 AND R12.YM = @12M
ORDER BY R.事業部, R.客戶簡稱
";
            sql = string.Format(sql, sqlWhr);

            DataTable dtDetailData = daoSelect.SqlSelect(sql);
            ViewState["S124P_DetailData"] = dtDetailData;
            gvDetailData.DataSource = dtDetailData;
            gvDetailData.DataBind();
            #endregion


            RowSpan(gvSumData);
            RowSpan(gvFailedData);
            RowSpan(gvDetailData);
        }
    }

    protected void RowSpan(GridView gv)
    {
        #region 彙總表合併儲存格
        int SpanIndx = -99;
        foreach (GridViewRow gvr in gv.Rows)
        {
            if (gvr.RowType == DataControlRowType.DataRow)
            {
                if (gvr.RowIndex == 0)
                    gvr.Cells[0].RowSpan = 1;
                else
                {            
                    string strBelongSectionName = ((Label)gvr.FindControl("labBelongSectionName")).Text;
                    string strBelongSectionName_pre = ((Label)gv.Rows[(gvr.RowIndex - 1)].FindControl("labBelongSectionName")).Text;

                    //比對如果名稱如果相同就合併(RowSpan+1)
                    if (strBelongSectionName == strBelongSectionName_pre)
                    {
                        SpanIndx = (SpanIndx == -99) ? gvr.RowIndex - 1 : SpanIndx;
                        gv.Rows[SpanIndx].Cells[0].RowSpan += 1;
                        gvr.Cells[0].Visible = false;
                    }
                    else
                    {
                        SpanIndx = -99;
                        gvr.Cells[0].RowSpan = 1;
                    }
                }
            }
        }
        #endregion
    }

    protected void gvData_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            GridView gv = (GridView)sender;
            string strBelongSectionName = ((Label)e.Row.FindControl("labBelongSectionName")).Text;
            Label labType = ((Label)e.Row.FindControl("labType"));

            if (strBelongSectionName == "全部")
            {
                e.Row.Enabled = false;
                e.Row.BackColor = System.Drawing.Color.LightYellow;
            }

            if (gv.ID == "gvSumData")
            {
                for (int i = 0; i < e.Row.Cells.Count; i++)
                {
                    string strColumnName = gv.HeaderRow.Cells[i].Text;

                    if (labType.Text == "報價成功率(%)")
                    {
                        labType.ForeColor = System.Drawing.Color.Blue;
                        labType.Font.Size = 12;
                        if (strColumnName.Contains("月"))
                        {
                            string strMM = strColumnName.Replace("月", "");
                            Label labMM = ((Label)e.Row.FindControl(string.Format("lab{0}M", strMM)));
                            double Value = 0;

                            double.TryParse(labMM.Text, out Value);

                            labMM.Font.Size = 12;
                            labMM.Font.Bold = true;
                            labMM.ForeColor = (Value >= 80) ? System.Drawing.Color.Green : System.Drawing.Color.Red;
                        }
                    }

                    #region 未來的月份 一律顯示空白
                    if (strColumnName.Contains("月"))
                    {
                        string strMM = strColumnName.Replace("月", "");
                        string QueryDate = ddlYear.SelectedValue + "/" + strMM.PadLeft(2, '0') + "/01";
                        string NowDate = DateTime.Now.ToString("yyyy/MM") + "/01";
                        if (Convert.ToDateTime(NowDate) < Convert.ToDateTime(QueryDate))
                        {
                            e.Row.Cells[i].Text = "";
                        }
                    }
                    #endregion
                }
            }

            if (gv.ID == "gvFailedData")
            {
                for (int i = 0; i < e.Row.Cells.Count; i++)
                {
                    string strColumnName = gv.HeaderRow.Cells[i].Text;

                    #region 未來的月份 一律顯示空白
                    if (strColumnName.Contains("月"))
                    {
                        string strMM = strColumnName.Replace("月", "");
                        Label labMM = ((Label)e.Row.FindControl(string.Format("lab{0}M", strMM)));
                        double Value = 0;
                        double.TryParse(labMM.Text, out Value);

                        if (Value >= 5)
                        {
                            labMM.Font.Bold = true;
                            labMM.Font.Size = 12;
                            labMM.ForeColor = System.Drawing.Color.Red;
                        }
                        else if (Value >= 3)
                            labMM.ForeColor = System.Drawing.Color.Red;
                        else if (Value == 0)
                            labMM.Text = "";
                    }
                    #endregion
                }
            }

            if (gv.ID == "gvDetailData")
            {
                DataTable dtTotalListData = new DataTable();
                if (ViewState["S124P_TotalListData"] != null)
                     dtTotalListData = (DataTable)ViewState["S124P_TotalListData"];

                for (int i = 0; i < e.Row.Cells.Count; i++)
                {
                    string strColumnName = gv.HeaderRow.Cells[i].Text;

                    #region 未來的月份 一律顯示空白
                    if (strColumnName.Contains("月"))
                    {
                        string strMM = strColumnName.Replace("月", "");
                        Label labMM = ((Label)e.Row.FindControl(string.Format("lab{0}M", strMM)));
                        string strValue = labMM.Text;

                        if (strValue == "0")
                            labMM.Text = "";
                        else
                        {
                            var list = dtTotalListData.AsEnumerable().Where(x => x["事業部"].ToString() == strBelongSectionName 
                                && x["YM"].ToString() == (ddlYear.SelectedValue + "/" + strMM.PadLeft(2, '0')) 
                                && x["客戶簡稱"].ToString() == labType.Text);

                            string strTooltip = "";
                            foreach (DataRow dr in list)
                                strTooltip += dr["短品名"].ToString() + "\r\n";

                            labMM.ToolTip += strTooltip;
                            //e.Row.Cells[i].ToolTip = "1213";
                        }
                    }
                    #endregion
                }
            }

            //double TaxSumAmount = 0;
            //double.TryParse(e.Row.Cells[6].Text, out TaxSumAmount);
            //e.Row.Cells[6].Text = String.Format("{0:###,###.##}", TaxSumAmount);

            //e.Row.Cells[6].Text = string.IsNullOrWhiteSpace(e.Row.Cells[6].Text) ? "0" : e.Row.Cells[6].Text;

            //e.Row.Cells[0].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[1].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[3].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[4].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[5].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[6].HorizontalAlign = HorizontalAlign.Right;
        }
    }

    protected void btnExcel_Click(object sender, EventArgs e)
    {
        ERP.EPPlusExcel excel = new ERP.EPPlusExcel();
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder();

        if (gvSumData.Rows.Count == 0) 
            JavaScriptFunction.ShowJsMSG(this, "請先執行查詢");
        else
        {
            ERP.DBDao daoSelect = new ERP.DBDao();
            string sql = "", sqlWhr = "";

            #region RFQ轉內訂總表
            DataTable dtSumData = (DataTable)ViewState["S124P_SumData"];
            excel.SetExcelWorkSheet("RFQ轉內訂總表");
            excel.AddContentData(0, 1, dtSumData, true, true, true);
            excel.SetAutoFit(0, 1, dtSumData.Columns.Count);
            #endregion

            #region 報價失敗分析
            DataTable dtFailedData = (DataTable)ViewState["S124P_FailedData"];
            excel.SetExcelWorkSheet("報價失敗分析");
            excel.AddContentData(0, 1, dtFailedData, true, true, true);
            excel.SetAutoFit(0, 1, dtFailedData.Columns.Count);
            #endregion

            #region RFQ明細By客戶
            DataTable dtDetailData = (DataTable)ViewState["S124P_DetailData"];
            excel.SetExcelWorkSheet("RFQ明細By客戶");
            excel.AddContentData(0, 1, dtDetailData, true, true, true);
            excel.SetAutoFit(0, 1, dtDetailData.Columns.Count);
            #endregion

            #region RFQ明細總表
            DataTable dtDetailListData = (DataTable)ViewState["S124P_DetailListData"];
            excel.SetExcelWorkSheet("RFQ明細總表");
            excel.AddContentData(0, 1, dtDetailListData, true, true, true);
            excel.SetAutoFit(0, 1, dtDetailListData.Columns.Count);
            #endregion

            bool isSuccess = false;
            isSuccess = excel.ExportToExcel(this, string.Format("報價成功率報表({0})", ddlYear.SelectedValue), out sbErrMsg);
            if (!isSuccess)
                JavaScriptFunction.ShowJsMSG(this, sbErrMsg.ToString());
        }
    }

    protected void btnResetData_Click(object sender, EventArgs e)
    {
        ERP.DBDao daoSelect = new ERP.DBDao();
        ERP.DBDao daoTrans = new ERP.DBDao();
        daoTrans.SqlCommand.CommandTimeout = 300;

        System.Text.StringBuilder sbMsg = new System.Text.StringBuilder();
        bool isSuccess = true;
        try
        {
            daoTrans.SqlConn.Open();
            daoTrans.BeginTransaction();

            string sql = daoSelect.SqlSelectToString("SELECT SQL FROM CHPT.dbo.SQLList WITH(NOLOCK) WHERE ID = '723'", "SQL");
            isSuccess = daoTrans.SqlUpdate(sql, ref sbMsg);
        }
        catch (Exception ex)
        {
            isSuccess = false;
            sbMsg.Append(ex);
        }
        finally
        {
            if (isSuccess)
            {
                daoTrans.Commit();
                JavaScriptFunction.ShowJsMSG(this, "[更新成功]");
                btnQuery_Click(null, null);
            }
            else
            {
                daoTrans.Rollback();
                JavaScriptFunction.ShowJsMSG(this, string.Format("[更新失敗]\\n{0}", sbMsg.ToString()));
            }
        }
    }

    private void CloseDialog(string dialogId)
    {
        string script = string.Format(@"closeDialog('{0}')", dialogId);
        ScriptManager.RegisterClientScriptBlock(this, typeof(Page), UniqueID, script, true);
    }
    private void ShowDialog(string dialogId)
    {
        string script = string.Format(@"showDialog('{0}')", dialogId);
        ScriptManager.RegisterClientScriptBlock(this, typeof(Page), UniqueID, script, true);
    }

    #region old Code 不使用 暫先不移除
    protected void btnQuery_Click_Old(object sender, EventArgs e)
    {
        string strAlertMsg = "", sql = "", sqlWhr = "";

        if (string.IsNullOrWhiteSpace(txtStartDate.Text) || string.IsNullOrWhiteSpace(txtEndDate.Text))
            strAlertMsg += "起訖區間不可空白";

        if (string.IsNullOrWhiteSpace(rblBelongSection.SelectedValue))
            strAlertMsg += "事業部 不可空白";

        if (!string.IsNullOrWhiteSpace(strAlertMsg))
            JavaScriptFunction.ShowJsMSG(this, strAlertMsg);
        else
        {
            ERP.DBDao daoSelect = new ERP.DBDao();

            #region Parameters設置
            sqlWhr += (string.IsNullOrWhiteSpace(txtStartDate.Text)) ? "" : " AND 報價日期 >= @StartDate AND 報價日期 <= @EndDate ";
            sqlWhr += (string.IsNullOrWhiteSpace(lsbObject.SelectedValue)) ? "" : " AND 客戶代號 = @ObjectNo ";
            sqlWhr += (string.IsNullOrWhiteSpace(lsbSalesID.SelectedValue)) ? "" : " AND 業務工號 = @SalesID ";

            if (rblCostSiteName.SelectedValue != "全部")
                sqlWhr += (string.IsNullOrWhiteSpace(rblCostSiteName.SelectedValue)) ? "" : " AND 成本據點 = @CostSiteName ";

            if (rblBelongSection.SelectedValue != "全部")
                sqlWhr += (string.IsNullOrWhiteSpace(rblBelongSection.SelectedValue)) ? "" : " AND 事業部代號 = @BelongSection ";

            string strSortTypeA = "";
            foreach (ListItem list in cblSortTypeA.Items)
            {
                if (string.IsNullOrWhiteSpace(strSortTypeA))
                    strSortTypeA += (list.Selected) ? string.Format("'{0}'", list.Text) : "";
                else
                    strSortTypeA += (list.Selected) ? string.Format(",'{0}'", list.Text) : "";
            }
            sqlWhr += (string.IsNullOrWhiteSpace(strSortTypeA)) ? "" : string.Format(" AND 產品屬性 IN ({0}) ", strSortTypeA);

            daoSelect.SqlCommand.Parameters.Clear();
            daoSelect.SqlCommand.Parameters.AddWithValue("@StartDate", txtStartDate.Text.Replace("-", "/"));
            daoSelect.SqlCommand.Parameters.AddWithValue("@EndDate", txtEndDate.Text.Replace("-", "/"));
            daoSelect.SqlCommand.Parameters.AddWithValue("@ObjectNo", lsbObject.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@SalesID", lsbSalesID.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@BelongSection", rblBelongSection.SelectedValue);
            daoSelect.SqlCommand.Parameters.AddWithValue("@CostSiteName", rblCostSiteName.SelectedValue);
            #endregion

            #region 取得明細表Data
            string sqlDetail = @"
WITH S124P AS
(
	SELECT * FROM RPS_S124P (NOLOCK)
	WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' --AND 短品名 = 'TMRF62X10_TFM'
	{0}
)

SELECT  計入件數, 計入金額, 狀態, 業務, 客戶簡稱, 產品屬性, 短品名, SUM(數量) AS 數量, FORMAT(SUM(總價), 'N0') AS 總價, 報價失敗原因, 狀態備註紀錄, 報價單號, 報價月份, 內部訂單, 訂單月份, 事業部
FROM 
(
		SELECT  計入件數, 計入金額, 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 數量, 總價, 報價失敗原因, 狀態備註紀錄, 
				(
					SELECT (
						SELECT  CAST(報價單號 AS NVARCHAR) + ';'
						FROM S124P R0 (NOLOCK)
						WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
						FOR XML PATH('')
					)
				) AS 報價單號, 
				(
					SELECT TOP(1) SUBSTRING(報價日期, 1, 7)
					FROM S124P R0 (NOLOCK)
					WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
					ORDER BY 報價單號
				) AS 報價月份,
				(
					SELECT (
						SELECT  CAST(內部訂單 AS NVARCHAR) + ';'
						FROM S124P R0 (NOLOCK)
						WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
						FOR XML PATH('')
					)
				) AS 內部訂單, 訂單月份, 
				事業部代號, 事業部, 成本據點  
		FROM S124P R (NOLOCK) 
		WHERE R.產品屬性 = '新設計案'
		UNION ALL
		SELECT  計入件數, 計入金額, 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 數量, 總價, 報價失敗原因, 狀態備註紀錄, 
		       報價單號, 
			   SUBSTRING(報價日期, 1, 7) AS 報價月份, 
		       內部訂單,訂單月份,
			   事業部代號, 事業部, 成本據點 
		FROM S124P R (NOLOCK) 
		WHERE R.產品屬性 != '新設計案' 
) t
GROUP BY  計入件數, 計入金額, 狀態, 業務, 客戶簡稱, 產品屬性, 短品名, 報價失敗原因, 狀態備註紀錄, 報價單號, 報價月份, 內部訂單, 訂單月份, 事業部 
ORDER BY 業務, 短品名, 計入件數, 狀態, 數量, 總價
";
            sqlDetail = string.Format(sqlDetail, sqlWhr);

            DataTable dtDetail = daoSelect.SqlSelect(sqlDetail);
            ViewState["S124P_DetailData"] = dtDetail;
            #endregion

            #region 取得彙總表Data
            sql = @"
WITH S124P AS
(
	SELECT * 
	FROM RPS_S124P (NOLOCK)
	WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' AND 計入金額 = 'Y' --AND 短品名 = 'TMMP08X8_WSRF_MLO'
	{0}
	--AND 報價日期 >= '2022/01/01' AND 報價日期 <= '2022/12/31'
),
S124P2 AS
(
	SELECT 事業部, 狀態, COUNT(*) AS TotalQTY 
	FROM RPS_S124P (NOLOCK)
	WHERE 狀態 != 'Other' AND 客戶代號 != 'CHPT-008' AND 計入件數 = 'Y' --AND 短品名 = 'TMMP08X8_WSRF_MLO'
	{0}
	--AND 報價日期 >= '2022/01/01' AND 報價日期 <= '2022/12/31'
	GROUP BY 事業部, 狀態
),
Bdata AS
(
	SELECT 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, SUM(數量) AS 數量, SUM(總價) AS 總價, 報價失敗原因, 狀態備註紀錄, 報價單號, 內部訂單, 事業部代號, 事業部, 成本據點  
	FROM 
	(
		SELECT 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 數量, 總價, 報價失敗原因, 狀態備註紀錄, 
						(
							SELECT (
								SELECT  CAST(報價單號 AS NVARCHAR) + ';'
								FROM S124P R0 (NOLOCK)
								WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
								FOR XML PATH('')
							)
						) AS 報價單號, 
						(
							SELECT (
								SELECT  CAST(內部訂單 AS NVARCHAR) + ';'
								FROM S124P R0 (NOLOCK)
								WHERE R0.短品名 = R.短品名 AND R0.狀態 = R.狀態 AND R0.產品屬性 = R.產品屬性
								FOR XML PATH('')
							)
						) AS 內部訂單, 
						 事業部代號, 事業部, 成本據點  
		FROM S124P R (NOLOCK) 
		WHERE R.產品屬性 = '新設計案'
		UNION ALL
		SELECT 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 數量, 總價, 報價失敗原因, 狀態備註紀錄, 報價單號, 內部訂單, 事業部代號, 事業部, 成本據點 
		FROM S124P R (NOLOCK) 
		WHERE R.產品屬性 != '新設計案'
	) t
	GROUP BY 狀態, 業務工號, 業務, 客戶代號, 客戶簡稱, 產品屬性, 短品名, 報價失敗原因, 狀態備註紀錄, 報價單號, 內部訂單, 事業部, 事業部代號, 事業部, 成本據點   
),
TotalByDept AS
(
	SELECT 事業部, ISNULL((SELECT SUM(TotalQTY) FROM S124P2 WHERE S124P2.事業部 = Bdata.事業部), 0) AS TotalQTY, SUM(總價) AS TotalAmount
	FROM Bdata 
	GROUP BY 事業部
),
StatusQTYByDept AS
(
	SELECT 事業部, 狀態, ISNULL((SELECT TotalQTY FROM S124P2 WHERE S124P2.事業部 = Bdata.事業部 AND S124P2.狀態 = Bdata.狀態), 0) AS TotalQTY, SUM(總價) AS TotalAmount
	FROM Bdata 
	GROUP BY 事業部, 狀態
),
ResultQTY AS
(
	SELECT T0.事業部, 
	ISNULL(T0.TotalQTY, 0) TotalQTY, 
	ISNULL(T1.TotalQTY, 0) AS SuccessQTY, 
	ISNULL(T2.TotalQTY, 0) AS WaitingQTY, 
	ISNULL(T3.TotalQTY, 0) AS FailedQTY,
	IIF(T0.TotalQTY > 0, ROUND(CAST(ISNULL(T1.TotalQTY, 0) AS FLOAT) / CAST(ISNULL(T0.TotalQTY, 0) AS FLOAT), 2) * 100, 0) AS SuccessRate
	--ROUND(CAST(ISNULL(T1.TotalQTY, 0) AS FLOAT) / CAST(ISNULL(T0.TotalQTY, 0) AS FLOAT), 2) * 100 AS SuccessRate
	FROM TotalByDept T0
	LEFT JOIN StatusQTYByDept T1 ON T0.事業部 = T1.事業部 AND T1.狀態 = 'Success'
	LEFT JOIN StatusQTYByDept T2 ON T0.事業部 = T2.事業部 AND T2.狀態 = 'Waiting'
	LEFT JOIN StatusQTYByDept T3 ON T0.事業部 = T3.事業部 AND T3.狀態 = 'Failed'
),
ResultAmount AS
(
	SELECT T0.事業部, 
	ISNULL(T0.TotalAmount, 0) TotalAmount, 
	ISNULL(T1.TotalAmount, 0) AS SuccessAmount, 
	ISNULL(T2.TotalAmount, 0) AS WaitingAmount, 
	ISNULL(T3.TotalAmount, 0) AS FailedAmount,
	IIF(T0.TotalQTY > 0, ROUND(CAST(ISNULL(T1.TotalQTY, 0) AS FLOAT) / CAST(ISNULL(T0.TotalQTY, 0) AS FLOAT), 2) * 100, 0) AS SuccessRate
	--ROUND(CAST(ISNULL(T1.TotalQTY, 0) AS FLOAT) / CAST(ISNULL(T0.TotalQTY, 0) AS FLOAT), 2) * 100 AS SuccessRate
	FROM TotalByDept T0
	LEFT JOIN StatusQTYByDept T1 ON T0.事業部 = T1.事業部 AND T1.狀態 = 'Success'
	LEFT JOIN StatusQTYByDept T2 ON T0.事業部 = T2.事業部 AND T2.狀態 = 'Waiting'
	LEFT JOIN StatusQTYByDept T3 ON T0.事業部 = T3.事業部 AND T3.狀態 = 'Failed'
),
ResultData AS
(
	SELECT ResultQTY.事業部, '件數'[類別], 
	CAST(ResultQTY.TotalQTY AS NVARCHAR) 總計, 
	CAST(ResultQTY.SuccessQTY AS NVARCHAR) 報價成功, 
	CAST(ResultQTY.WaitingQTY AS NVARCHAR) 待客戶回覆, 
	CAST(ResultQTY.FailedQTY AS NVARCHAR) 報價失敗, 
	ResultQTY.SuccessRate 報價成功率
	FROM ResultQTY
	UNION ALL
	SELECT ResultAmount.事業部, '金額'[Type], 
	FORMAT(ResultAmount.TotalAmount, 'N0'), 
	FORMAT(ResultAmount.SuccessAmount, 'N0'), 
	FORMAT(ResultAmount.WaitingAmount, 'N0'), 
	FORMAT(ResultAmount.FailedAmount, 'N0'), 
	ResultAmount.SuccessRate
	FROM ResultAmount
)

--SELECT 
--ResultQTY.事業部, 
--ResultQTY.TotalQTY, ResultAmount.TotalAmount, 
--ResultQTY.SuccessQTY, ResultAmount.SuccessAmount, 
--ResultQTY.WaitingQTY, ResultAmount.WaitingAmount,
--ResultQTY.FailedQTY, ResultAmount.FailedAmount,
--ResultQTY.SuccessRate
--FROM ResultQTY
--LEFT JOIN ResultAmount ON ResultQTY.事業部 = ResultAmount.事業部
--ORDER BY ResultQTY.事業部


SELECT * FROM 
(
	SELECT 1 AS Seq, * FROM ResultData
	UNION ALL
	SELECT 0 AS Seq, '全部', '件數' AS [類別], 
	CAST(SUM(ResultQTY.TotalQTY) AS NVARCHAR) 總計, 
	CAST(SUM(ResultQTY.SuccessQTY) AS NVARCHAR) 報價成功, 
	CAST(SUM(ResultQTY.WaitingQTY) AS NVARCHAR) 待客戶回覆, 
	CAST(SUM(ResultQTY.FailedQTY) AS NVARCHAR) 報價失敗, 
	IIF(SUM(CAST(TotalQTY AS FLOAT)) > 0, ROUND(SUM(CAST(SuccessQTY AS FLOAT)) / SUM(CAST(TotalQTY AS FLOAT)), 2) * 100, 0) 報價成功率
	--ROUND(SUM(CAST(SuccessQTY AS FLOAT)) / SUM(CAST(TotalQTY AS FLOAT)), 2) * 100 報價成功率
	FROM ResultQTY
	UNION ALL
	SELECT 0 AS Seq, '全部', '金額' AS [類別], 
	FORMAT(SUM(ResultAmount.TotalAmount), 'N0'), 
	FORMAT(SUM(ResultAmount.SuccessAmount), 'N0'), 
	FORMAT(SUM(ResultAmount.WaitingAmount), 'N0'), 
	FORMAT(SUM(ResultAmount.FailedAmount), 'N0'), 
	(SELECT IIF(SUM(CAST(TotalQTY AS FLOAT)) > 0, ROUND(SUM(CAST(SuccessQTY AS FLOAT)) / SUM(CAST(TotalQTY AS FLOAT)), 2) * 100, 0) FROM ResultQTY ) 
    --(SELECT ROUND(SUM(CAST(SuccessQTY AS FLOAT)) / SUM(CAST(TotalQTY AS FLOAT)), 2) * 100 FROM ResultQTY )
	FROM ResultAmount
) t
ORDER BY Seq, 事業部, [類別]
";

            sql = string.Format(sql, sqlWhr);
            DataTable dt = daoSelect.SqlSelect(sql);
            gvContent.DataSource = dt;
            gvContent.DataBind();
            ViewState["S124P_TotalData"] = dt;
            #endregion

            #region 彙總表合併儲存格
            foreach (GridViewRow gvr in gvContent.Rows)
            {
                if (gvr.RowType == DataControlRowType.DataRow)
                {
                    if (gvr.RowIndex != 0)
                    {
                        string strBelongSectionName = ((Label)gvr.FindControl("labBelongSectionName")).Text;
                        string strBelongSectionName_pre = ((Label)gvContent.Rows[(gvr.RowIndex - 1)].FindControl("labBelongSectionName")).Text;

                        //比對如果名稱如果相同就合併(RowSpan+1)
                        if (strBelongSectionName == strBelongSectionName_pre)
                        {
                            gvContent.Rows[(gvr.RowIndex - 1)].Cells[0].RowSpan += 1;
                            gvr.Cells[0].Visible = false;

                            gvContent.Rows[(gvr.RowIndex - 1)].Cells[6].RowSpan += 1;
                            gvr.Cells[6].Visible = false;
                        }
                        else
                        {
                            gvr.Cells[0].RowSpan = 1;
                            gvr.Cells[6].RowSpan = 1;
                        }
                    }
                    else
                    {
                        gvr.Cells[0].RowSpan = 1;
                        gvr.Cells[6].RowSpan = 1;
                    }
                }
            }
            #endregion
        }
    }

    protected void gvContent_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            string strBelongSectionName = ((Label)e.Row.FindControl("labBelongSectionName")).Text;
            if (e.Row.RowIndex == 0 || e.Row.RowIndex == 1)
            {
                e.Row.Enabled = false;
                e.Row.BackColor = System.Drawing.Color.LightYellow;
            }

            //double TaxSumAmount = 0;
            //double.TryParse(e.Row.Cells[6].Text, out TaxSumAmount);
            //e.Row.Cells[6].Text = String.Format("{0:###,###.##}", TaxSumAmount);

            //e.Row.Cells[6].Text = string.IsNullOrWhiteSpace(e.Row.Cells[6].Text) ? "0" : e.Row.Cells[6].Text;

            //e.Row.Cells[0].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[1].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[2].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[3].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[4].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[5].HorizontalAlign = HorizontalAlign.Center;
            //e.Row.Cells[6].HorizontalAlign = HorizontalAlign.Right;
        }
    }

    protected void btnStatus_Click(object sender, EventArgs e)
    {
        Button btn = (Button)sender;
        string strBelongSection = ((Label)(btn.FindControl("labBelongSectionName"))).Text;
        string strOrderStatus = btn.ID.Replace("btn", "");

        labPopBelongSection.Text = strBelongSection;
        labPopOrderStatus.Text = strOrderStatus;

        DataTable dtDetail = ((DataTable)ViewState["S124P_DetailData"]);

        var Detail = dtDetail.AsEnumerable()
             .Where(d => d["狀態"].ToString() == strOrderStatus)
             .Where(d => d["事業部"].ToString() == strBelongSection)
             .OrderBy(d => d.Field<string>("業務"))
             .OrderBy(d => d.Field<string>("狀態"));

        gvDialogDetail.DataSource = Detail.CopyToDataTable();
        gvDialogDetail.DataBind();
        ShowDialog("dialogDetail");
    }

    protected void gvDialogDetail_RowDataBound(object sender, GridViewRowEventArgs e)
    {
        if (e.Row.RowType == DataControlRowType.DataRow)
        {
            e.Row.Cells[8].HorizontalAlign = HorizontalAlign.Right;
            e.Row.Cells[8].ForeColor = System.Drawing.Color.Blue;
        }
    }

    protected void btnExcel_Click_Old(object sender, EventArgs e)
    {
        ERP.EPPlusExcel excel = new ERP.EPPlusExcel();
        System.Text.StringBuilder sbErrMsg = new System.Text.StringBuilder();

        if (gvContent.Rows.Count == 0)
            JavaScriptFunction.ShowJsMSG(this, "請先執行查詢");
        else
        {
            ERP.DBDao daoSelect = new ERP.DBDao();
            string sql = "", sqlWhr = "";

            #region 彙總表
            DataTable dtTotalData = (DataTable)ViewState["S124P_TotalData"];
            dtTotalData.Columns.Remove("Seq");
            excel.SetExcelWorkSheet("彙總表");
            excel.AddContentData(0, 1, dtTotalData, true, true, true);
            excel.SetAutoFit(0, 1, dtTotalData.Columns.Count);
            #endregion

            #region 明細表
            DataTable dtDetailData = (DataTable)ViewState["S124P_DetailData"];
            excel.SetExcelWorkSheet("明細表");
            excel.AddContentData(0, 1, dtDetailData, true, true, true);
            excel.SetAutoFit(0, 1, dtDetailData.Columns.Count);
            #endregion

            bool isSuccess = false;
            isSuccess = excel.ExportToExcel(this, string.Format("報價成功率報表({0}-{1})", txtStartDate.Text.Replace("-", "/"), txtEndDate.Text.Replace("-", "/")), out sbErrMsg);
            if (!isSuccess)
                JavaScriptFunction.ShowJsMSG(this, sbErrMsg.ToString());
        }
    }
    #endregion 
}


