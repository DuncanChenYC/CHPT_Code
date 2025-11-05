using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// PopupByJson 的摘要描述
/// </summary>
public class PopupByJson
{
    public PopupByJson()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
    }

    public class JsonPopupSearchSqlData
    {
        public string strJSqlSelectFromWhere { get; set; }
        public string strJSqlGroupByHavingOrderBy { get; set; }
        public string strJSqlHorizontalAlign { get; set; }        
    }

    public enum PopupActionType
    {
        ondblclick,
        onclick
    }

    /// <summary>
    /// 呼叫JSqlPopupSearch視窗
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="targetID"></param>
    /// <param name="fromWhere"></param>
    /// <param name="orderBy"></param>
    /// <param name="attributesAction">[ondblclick][onclick]</param>
    public void SetControlJSWinOpenqlPopupSearchAction(object sender, string targetID, string fromWhere, string orderBy, PopupActionType popType, string maxRows)
    {
        this.SetControlJSWinOpenqlPopupSearchAction(sender, targetID, fromWhere, orderBy, "", popType, maxRows, "");
    }

    public void SetControlJSWinOpenqlPopupSearchAction(object sender, string targetID, string fromWhere, string orderBy, string horizontalAlign, PopupActionType popType, string maxRows)
    {
        this.SetControlJSWinOpenqlPopupSearchAction(sender, targetID, fromWhere, orderBy, horizontalAlign, popType, maxRows, "");
    }

    public void SetControlJSWinOpenqlPopupSearchAction(object sender, string targetID, string fromWhere, string orderBy, string horizontalAlign, PopupActionType popType, string maxRows, string dbName)
    {
        JsonPopupSearchSqlData jpSQL = new JsonPopupSearchSqlData
        {
            strJSqlSelectFromWhere = fromWhere,
            strJSqlGroupByHavingOrderBy = orderBy,
            strJSqlHorizontalAlign = horizontalAlign
        };

        System.Web.Script.Serialization.JavaScriptSerializer jssSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
        string strJSON = jssSerializer.Serialize(jpSQL);

        string javascriptAction = "";
        if(dbName == "")
            javascriptAction = string.Format("JSWinOpenqlPopupSearch(event, '{0}', {1}, '{2}');", targetID, strJSON, maxRows);
        else
            javascriptAction = string.Format("JSWinOpenqlPopupSearchByDBName(event, '{0}', {1}, '{2}', '{3}');", targetID, strJSON, maxRows, dbName);

        if (sender.GetType().Name == "TextBox")
            ((System.Web.UI.WebControls.TextBox)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "Button")
            ((System.Web.UI.WebControls.Button)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "LinkButton")
            ((System.Web.UI.WebControls.LinkButton)sender).Attributes[popType.ToString()] = javascriptAction;
    }

    public void SetControlJSWinOpenqlPopupSearchActionMultiSelect
        (object sender, string targetID, string fromWhere, string orderBy, string horizontalAlign, PopupActionType popType, string maxRows, string dbName, string uniqueValueCol)
    {
        JsonPopupSearchSqlData jpSQL = new JsonPopupSearchSqlData
        {
            strJSqlSelectFromWhere = fromWhere,
            strJSqlGroupByHavingOrderBy = orderBy,
            strJSqlHorizontalAlign = horizontalAlign
        };

        System.Web.Script.Serialization.JavaScriptSerializer jssSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
        string strJSON = jssSerializer.Serialize(jpSQL);

        string javascriptAction = "";
        if (dbName == "")
            javascriptAction = string.Format("JSWinOpenqlPopupSearchMultiSelect(event, '{0}', {1}, '{2}', '{3}');", targetID, strJSON, maxRows, uniqueValueCol);
        else
            javascriptAction = string.Format("JSWinOpenqlPopupSearchByDBNameMultiSelect(event, '{0}', {1}, '{2}', '{3}', '{4}');", targetID, strJSON, maxRows, uniqueValueCol, dbName);

        if (sender.GetType().Name == "TextBox")
            ((System.Web.UI.WebControls.TextBox)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "Button")
            ((System.Web.UI.WebControls.Button)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "LinkButton")
            ((System.Web.UI.WebControls.LinkButton)sender).Attributes[popType.ToString()] = javascriptAction;
    }

    /// <summary>
    /// 呼叫JSqlPopupSearch視窗
    /// </summary>
    /// <param name="sender"></param>
    /// <param name="targetID"></param>
    /// <param name="fromWhere"></param>
    /// <param name="orderBy"></param>
    /// <param name="attributesAction">[ondblclick][onclick]</param>
    public void SetControlJSqlPopupSearchAction(object sender, string targetID, string fromWhere, string orderBy, PopupActionType popType)
    {
        JsonPopupSearchSqlData jpSQL = new JsonPopupSearchSqlData
        {
            strJSqlSelectFromWhere = fromWhere,
            strJSqlGroupByHavingOrderBy = orderBy
        };

        System.Web.Script.Serialization.JavaScriptSerializer jssSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
        string strJSON = jssSerializer.Serialize(jpSQL);

        string javascriptAction = string.Format("JSqlPopupSearch(event, {0},'{1}');", strJSON, targetID);
        if (sender.GetType().Name == "TextBox")
            ((System.Web.UI.WebControls.TextBox)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "Button")
            ((System.Web.UI.WebControls.Button)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "LinkButton")
            ((System.Web.UI.WebControls.LinkButton)sender).Attributes[popType.ToString()] = javascriptAction;
    }

    public void SetControlJSqlPopupSearchWithPostbackAction(object sender, string targetID, string fromWhere, string orderBy, PopupActionType popType)
    {
        JsonPopupSearchSqlData jpSQL = new JsonPopupSearchSqlData
        {
            strJSqlSelectFromWhere = fromWhere,
            strJSqlGroupByHavingOrderBy = orderBy
        };

        System.Web.Script.Serialization.JavaScriptSerializer jssSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
        string strJSON = jssSerializer.Serialize(jpSQL);

        string javascriptAction = string.Format("JSqlPopupSearchWithPostBack(event, {0}, '{1}', '{2}');", strJSON, targetID, true);
        if (sender.GetType().Name == "TextBox")
            ((System.Web.UI.WebControls.TextBox)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "Button")
            ((System.Web.UI.WebControls.Button)sender).Attributes[popType.ToString()] = javascriptAction;
        else if (sender.GetType().Name == "LinkButton")
            ((System.Web.UI.WebControls.LinkButton)sender).Attributes[popType.ToString()] = javascriptAction;
    }

    public void SetButtonDisplayNone(System.Web.UI.WebControls.Button btn)
    {
        btn.Attributes.Add("style", "display:none");
    }
}