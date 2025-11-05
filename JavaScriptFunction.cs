/// <summary>
/// JavaScriptFunction 的摘要描述
/// </summary>
public class JavaScriptFunction
{
    public JavaScriptFunction()
    {
        //
        // TODO: 在這裡新增建構函式邏輯
        //
    }

    /// <summary>
    /// 字串要被宣告成JavaScript字串時的特殊處理。
    /// 例如: Response.Write("&lt;script&gt;alert('"+msg+"');&lt;/script&gt;");中的msg有特殊字元如換行時，會造成JavaScript Error。
    /// 此函數可以進行特殊字元的置換處理。
    /// </summary>
    /// <param name="raw">原始字串</param>
    /// <param name="htmlAttribute">是否要宣告在HTML Attribute中，對單雙引號加入特別處理。
    /// 例如: "&lgt;a onclick="alert('"+msg+"');"&gt;時，注意Attribute請使用雙引號包含字串</param>
    /// <returns>置換特殊字元之後的字串</returns>
    public static string JSStringEscape(string raw, bool inHtmlAttribute)
    {
        //處理跳脫字元『/』
        raw = raw.Replace("\\n", "\n").Replace("\\", "\\\\");
        //處理換行
        raw = raw.Replace("\r\n", "\\n").Replace("\r", "").Replace("\n", "\\n");        
        if (inHtmlAttribute)
            raw = raw.Replace("\"", "&quot;").Replace("'", "\\'");
        else
            raw = raw.Replace("'", "\\'").Replace("\"", "\\\"");
        return raw;
    }

    public static void ShowJsMSG(System.Web.UI.Page page, string alertMsg)
    {
        System.Web.UI.ScriptManager.RegisterStartupScript(page, page.GetType(), "Alert", string.Format("alert('{0}');", JSStringEscape(alertMsg, true)), true);
    }
    public static void ShowJsMSG(System.Web.UI.Page page, string alertMsg, bool inHtmlAttribute)
    {
        System.Web.UI.ScriptManager.RegisterStartupScript(page, page.GetType(), "Alert", string.Format("alert('{0}');", JSStringEscape(alertMsg, inHtmlAttribute)), true);
    }

    public static void ShowJsMSG(System.Web.UI.Page page, string pnlID, string alertMsg)
    {
        ShowJsHideModalPopup(page, pnlID);
        ShowJsMSG(page, alertMsg);
    }

    public static void ShowJsMSGWithModalPopup(System.Web.UI.Page page, string pnlID, string alertMsg)
    {
        ShowJsHidePanel(page, pnlID);
        ShowJsMSG(page, alertMsg);
        ShowJsShowPanel(page, pnlID);
    }

    public static void ShowJsMSG(System.Web.UI.Page page, string pnlModalPopupID, string pnlOtherID, string alertMsg)
    {
        ShowJsHidePanel(page, pnlOtherID);
        ShowJsHideModalPopup(page, pnlModalPopupID);
        ShowJsMSG(page, alertMsg);
        ShowJsShowPanel(page, pnlOtherID);
    }

    public static void ShowJsHideModalPopup(System.Web.UI.Page page, string pnlID)
    {
        string hideModalPopup = "function HideModalPopup() { var idName = '" + pnlID + "'; $('[id$=' + idName + ']').hide(); } HideModalPopup();";
        System.Web.UI.ScriptManager.RegisterStartupScript(page, page.GetType(), "HideModalPopup", hideModalPopup, true);
    }

    public static void ShowJsHidePanel(System.Web.UI.Page page, string pnlID)
    {
        string hidePanel = "function HidePanel() { var idName = '" + pnlID + "'; $('[id$=' + idName + ']').hide(); } HidePanel();";
        System.Web.UI.ScriptManager.RegisterStartupScript(page, page.GetType(), "HidePanel", hidePanel, true);
    }

    public static void ShowJsShowPanel(System.Web.UI.Page page, string pnlID)
    {
        string showPanel = "function ShowPanel() { var idName = '" + pnlID + "'; $('[id$=' + idName + ']').show(); } ShowPanel();";
        System.Web.UI.ScriptManager.RegisterStartupScript(page, page.GetType(), "ShowPanel", showPanel, true);
    }
}