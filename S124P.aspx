<%@ Page Title="" Language="C#" MasterPageFile="~/MasterPage.master" AutoEventWireup="true" CodeFile="S124P.aspx.cs" Inherits="S124P" %>

<%@ Register Src="../UserControl/Pager.ascx" TagName="Pager" TagPrefix="uc2" %>

<asp:Content ID="Content1" ContentPlaceHolderID="contentplaceholderContent" runat="Server">
        <!--Font-awesome-->
    <link href="../CSS/font-awesome-4.7.0/css/font-awesome.css" rel="stylesheet" />

	<link href="../JQUERY/jquery-ui-1.12.1/css/jquery-ui.css" rel="stylesheet" />
	<link href="../JQUERY/select2/select2.css" rel="stylesheet" />
    <link href="../CSS/NewERPStyle.css" rel="stylesheet" />
    <link href="../CSS/ProgressCSS.css" rel="stylesheet" />
    <script src="../JQUERY/jquery-ui-1.12.1/jquery-1.12.4.js"></script>
    <script src="../JQUERY/jquery-ui-1.12.1/jquery-ui-1.12.1.js"></script>
    <script src="../JQUERY/select2/select2.js"></script>
    <script src="../JQUERY/jquery.blockUI.js"></script>
    <script src="../COMMON/SQLDynamicJS.js?20190802"></script>
    <script src="../COMMON/gridviewScroll.min.js"></script>
    <script src="../COMMON/SetGridviewScroll.js"></script>
    <script src="../COMMON/JSqlPopupSearch.js"></script>
    <script src="../COMMON/SetPanelScrollPos.js"></script>
    <script>

        $(document).ready(function () {
            LoadTab();
            SetIFrameHeigh();
            LoadSelect2()
            SetDialogCenter("dialogDetail", "明細資料");
	        BlockUI("main");
        });

        function LoadSelect2() {
            $('#<%= lsbSalesID.ClientID %>').select2({ placeholder: '請選擇', allowClear: true });
                    $('#<%= lsbObject.ClientID %>').select2({ placeholder: '請選擇', allowClear: true });
                }

        function pageLoad() {
<%--            GridviewScroll($('#<%=this.gvContent.ClientID%>'), $("#<%=hfV_gvContent.ClientID%>"), $("#<%=hfH_gvContent.ClientID%>"), 600, 1250, 0, 0);
            GridviewScroll($('#<%=this.gvDialogDetail.ClientID%>'), $("#<%=hfV1_gvDialogDetail.ClientID%>"), $("#<%=hfH1_gvDialogDetail.ClientID%>"), 500, 1150, 0, 0);--%>

            GridviewScroll($('#<%=this.gvSumData.ClientID%>'), $("#<%=hfV_gvSumData.ClientID%>"), $("#<%=hfH_gvSumData.ClientID%>"), 500, 1500, 0, 0);
            GridviewScroll($('#<%=this.gvFailedData.ClientID%>'), $("#<%=hfV_gvFailedData.ClientID%>"), $("#<%=hfH_gvFailedData.ClientID%>"), 500, 1500, 0, 0);
            GridviewScroll($('#<%=this.gvDetailData.ClientID%>'), $("#<%=hfV_gvDetailData.ClientID%>"), $("#<%=hfH_gvDetailData.ClientID%>"), 500, 1500, 0, 0);
        }

	    function SetIFrameHeigh() {
	        var body = document.body,
            html = document.documentElement;
	        var height = $('#<%= table_FormContent.ClientID %>').height() + 30;
	        height = 'IFrameHeigh#' + height;
	        parent.postMessage(height, '*');
	    }

	    var prm = Sys.WebForms.PageRequestManager.getInstance();
        prm.add_endRequest(function () {
            LoadTab();
            SetIFrameHeigh();
            LoadSelect2();
            SetDialogCenter("dialogDetail", "明細資料");
	    });

        //顯示彈出視窗
        function showDialog(id) {
            $('#' + id).dialog("open");
        }

        //關閉彈出視窗
        function closeDialog(id) {
            $('#' + id).dialog("close");
        }

        function LoadTab() {
            var _showTab = $('#<%= selected_tab.ClientID %>').val();

	        $('.abgne_tab').each(function () {
	            var $tab = $(this);
	            var $defaultLi = $('ul.tabs li', $tab).eq(_showTab).addClass('active');
	            $($defaultLi.find('a').attr('href')).siblings().hide();

	            $('ul.tabs li', $tab).click(function () {
	                var $this = $(this),
                        _clickTab = $this.find('a').attr('href');

	                //var newIdx = $("li").index($this);
	                var newIdx = $this.index();
                    $('#<%= selected_tab.ClientID %>').val(newIdx);
                    $this.addClass('active').siblings('.active').removeClass('active');
                    $(_clickTab).stop(false, true).fadeIn().siblings().hide();
                    SetIFrameHeigh();
                    return false;
                }).find('a').focus(function () {
                    this.blur();
                });
            });
                }

	    function BlockUI(elementID) {
	        var prm = Sys.WebForms.PageRequestManager.getInstance();
	        prm.add_beginRequest(function () {

	            $.blockUI({
	                message: '<table align = "center"><tr><td>' +
                     '<img src="../images/loadingAnim.gif"/></td></tr></table>',
	                css: {},
	                //bindEvents: true,
	                overlayCSS: {
	                    backgroundColor: '#000000', opacity: 0.6, border: '3px solid #63B2EB'
	                }
	            });
	        });
	        prm.add_endRequest(function () {
	            $.unblockUI();
	        });
	    }





    </script>
	<style type="text/css">
		input::-webkit-outer-spin-button,
		input::-webkit-inner-spin-button {
			-webkit-appearance: none !important;
			margin: 0;
		}

        .TopMenuButton {
            background-color: transparent;
            border: none;
            color: #484444; /*white;*/
            padding: 6px 26px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            font-weight: 700;
        }

        .TopMenuButton:hover {
            background-color: transparent;
            border: none;
            color: #034af3;/*yellow;*/
            padding: 6px 26px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            font-weight: 700;
        }

        .TopMenuButton_disabled {
            background-color: transparent;
            border: none;
            color: silver;
            padding: 6px 26px;
            text-align: center;
            text-decoration: none;
            display: inline-block;
            font-size: 16px;
            font-weight: 700;
        }

		.grid-header {
			background: #e7ebef url('../images/GridView.header/bg-header-silver.gif') repeat-x left top;
			background-repeat: repeat-x;
			font-weight: bold;
			position: relative;
			top: expression(this.offsetParent.scrollTop);
			z-index: 10;
		}

		ul, li {
			margin: 0;
			padding: 0;
			list-style: none;
		}

		.abgne_tab {
			clear: left;
			width: 1150px;
			margin: 10px 0;
		}

		ul.tabs {
			width: 100%;
			height: 32px;
			border-bottom: 1px solid #999;
			border-left: 1px solid #999;
		}

			ul.tabs li {
				float: left;
				height: 31px;
				line-height: 31px;
				overflow: hidden;
				position: relative;
				margin-bottom: -1px; /* 讓 li 往下移來遮住 ul 的部份 border-bottom */
				border: 1px solid #999;
				border-left: none;
				background: #e1e1e1;
			}

				ul.tabs li a {
					display: block;
					padding: 0 20px;
					color: #000;
					border: 1px solid #fff;
					text-decoration: none;
				}

					ul.tabs li a:hover {
						background: #ccc;
					}

				ul.tabs li.active {
					background: #fff;
					border-bottom: 1px solid #fff;
				}

					ul.tabs li.active a:hover {
						background: #fff;
					}

		div.tab_container {
			clear: left;
			width: 100%;
			border: 1px solid #999;
			border-top: none;
			background: #fff;
		}

			div.tab_container .tab_content {
				padding: 20px;
			}

				div.tab_container .tab_content h2 {
					margin: 0 0 20px;
				}

		.ui-state-disabled {
			display: none;
		}

		.tooltip {
			border-bottom: 1px none #000000;
			color: #000000;
			outline: none;
			cursor: help;
			text-decoration: none;
			position: relative;
			left: 10px;
		}

			.tooltip span {
				margin-left: -999em;
				position: absolute;
			}

			.tooltip:hover span {
				border-radius: 5px 5px;
				-moz-border-radius: 5px;
				-webkit-border-radius: 5px;
				box-shadow: 5px 5px 5px rgba(0, 0, 0, 0.1);
				-webkit-box-shadow: 5px 5px rgba(0, 0, 0, 0.1);
				-moz-box-shadow: 5px 5px rgba(0, 0, 0, 0.1);
				font-family: 'Microsoft JhengHei', Calibri;
				position: absolute;
				left: 1em;
				top: 2em;
				z-index: 99;
				margin-left: 0;
				width: 280px;
			}

			.tooltip:hover img {
				border: 0;
				margin: -10px 0 0 -55px;
				float: left;
				position: absolute;
			}

			.tooltip:hover em {
				font-family: 'Microsoft JhengHei', Calibri;
				font-size: 1.2em;
				font-weight: bold;
				display: block;
				padding: 0.1em 0 0.1em 0;
			}

		.classic {
			padding: 0.8em 1em;
		}

		.custom {
			padding: 0.5em 0.8em 0.8em 1em;
		}

		* html a:hover {
			background: transparent;
		}

		.classic {
			background: #FFFFAA;
			border: 1px solid #FFAD33;
		}

		.critical {
			background: #FFCCAA;
			border: 1px solid #FF3334;
		}

		.help {
			background: #9FDAEE;
			border: 1px solid #2BB0D7;
		}

		.info {
			background: #9FDAEE;
			border: 1px solid #2BB0D7;
		}

	    .warning {
	        background: #FFFFAA;
	        border: 1px solid #FFAD33;
	    }
	</style>
    <asp:ScriptManager ID="ScriptManager1" runat="server"></asp:ScriptManager>

    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
		<ContentTemplate>
            <div id="main">
                <table id="table_FormContent" runat="server" style="width: 1150px; padding-left: 30px;">
                    <tr>
                        <td>
                            <!--表單標題-->
                            <p style="text-align: center; font-family: 'Microsoft JhengHei',Verdana; font-size: 20px; font-weight: 700;">
                                報價成功率報表<br />
                            </p>
                            <!--表單內容-->
                            <table style="width: 100%; font-family: 'Microsoft JhengHei',Verdana; font-size: 14px;">
                                <tr>
                                    <td colspan="2">
                                        <table style="width: 100%; border-collapse: collapse; border: 1px solid #95B8E7;">
                                            <tr style="background-color: #E4EFFF; font-size: 16px; font-weight: 700; line-height: 24px;">
                                                <td style="text-align: left; font-size: 14px; font-family: 'Microsoft JhengHei',Verdana; padding: 4px 7px 2px 4px; border: 1px solid #95B8E7; color: #0E2D5F;">查詢條件</td>
                                            </tr>
                                            <tr>
                                                <td>
                                                    <asp:UpdatePanel ID="upFileCompare" runat="server" UpdateMode="Conditional">
                                                        <ContentTemplate>
                                                            <table style="width: 100%;" class="NewERPFormTable">
                                                                <tr>
                                                                    <td style="text-align: center;">
                                                                        &nbsp;<asp:Label ID="Label18" runat="server" Text="查詢年度"></asp:Label>
                                                                        &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;</td>
                                                                    <td style="text-align: center;">
                                                                        <asp:TextBox ID="txtStartDate" runat="server" TextMode="Date" Width="120px" Visible="False"></asp:TextBox>
                                                                        <asp:TextBox ID="txtEndDate" runat="server" TextMode="Date" Width="120px" Visible="False"></asp:TextBox>
                                                                        <asp:DropDownList ID="ddlYear" runat="server">
                                                                            <asp:ListItem>2022</asp:ListItem>
                                                                            <asp:ListItem>2023</asp:ListItem>
                                                                            <asp:ListItem>2024</asp:ListItem>
                                                                            <asp:ListItem>2025</asp:ListItem>
                                                                            <asp:ListItem>2026</asp:ListItem>
                                                                        </asp:DropDownList>
                                                                    </td>
                                                                    <td>
                                                                        <asp:Label ID="Label7" runat="server" Text="成本據點" Font-Bold="True" />
                                                                    </td>
                                                                    <td>
                                                                        <asp:RadioButtonList ID="rblCostSiteName" runat="server">
                                                                            <asp:ListItem Selected="True">全部</asp:ListItem>
                                                                            <asp:ListItem>平鎮</asp:ListItem>
                                                                            <asp:ListItem>蘇州</asp:ListItem>
                                                                        </asp:RadioButtonList>
                                                                    </td>
                                                                    <td style="text-align: center;">
                                                                        <asp:Label ID="Label19" runat="server" Text="事業部"></asp:Label>
                                                                    </td>
                                                                    <td style="text-align: left;">
                                                                        <asp:RadioButtonList ID="rblBelongSection" runat="server" RepeatColumns="3">
                                                                        </asp:RadioButtonList>
                                                                    </td>
                                                                    <td style="text-align: center;">
                                                                        <asp:Label ID="Label20" runat="server" Text="案件屬性"></asp:Label>
                                                                    </td>
                                                                    <td style="text-align: left;">
                                                                        <asp:CheckBoxList ID="cblSortTypeA" runat="server">
                                                                            <asp:ListItem Value="NEWDESIGN" Selected="True">新設計案</asp:ListItem>
                                                                            <asp:ListItem Value="REORDER" Selected="True">Re-Order</asp:ListItem>
                                                                            <asp:ListItem Value="GERBER" Selected="True">Gerber</asp:ListItem>
                                                                        </asp:CheckBoxList>
                                                                    </td>
                                                                    <td>
                                                                        <asp:Label ID="Label6" runat="server" Text="業務" />
                                                                        <br />
                                                                        <br />
                                                                        <asp:Label ID="Label5" runat="server" Text="客戶" />
                                                                        <br />
                                                                    </td>
                                                                    <td style="text-align: left;">
                                                                        <asp:ListBox ID="lsbSalesID" runat="server" Width="200"></asp:ListBox>
                                                                        <br />
                                                                        <br />
                                                                        <asp:ListBox ID="lsbObject" runat="server" Width="200"></asp:ListBox>
                                                                        <br />
                                                                    </td>
                                                                    <td style="text-align: center;" colspan="10">
                                                                        <asp:Button ID="btnQuery" runat="server" ForeColor="Blue" OnClick="btnQuery_Click" Text="查詢" Width="80px" />
                                                                        <br />
                                                                        <br />
                                                                        <asp:Button ID="btnExcel" runat="server" OnClick="btnExcel_Click" Text="匯出" Width="80px" />
                                                                    </td>
                                                                    <td style="text-align: left; width: 500px;">1.資料為<asp:Label ID="Label8" runat="server" Font-Bold="True" ForeColor="Blue" Text="前一日23:00" Font-Underline="False"></asp:Label>由系統排程更新 
                                                                <br />
                                                                        2.如需最新資料可<asp:Button ID="btnResetData" runat="server" Text="手動更新資料" ForeColor="blue" OnClick="btnResetData_Click" />
                                                                        (約需40-60Sec)
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </ContentTemplate>
                                                    </asp:UpdatePanel>
                                                </td>
                                            </tr>
                                            <tr style="background-color: #E4EFFF; font-size: 16px; font-weight: 700; line-height: 24px;">
                                                <td style="text-align: left; font-size: 14px; font-family: 'Microsoft JhengHei',Verdana; padding: 4px 7px 2px 4px; border: 1px solid #95B8E7; color: #0E2D5F;">查詢結果
                                                </td>
                                            </tr>
                                            <tr style="text-align: center; width: 100%;">
                                                <td style="text-align: center; width: 100%;">
                                                    <table style="width: 100%; text-align: center;">
                                                        <tr style="text-align: center;">
                                                            <td style="text-align: center;"></td>
                                                        </tr>
                                                    </table>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td style="padding-left: 20px; padding-bottom: 10px;">
                                                    <div class="abgne_tab" id="tabs" style="text-align: center; width: 100%;">
                                                        <ul class="tabs">
                                                            <li id="litab1" runat="server"><a href="#tab1">RFQ轉內訂總表</a></li>
                                                            <li id="litab2" runat="server"><a href="#tab2">報價失敗分析</a></li>
                                                            <li id="litab3" runat="server"><a href="#tab3">RFQ明細By客戶</a></li>
                                                        </ul>
                                                        <div class="tab_container">
                                                            <div id="tab1" class="tab_content">
                                                                <asp:GridView ID="gvSumData" runat="server" Align="Center" Width="100%" EmptyDataText="No data to display" OnRowDataBound="gvData_RowDataBound" AutoGenerateColumns="False">
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="事業部" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labBelongSectionName" runat="server" Text='<%# Bind("事業部") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="報價狀態" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labType" runat="server" Text='<%# Bind("狀態") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="1月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab1M" runat="server" Text='<%# Bind("1月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="2月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab2M" runat="server" Text='<%# Bind("2月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="3月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab3M" runat="server" Text='<%# Bind("3月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="4月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab4M" runat="server" Text='<%# Bind("4月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="5月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab5M" runat="server" Text='<%# Bind("5月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="6月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab6M" runat="server" Text='<%# Bind("6月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="7月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab7M" runat="server" Text='<%# Bind("7月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="8月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab8M" runat="server" Text='<%# Bind("8月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="9月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab9M" runat="server" Text='<%# Bind("9月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="10月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab10M" runat="server" Text='<%# Bind("10月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="11月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab11M" runat="server" Text='<%# Bind("11月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="12月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab12M" runat="server" Text='<%# Bind("12月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                    </Columns>
                                                                    <EmptyDataTemplate>
                                                                        No Data To Dispaly                                                         
                                                                    </EmptyDataTemplate>
                                                                    <HeaderStyle CssClass="grid-header"></HeaderStyle>
                                                                </asp:GridView>
                                                                <table>
                                                                    <tr>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfV_gvSumData" runat="server" Value="0" />
                                                                        </td>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfH_gvSumData" runat="server" Value="0" />
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                            <div id="tab2" class="tab_content">
                                                                <asp:GridView ID="gvFailedData" runat="server" Align="Center" Width="100%" EmptyDataText="No data to display" OnRowDataBound="gvData_RowDataBound" AutoGenerateColumns="False">
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="事業部" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labBelongSectionName" runat="server" Text='<%# Bind("事業部") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="報價失敗原因" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labType" runat="server" Text='<%# Bind("報價失敗原因") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="1月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab1M" runat="server" Text='<%# Bind("1月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="2月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab2M" runat="server" Text='<%# Bind("2月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="3月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab3M" runat="server" Text='<%# Bind("3月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="4月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab4M" runat="server" Text='<%# Bind("4月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="5月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab5M" runat="server" Text='<%# Bind("5月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="6月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab6M" runat="server" Text='<%# Bind("6月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="7月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab7M" runat="server" Text='<%# Bind("7月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="8月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab8M" runat="server" Text='<%# Bind("8月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="9月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab9M" runat="server" Text='<%# Bind("9月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="10月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab10M" runat="server" Text='<%# Bind("10月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="11月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab11M" runat="server" Text='<%# Bind("11月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="12月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab12M" runat="server" Text='<%# Bind("12月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                    </Columns>
                                                                    <EmptyDataTemplate>
                                                                        No Data To Dispaly                                                         
                                                                    </EmptyDataTemplate>
                                                                    <HeaderStyle CssClass="grid-header"></HeaderStyle>
                                                                </asp:GridView>
                                                                <table>
                                                                    <tr>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfV_gvFailedData" runat="server" Value="0" />
                                                                        </td>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfH_gvFailedData" runat="server" Value="0" />
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                            <div id="tab3" class="tab_content">
                                                                <asp:GridView ID="gvDetailData" runat="server" Align="Center" Width="100%" EmptyDataText="No data to display" OnRowDataBound="gvData_RowDataBound" AutoGenerateColumns="False">
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="事業部" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labBelongSectionName" runat="server" Text='<%# Bind("事業部") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="客戶" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labType" runat="server" Text='<%# Bind("客戶簡稱") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="1月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab1M" runat="server" Text='<%# Bind("1月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="2月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab2M" runat="server" Text='<%# Bind("2月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="3月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab3M" runat="server" Text='<%# Bind("3月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="4月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab4M" runat="server" Text='<%# Bind("4月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="5月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab5M" runat="server" Text='<%# Bind("5月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="6月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab6M" runat="server" Text='<%# Bind("6月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="7月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab7M" runat="server" Text='<%# Bind("7月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="8月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab8M" runat="server" Text='<%# Bind("8月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="9月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab9M" runat="server" Text='<%# Bind("9月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="10月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab10M" runat="server" Text='<%# Bind("10月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="11月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab11M" runat="server" Text='<%# Bind("11月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="12月" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="lab12M" runat="server" Text='<%# Bind("12月") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                    </Columns>
                                                                    <EmptyDataTemplate>
                                                                        No Data To Dispaly                                                         
                                                                    </EmptyDataTemplate>
                                                                    <HeaderStyle CssClass="grid-header"></HeaderStyle>
                                                                </asp:GridView>
                                                                <table>
                                                                    <tr>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfV_gvDetailData" runat="server" Value="0" />
                                                                        </td>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfH_gvDetailData" runat="server" Value="0" />
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                            <div id="tab4" class="tab_content">
                                                                <asp:GridView ID="gvContent" runat="server" Align="Center" Width="100%" EmptyDataText="No data to display" OnRowDataBound="gvContent_RowDataBound" AutoGenerateColumns="False" Visible="False">
                                                                    <Columns>
                                                                        <asp:TemplateField HeaderText="事業部" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labBelongSectionName" runat="server" Text='<%# Bind("事業部") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="類別" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labType" runat="server" Text='<%# Bind("類別") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="總計" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labTotal" runat="server" Text='<%# Bind("總計") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="報價成功" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Button ID="btnSuccess" runat="server" Font-Bold="true" ForeColor="Blue" BorderWidth="0" BackColor="White" Text='<%# Bind("報價成功") %>' OnClick="btnStatus_Click" />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="待客戶回覆" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Button ID="btnWaiting" runat="server" ForeColor="Green" BorderWidth="0" BackColor="White" Text='<%# Bind("待客戶回覆") %>' OnClick="btnStatus_Click" />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="報價失敗" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Button ID="btnFailed" runat="server" ForeColor="Red" BorderWidth="0" BackColor="White" Text='<%# Bind("報價失敗") %>' OnClick="btnStatus_Click" />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                        <asp:TemplateField HeaderText="報價成功率%" ItemStyle-HorizontalAlign="Center">
                                                                            <ItemTemplate>
                                                                                <asp:Label ID="labSuccessRate" runat="server" Font-Bold="true" ForeColor="Blue" Text='<%# Bind("報價成功率") %>' />
                                                                            </ItemTemplate>
                                                                            <ItemStyle HorizontalAlign="Center" />
                                                                        </asp:TemplateField>
                                                                    </Columns>
                                                                    <EmptyDataTemplate>
                                                                        No Data To Dispaly                                                         
                                                                    </EmptyDataTemplate>
                                                                    <HeaderStyle CssClass="grid-header"></HeaderStyle>
                                                                </asp:GridView>
                                                                <table>
                                                                    <tr>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfV_gvContent" runat="server" Value="0" />
                                                                        </td>
                                                                        <td>
                                                                            <asp:HiddenField ID="hfH_gvContent" runat="server" Value="0" />
                                                                        </td>
                                                                    </tr>
                                                                </table>
                                                            </div>
                                                        </div>
                                                </td>
                                            </tr>
                                        </table>
                                    </td>
                                </tr>
                            </table>
                            <asp:HiddenField ID="selected_tab" runat="server" />
                        </td>
                    </tr>
                </table>
            </div>
        </ContentTemplate>
        <Triggers>
            <asp:PostBackTrigger ControlID="btnExcel" />
        </Triggers>
    </asp:UpdatePanel>
    <br />
    <div id="dialogDetail" style="width: 1050px;">
        <asp:UpdatePanel runat="server" Visible="false">
            <ContentTemplate>
                <table style="width: 100%;" class="NewERPFormTable">
                    <tr>
                        <td colspan="2">事業部：<asp:Label ID="labPopBelongSection" runat="server" ForeColor="Blue"></asp:Label>
                            <br />
                            報價狀態：<asp:Label ID="labPopOrderStatus" runat="server" ForeColor="Blue"></asp:Label>
                            <br />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2">
                            <asp:GridView ID="gvDialogDetail" runat="server" Width="100%" EmptyDataText="No data to display" ShowHeaderWhenEmpty="True" OnRowDataBound="gvDialogDetail_RowDataBound">
                                <EmptyDataTemplate>
                                    No Data To Display
                                </EmptyDataTemplate>
                                <HeaderStyle CssClass="grid-header"></HeaderStyle>
                                <RowStyle HorizontalAlign="Center" VerticalAlign="Middle" />
                            </asp:GridView>
                            <asp:HiddenField ID="hfV1_gvDialogDetail" runat="server" Value="0" />
                            <asp:HiddenField ID="hfH1_gvDialogDetail" runat="server" Value="0" />
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: center;">
                            <asp:Button runat="server" Text="關閉" Width="80" OnClientClick="closeDialog('dialogDetail');" />
                        </td>
                    </tr>
                </table>
            </ContentTemplate>
        </asp:UpdatePanel>
    </div>
</asp:Content>

