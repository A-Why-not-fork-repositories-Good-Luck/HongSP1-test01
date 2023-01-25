<% @Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="UTF-8"
Session.CodePage="65001"
Response.CodePage="65001"
%>
<!-- #include file="../common/auth_chk.asp" -->
<!-- #include file="../common/db_conn.asp" -->
<!DOCTYPE html>
<html lang="ko">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<meta charset="utf-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge,chrome=1">
<meta http-equiv="Cache-Control" content="no-cache">
<meta http-equiv="Pragma" content="no-cache">
<meta http-equiv="Expires" content="-1">
<meta name="description" content="TRF 자산관리 시스템">
<meta property="og:description" content="TRF 자산관리 시스템">
<meta name="viewport" content="viewport-fit=cover, width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no, target-densitydpi=medium-dpi">
<title>TRF 자산관리 시스템</title>

<!-- Le styles -->
<link href="../css/bootstrap.css" rel="stylesheet">
<link href="../css/bootstrap-responsive.css" rel="stylesheet">
<link href="../css/common.css" rel="stylesheet">
<link href="../css/jquery-ui.css" rel="stylesheet">

<!-- HTML5 shim, for IE6-8 support of HTML5 elements -->
<!--[if lt IE 9]>
    <script src="js/html5shiv.js"></script>
<![endif]-->

<!-- Fav and touch icons -->
<link rel="apple-touch-icon-precomposed" sizes="144x144" href="ico/apple-touch-icon-144-precomposed.png">
<link rel="apple-touch-icon-precomposed" sizes="114x114" href="ico/apple-touch-icon-114-precomposed.png">
<link rel="apple-touch-icon-precomposed" sizes="72x72" href="ico/apple-touch-icon-72-precomposed.png">
<link rel="apple-touch-icon-precomposed" href="ico/apple-touch-icon-57-precomposed.png">
<link rel="shortcut icon" href="ico/favicon.png">

<!-- google web font -->
<!-- <link rel="preconnect" href="https://fonts.gstatic.com">
<link href="https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@100;300;400;500;700;900&display=swap" rel="stylesheet"> -->

<script src="http://code.jquery.com/jquery-latest.js"></script>
<!--script src="../js/jquery.min.js"></script-->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/jquery-ui.js"></script>

<script>
function lfn_getday() {
    var now = new Date();
    var year = now.getFullYear();
    var month = (now.getMonth() + 1) >=10 ? (now.getMonth() + 1) : "0" + (now.getMonth() + 1);
    var date  = (now.getDate()) >= 10 ? (now.getDate()) : "0" + (now.getDate());

    return year + "/" + month;
}

$(window).load(function() {
<%
    If Not (session("system_admin_flag") = "1" Or session("notice_admin_flag") = "1") Then
        Response.write "$(""#cbo_mgmt_com"").val(""" & session("mgmt_com_cd") & """);"
        Response.write "$(""#cbo_mgmt_com"").attr(""disabled"", ""true"");"
    End if
%>
});

$(function() {
    $("#grd01_body").empty();

    $("#btn_search").click(function() {
        $("#txt_cbo_mgmt_com").val($("#cbo_mgmt_com").val());
        var f = document.frm;
        f.txt_hidden_flag.value="btn_search";
        f.action = "menu04_proc.asp"
        f.method = "post";
        f.target = "tmp";
        f.submit();
    });
    function slideMenu() {
        var activeState = $("#menu-container .menu-list").hasClass("active");
        $("#menu-container .menu-list").animate({left: activeState ? "0%" : "-100%"}, 400);
    }
    $("#menu-wrapper").click(function(event) {
        event.stopPropagation();
        $("#hamburger-menu").toggleClass("open");
        $("#menu-container .menu-list").toggleClass("active");
        slideMenu();

        $("body").toggleClass("overflow-hidden");
    });

    $(".menu-list").find(".accordion-toggle").click(function() {
        $(this).next().toggleClass("open").slideToggle("fast");
        $(this).toggleClass("active-tab").find(".menu-link").toggleClass("active");

        $(".menu-list .accordion-content").not($(this).next()).slideUp("fast").removeClass("open");
        $(".menu-list .accordion-toggle").not(jQuery(this)).removeClass("active-tab").find(".menu-link").removeClass("active");
    });
});
</script>

<script>
$(document).ready(function(){
    $(".btn-open").click(function(){
        $(".hidden-table").removeClass("none");
        $(".btn-open-tr").hide();
    });
    $(".btn-close").click(function(){
        $(".hidden-table").addClass("none");
        $(".btn-open-tr").show();
    });
});
</script>

<script>
$(function() {
    $("#datepicker1" ).datepicker({
        showOn: "button",
        buttonImage: "../img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date",
        dateFormat: "yy/mm"
    });
    $("#datepicker1").val(lfn_getday());
    $("#datepicker2" ).datepicker({
        showOn: "button",
        buttonImage: "../img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date",
        dateFormat: "yy/mm"
    });
});
</script>

</head>

<body>

<!-- 전체메뉴 -->
<div id="menu-container">
    <!-- menu-wrapper -->
    <ul class="menu-list accordion">
        <li><a href="menu01.asp">자산변동등록</a></li>
        <li><a href="menu02.asp">자산실사등록</a></li>
        <li><a href="menu03.asp">자산정보조회</a></li>
        <li><a href="menu04.asp">실사차이조회</a></li>
    </ul>
    <!-- menu-list accordion-->
</div>
<!-- //전체메뉴 -->

<!-- header -->
<div class="header">
    <h1 class="blind">TRF 자산관리 시스템</h1>
    <div class="header-level-01">
        <h2><a href="../main.asp" title="메인페이지로 이동" class="btn-home">TRF 자산관리 시스템</a></h2>
        <div id="menu-wrapper">
            <button id="hamburger-menu" title="전체메뉴 펼침"><span></span><span></span><span></span></button>
        </div>
    </div>
    <div class="header-level-02">
        <a href="../main.asp" title="이전페이지로 이동" class="btn-prev"></a>
        <h3>재고실사 차이 조회</h3>
    </div>
</div>
<!-- //header -->

<form name="frm" onkeydown="if (event.keyCode == 13) return false;">
<!-- container -->
<div class="container">

    <div class="sub-inquiry">
        <div class="title-area">
            <h4>재고실사 차이 조회</h4>
        </div>
        <div class="inquiry-area">

            <table class="table table-condensed">
                <colgroup>
                    <col style="width:79px;">
                    <col>
                </colgroup>
                <tbody>
                <tr>
                    <th>년월</th>
                    <td><input type="text" id="datepicker1" name="txt_month" readonly="readonly"></td>
                </tr>
                <tr>
                    <th>관리회사</th>
                    <td>
                        <select id="cbo_mgmt_com" name="cbo_mgmt_com">
                            <option selected>회로/모듈</option>
                            <option>옵션</option>
                        </select>
                        <input type="hidden" id="txt_cbo_mgmt_com" name="txt_cbo_mgmt_com">
                    </td>
                </tr>
                <tr>
                    <th>관리부서</th>
                    <td>
                        <select id="cbo_mgmt_class" name="cbo_mgmt_class">
                            <option selected>총무</option>
                            <option>옵션</option>
                        </select>
                    </td>
                </tr>
                </tbody>
            </table>

            <div class="tcenter mt-10">
                <div class="btn-bottom">
                    <button class="btn btn-success" id="btn_search">검색</button>
                </div>
            </div>

        </div>
    </div>

    <!-- 데이터 테이블 -->
    <div class="sub-inquiry grid-table">
        <div class="inquiry-area" style="overflow:auto;">
            <table class="table table-condensed" style="width:1000px;" id="grd01">
                <colgroup>
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                    <col style="width:;">
                <thead>
                <tr>
                    <th rowspan="2">자산번호</th>
                    <th rowspan="2">자산</th>
                    <th colspan="5">재고</th>
                    <th colspan="5">실사</th>
                </tr>
                <tr>
                    <th>회사</th>
                    <th>관리부서</th>
                    <th>구매법인</th>
                    <th>관리자(정)</th>
                    <th>관리자(부)</th>
                    <th>회사</th>
                    <th>관리부서</th>
                    <th>구매법인</th>
                    <th>관리자(정)</th>
                    <th>관리자(부)</th>
                </tr>
                </thead>
                <tbody id="grd01_body">
                <tr>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                    <td></td>
                </tr>
                </tbody>
            </table>
        </div>
    </div>
    <!-- //데이터 테이블 -->

    <div class="btn-bottom">
        <button class="btn btn-info" type="button" onclick="javascript:location.replace('../main.asp')">닫기</button>
    </div>

</div>
<!-- //container -->

<input type="hidden" id="txt_hidden_flag" name="txt_hidden_flag">
</form>
<iframe src="" name="tmp" width="0" height="0" style="display:none"></iframe>
</body>
</html>
<%
    Dim strSQL, cnt

    strSQL = ""
    strSQL = strSQL & " SELECT cd_data AS cdcode "
    strSQL = strSQL & "      , cd_nm_loc as cdname "
    strSQL = strSQL & "   FROM mco_code "
    strSQL = strSQL & "  WHERE com_cd = N'" & session("com_cd") & "' "
    strSQL = strSQL & "    AND cd_class = N'ASMGCM' "
    strSQL = strSQL & "    AND cd_data <> N'000000' "
    strSQL = strSQL & "    AND use_flag = N'1' "

    objRS.open strSQL, objDBConn
    If Not (objRS.eof Or objRS.bof) Then
        Response.write "<script>"
        Response.write "var v_cbo_mgmt_com = document.frm.cbo_mgmt_com;"
        Response.write "v_cbo_mgmt_com.options.length=0;"
        Response.write "v_cbo_mgmt_com.options[0] = new Option('선택하세요', '');"
        cnt = 1
        While Not objRS.eof
            Response.write "v_cbo_mgmt_com.options[" & cnt & "] = new Option('" & objRS("cdname") & "', '" & objRS("cdcode") & "');"
            objRS.movenext
            cnt = cnt + 1
        Wend
        Response.write "</script>"
    End If
    objRS.Close

    strSQL = ""
    strSQL = strSQL & " SELECT cd_data AS cdcode "
    strSQL = strSQL & "      , cd_nm_loc as cdname "
    strSQL = strSQL & "   FROM mco_code "
    strSQL = strSQL & "  WHERE com_cd = N'" & session("com_cd") & "' "
    strSQL = strSQL & "    AND cd_class = N'ASMGCL' "
    strSQL = strSQL & "    AND cd_data <> N'000000' "
    strSQL = strSQL & "    AND use_flag = N'1' "

    objRS.open strSQL, objDBConn
    If Not (objRS.eof Or objRS.bof) Then
        Response.write "<script>"
        Response.write "var v_cbo_mgmt_class = document.frm.cbo_mgmt_class;"
        Response.write "v_cbo_mgmt_class.options.length=0;"
        Response.write "v_cbo_mgmt_class.options[0] = new Option('선택하세요', '');"
        cnt = 1
        While Not objRS.eof
            Response.write "v_cbo_mgmt_class.options[" & cnt & "] = new Option('" & objRS("cdname") & "', '" & objRS("cdcode") & "');"
            objRS.movenext
            cnt = cnt + 1
        Wend
        Response.write "</script>"
    End If
    objRS.Close: objDBConn.Close
%>
