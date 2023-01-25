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

<script src='../js/jquery.min.js'></script>
<script src="../js/bootstrap.min.js"></script>
<script src="../js/jquery-ui.js"></script>

<script>
function lfn_init() {
    document.getElementById("txt_asset_cd").value = "";
    document.getElementById("txt_asset_nm").value = "";
    document.getElementById("txt_spec").value = "";
    document.getElementById("txt_asset_class_nm").value = "";
    document.getElementById("txt_mgmt_class_nm").value = "";
    document.getElementById("txt_make_vend_nm").value = "";
    document.getElementById("txt_make_dt").value = "";
    document.getElementById("txt_puchase_vend_nm").value = "";
    document.getElementById("txt_s_no").value = "";
    document.getElementById("txt_buy_dt").value = "";
    document.getElementById("txt_purchase_vend_tel").value = "";
    document.getElementById("txt_person_in_charge").value = "";
    document.getElementById("txt_buy_amount").value = "";
    document.getElementById("txt_specialty").value = "";
    document.getElementById("txt_asset_status_nm").value = "";
    document.getElementById("txt_mgmt_com_nm").value = "";
    document.getElementById("txt_manage_dept_nm").value = "";
    document.getElementById("txt_manager_emp_nm").value = "";
    document.getElementById("txt_manager_emp_nm_b").value = "";
}

$(function() {
    //lfn_init();

    $("#grd01_body").empty();

    $("#txt_asset_cd").keyup(function(event) {
        if (event.keyCode == 13) {
            var f = document.frm;
            if (f.txt_asset_cd.value == "") {
                alert("자산번호를 입력하여 주시기 바랍니다.");
                f.txt_asset_cd.focus();
                return;
            }
            f.txt_hidden_flag.value="asscd_cd_search";
            f.action = "menu03_proc.asp"
            f.method = "post";
            f.target = "tmp";
            f.submit();
        }
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
    $("#datepicker1").datepicker({
        showOn: "button",
        buttonImage: "../img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date"
    });
    $("#datepicker2").datepicker({
        showOn: "button",
        buttonImage: "../img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date"
    });
});
</script>
</head>
<body onload="document.getElementById('txt_asset_cd').focus()">

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
        <h3>자산 현황</h3>
    </div>
</div>
<!-- //header -->

<form name="frm" onkeydown="if (event.keyCode == 13) return false;">
<!-- container -->
<div class="container">

    <div class="sub-inquiry">
        <div class="inquiry-area">
            <table class="table table-condensed">
                <colgroup>
                    <col style="width:79px;">
                    <col>
                </colgroup>
                <tbody>
                <tr>
                    <th>자산번호</th>
                    <td><input type="text" value="" class="alone" id="txt_asset_cd" name="txt_asset_cd"></td>
                </tr>
                <tr>
                    <th>자산</th>
                    <td><input type="text" value="" class="alone" readonly="readonly" id="txt_asset_nm" name="txt_asset_nm"></td>
                </tr>
                </tbody>
            </table>
        </div>
    </div>

    <!-- 스크롤 바디 -->
    <div style="margin-top:15px; height:350px; overflow-y:auto;">

        <div class="sub-inquiry mt-0">
            <div class="inquiry-area">
                <table class="table table-condensed">
                    <colgroup>
                        <col style="width:79px;">
                        <col>
                    </colgroup>
                    <tbody>
                    <tr>
                        <th>규격</th>
                        <td>
                            <textarea rows="3" id="txt_spec" name="txt_spec" readonly></textarea>
                        </td>
                    </tr>
                    <tr>
                        <th>자산구분</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_asset_class_nm" name="txt_asset_class_nm"></td>
                    </tr>
                    <tr>
                        <th>관리구분</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_mgmt_class_nm" name="txt_mgmt_class_nm"></td>
                    </tr>
                    <tr>
                        <th>제조사</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_make_vend_nm" name="txt_make_vend_nm"></td>
                    </tr>
                    <tr>
                        <th>제작일</th>
                        <td><input type="text" class="alone" readonly="readonly" id="txt_make_dt" name="txt_make_dt"></td>
                    </tr>
                    <tr>
                        <th>구매처</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_puchase_vend_nm" name="txt_puchase_vend_nm"></td>
                    </tr>
                    <tr>
                        <th>Serial No</th>
                        <td><input type="text" class="alone" readonly="readonly" id="txt_s_no" name="txt_s_no"></td>
                    </tr>
                    <tr>
                        <th>구매일</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_buy_dt" name="txt_buy_dt"></td>
                    </tr>
                    <tr>
                        <th>구매처 전화</th>
                        <td><input type="text" class="alone" readonly="readonly" id="txt_purchase_vend_tel" name="txt_purchase_vend_tel"></td>
                    </tr>
                    <tr>
                        <th>구매처담당</th>
                        <td><input type="text" class="alone" readonly="readonly" id="txt_person_in_charge" name="txt_person_in_charge"></td>
                    </tr>
                    <tr>
                        <th>구매금액</th>
                        <td><input type="text" class="alone" readonly="readonly" id="txt_buy_amount" name="txt_buy_amount"></td>
                    </tr>
                    <tr>
                        <th>비고</th>
                        <td><input type="text" class="alone" readonly="readonly" id="txt_specialty" name="txt_specialty"></td>
                    </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <div class="sub-inquiry">
            <div class="inquiry-area">
                <table class="table table-condensed">
                    <colgroup>
                        <col style="width:79px;">
                        <col>
                    </colgroup>
                    <tbody>
                    <tr>
                        <th>상태</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_asset_status_nm" name="txt_asset_status_nm"></td>
                    </tr>
                    <tr>
                        <th>관리회사</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_mgmt_com_nm" name="txt_mgmt_com_nm"></td>
                    </tr>
                    <tr>
                        <th>사용부서</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_manage_dept_nm" name="txt_manage_dept_nm"></td>
                    </tr>
                    <tr>
                        <th>구매법인</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_purchase_co_nm" name="txt_purchase_co_nm"></td>
                    </tr>
                    <tr>
                        <th>관리자 (정)</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_manager_emp_nm" name="txt_manager_emp_nm"></td>
                    </tr>
                    <tr>
                        <th>관리자 (부)</th>
                        <td><input type="text" value="" class="alone" readonly="readonly" id="txt_manager_emp_nm_b" name="txt_manager_emp_nm_b"></td>
                    </tr>
                    </tbody>
                </table>
            </div>
        </div>

        <!-- 데이터 테이블 -->
        <div class="sub-inquiry grid-table auto">
            <div class="title-area">
                <h4>변동이력</h4>
            </div>
            <div class="inquiry-area">
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
                    <thead>
                    <tr>
                        <th>변동일</th>
                        <th>변동</th>
                        <th>회사</th>
                        <th>사용부서</th>
                        <th>구매법인</th>
                        <th>관리자(정)</th>
                        <th>관리자(부)</th>
                        <th>생성구분</th>
                        <th>비고</th>
                    </tr>
                    </thead>
                    <tbody id="grd01_body">
                    <tr>
                        <td class="tcenter"></td>
                        <td class="tcenter">취득</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td class="tcenter"></td>
                        <td class="tcenter"></td>
                        <td class="tcenter">정상</td>
                        <td></td>
                    </tr>
                    <tr>
                        <td class="tcenter"></td>
                        <td class="tcenter">취득</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td class="tcenter"></td>
                        <td class="tcenter"></td>
                        <td class="tcenter">정상</td>
                        <td></td>
                    </tr>
                    <tr>
                        <td class="tcenter"></td>
                        <td class="tcenter">취득</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td class="tcenter"></td>
                        <td class="tcenter"></td>
                        <td class="tcenter">정상</td>
                        <td></td>
                    </tr>
                    <tr>
                        <td class="tcenter"></td>
                        <td class="tcenter">취득</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td class="tcenter"></td>
                        <td class="tcenter"></td>
                        <td class="tcenter">정상</td>
                        <td></td>
                    </tr>
                    <tr>
                        <td class="tcenter"></td>
                        <td class="tcenter">취득</td>
                        <td></td>
                        <td></td>
                        <td></td>
                        <td class="tcenter"></td>
                        <td class="tcenter"></td>
                        <td class="tcenter">정상</td>
                        <td></td>
                    </tr>
                    </tbody>
                </table>
            </div>
        </div>
        <!-- //데이터 테이블 -->

    </div>
    <!-- //스크롤 바디 -->

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
