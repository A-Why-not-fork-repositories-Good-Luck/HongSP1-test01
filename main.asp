<% @Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="UTF-8"
Session.CodePage="65001"
Response.CodePage="65001"
%>
<!-- #include file="common/auth_chk.asp" -->
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
<link href="css/bootstrap.css" rel="stylesheet">
<link href="css/bootstrap-responsive.css" rel="stylesheet">
<link href="css/common.css" rel="stylesheet">
<link href="css/jquery-ui.css" rel="stylesheet">

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

<script src='js/jquery.min.js'></script>
<script src="js/bootstrap.min.js"></script>
<script src="js/jquery-ui.js"></script>

<script>
$(function() {
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
        buttonImage: "img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date"
    });
    $("#datepicker2" ).datepicker({
        showOn: "button",
        buttonImage: "img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date"
    });
});
</script>

</head>

<body>

<!-- 전체메뉴 -->
<div id="menu-container">
    <!-- menu-wrapper -->
    <ul class="menu-list accordion">
        <li><a href="menu/menu01.asp">자산변동등록</a></li>
        <li><a href="menu/menu02.asp">자산실사등록</a></li>
        <li><a href="menu/menu03.asp">자산정보조회</a></li>
        <li><a href="menu/menu04.asp">실사차이조회</a></li>
    </ul>
    <!-- menu-list accordion-->
</div>
<!-- //전체메뉴 -->

<!-- header -->
<div class="header main">
    <h1 class="blind">TRF 자산관리 시스템</h1>
    <div class="header-level-01">
        <h2><a href="main.asp" title="메인페이지로 이동" class="btn-home">TRF 자산관리 시스템</a></h2>
        <!-- <p><a href="#" title="전체메뉴 펼침" id="hamburger-menu" class="btn-menu"></a></p> -->
        <div id="menu-wrapper">
            <button id="hamburger-menu" title="전체메뉴 펼침"><span></span><span></span><span></span></button>
        </div>
    </div>
</div>
<!-- //header -->

<!-- container -->
<div class="container">
    <ul class="list-main">
        <li><a href="menu/menu01.asp">자산변동등록</a></li>
        <li><a href="menu/menu02.asp">자산실사등록</a></li>
        <li><a href="menu/menu03.asp">자산정보조회</a></li>
        <li><a href="menu/menu04.asp">실사차이조회</a></li>
    </ul>
</div>
<iframe src="" name="tmp" width="0" height="0" style="display:none"></iframe>
<!-- //container -->

</body>
</html>
