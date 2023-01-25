<% @Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="UTF-8"
Session.CodePage="65001"
Response.CodePage="65001"
Session.abandon
%>
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
<title>TRF 자산관리 시스템 01</title>

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
<script>
    $(document).ready(function() {
        $("#edt_user_id").keyup(function(event) {
            if (event.keyCode == 13) document.getElementById("edt_user_pw").focus();
        });
        $("#edt_user_pw").keyup(function(event) {
            if (event.keyCode == 13) lfn_login(document.frm);
        });
        $(".input-id input").keyup(function() {
            if ($("#edt_user_id").val().length > 0) {
                $(".input-id").addClass("eraser");
                $(".input-id > .btn-eraser").show();
            } else {
                $(".input-id").removeClass("eraser");
                $(".input-id > .btn-eraser").hide();
            }
        });
        $(".input-pw input").keyup(function() {
            if ($("#edt_user_pw").val().length > 0) {
                $(".input-pw").addClass("eraser");
                $(".input-pw > .btn-eraser").show();
            } else {
                $(".input-pw").removeClass("eraser");
                $(".input-pw > .btn-eraser").hide();
            }
        });
        $("#id_clear").click(function() {
            $(".input-id").removeClass("eraser");
            $(".input-id > .btn-eraser").hide();
            $("#edt_user_id").focus().val("");
        });
        $("#pw_clear").click(function() {
            $(".input-pw").removeClass("eraser");
            $(".input-pw > .btn-eraser").hide();
            $("#edt_user_pw").focus().val("");
        });
    });

    function lfn_login(f) {
        if (f.edt_user_id.value == "") {
            alert("아이디를 입력하여 주시기 바랍니다.");
            f.edt_user_id.focus();
            return;
        } else if (f.edt_user_pw.value == "") {
            alert("비밀번호를 입력하여 주시기 바랍니다.");
            f.edt_user_pw.focus();
            return;
        }

        f.target = "tmp";
        f.action = "common/login_chk.asp";
        f.method = "post";
        f.submit();
    }
</script>
</head>

<body onload="document.frm.edt_user_id.focus()">
<!-- container -->
<div class="container">
    <form class="form-signin" name="frm">
        <h1 class="form-signin-heading mb-20">TRF 자산관리 시스템</h1>
        <span class="input-id"><input type="text" class="input-block-level" placeholder="아이디" name="edt_user_id" id="edt_user_id" onfocus="this.select()" tabindex="1"><button type="button" id="id_clear" class="btn-eraser" tabindex="-1"></button></span>
        <span class="input-pw"><input type="password" class="input-block-level" placeholder="비밀번호" name="edt_user_pw" id="edt_user_pw" tabindex="2"><button type="button" id="pw_clear" class="btn-eraser" tabindex="-1"></button></span>
        <input type="hidden" name="conn" value="login">
        <input type="button" value="로그인" onclick="lfn_login(this.form)" class="btn btn-large btn-primary btn-sign-in" tabindex="3">
    </form>
</div>
<iframe src="" name="tmp" width="0" height="0" style="display:none"></iframe>
<!-- //container -->
</body>
</html>
