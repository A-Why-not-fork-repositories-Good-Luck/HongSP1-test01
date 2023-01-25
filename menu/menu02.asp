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

<script type="text/javascript" src="http://code.jquery.com/jquery-latest.js"></script>
<!--script src="../js/jquery.min.js"></script-->
<script src="../js/bootstrap.min.js"></script>
<script src="../js/jquery-ui.js"></script>

<script>
function lfn_getday() {
    var now = new Date();
    var year = now.getFullYear();
    var month = (now.getMonth() + 1) >=10 ? (now.getMonth() + 1) : "0" + (now.getMonth() + 1);
    var date  = (now.getDate()) >= 10 ? (now.getDate()) : "0" + (now.getDate());

    return year + "/" + month + "/" + date;
}
function lfn_modal_emp_search_a_rtn(a, b) {
    $("#myModal2").modal("hide");
    document.getElementById("txt_a_manager_emp_id").value = a;
    document.getElementById("txt_a_manager_emp_nm").value = b;
}
function lfn_modal_emp_search_b_rtn(a, b) {
    $("#myModal3").modal("hide");
    document.getElementById("txt_b_manager_emp_id").value = a;
    document.getElementById("txt_b_manager_emp_nm").value = b;
}

function lfn_mgmt_com_change_event(f) {
    f.txt_hidden_flag.value="mgmt_com_change";
    f.action = "menu02_proc.asp"
    f.method = "post";
    f.target = "tmp";
    f.submit();
}

$(window).load(function() {
<%
    If Not (session("system_admin_flag") = "1" Or session("notice_admin_flag") = "1") Then
        Response.write "$(""#cbo_mgmt_com"").val(""" & session("mgmt_com_cd") & """);"
        Response.write "lfn_mgmt_com_change_event(document.frm);"
        Response.write "$(""#cbo_mgmt_com"").attr(""disabled"", ""true"");"
    End if
%>
});

$(function() {
    lfn_mgmt_com_change_event(document.frm);

    $(".middle-box").css({"background-color":"#f4f8fb"});
    $("#lbl_result").text("");

    $("#cbo_mgmt_com").change(function() {
        lfn_mgmt_com_change_event(document.frm);
    });

    $("#btn_modal_emp_search_a").click(function() {
        var f = document.frm;
        f.txt_hidden_flag.value="modal_emp_search_a";
        f.action = "menu01_proc.asp"
        f.method = "post";
        f.target = "tmp";
        f.submit();
    });

    $("#btn_modal_emp_search_b").click(function() {
        var f = document.frm;
        f.txt_hidden_flag.value="modal_emp_search_b";
        f.action = "menu01_proc.asp"
        f.method = "post";
        f.target = "tmp";
        f.submit();
    });

    function lfn_asset_search() {
        var f = document.frm;
        if (f.txt_asset_cd.value == "") {
            alert("자산번호는 필수 입력 항목입니다.");
            f.txt_asset_cd.focus();
            return;
        } else if (f.datepicker1.value == "") {
            alert("실사일은 필수 입력 항목입니다.");
            f.txt_asset_cd.value = "";
            f.datepicker1.focus();
            return;
        } else if (f.cbo_mgmt_com.value == "") {
            alert("관리회사는 필수 입력 항목입니다.");
            f.txt_asset_cd.value = "";
            f.cbo_mgmt_com.focus();
            return;
        } else if (f.cbo_manage_dept.value == "") {
            alert("사용부서는 필수 입력 항목입니다.");
            f.txt_asset_cd.value = "";
            f.cbo_manage_dept.focus();
            return;
        }

        $("#txt_hidden_flag").val("asscd_cd_search");
        $.ajax({
            url: "menu02_proc.asp",
            data: $("#frm").serialize(),
            method: "POST",
            dataType: "json"
        })
        .done(function(json) {
            if (json.cnt == 0) {
                alert('자산번호를 확인하시기 바랍니다.');
                document.getElementById('txt_asset_nm').value = '';
                document.getElementById('txt_asset_status_nm').value = '';
                document.getElementById('txt_mgmt_com_cd').value = '';
                document.getElementById('txt_mgmt_com_nm').value = '';
                document.getElementById('txt_manage_dept_id').value = '';
                document.getElementById('txt_manage_dept_nm').value = '';
                document.getElementById('txt_purchase_co').value = '';
                document.getElementById('txt_purchase_co_nm').value = '';
                document.getElementById('txt_manager_emp_id').value = '';
                document.getElementById('txt_manager_emp_nm').value = '';
                document.getElementById('txt_manager_emp_id_b').value = '';
                document.getElementById('txt_manager_emp_nm_b').value = '';
                document.getElementById('txt_vary_specialty').value = '';
                $('.middle-box').css({'background-color':'#f4f8fb'});
                $('#lbl_result').text('');
                document.getElementById('txt_asset_cd').value = '';
                return;
            }

            if (json.result == "1") {
                $('.middle-box').css({'background-color':'#4f81bd'});
                $('#lbl_result').text('일치');
            } else {
                $('.middle-box').css({'background-color':'#ff0000'});
                $('#lbl_result').text('불일치');
            }

            document.getElementById('txt_asset_nm').value = json.j_txt_asset_nm;
            document.getElementById('txt_asset_status_nm').value = json.j_txt_asset_status_nm;
            document.getElementById('txt_mgmt_com_cd').value = json.j_txt_mgmt_com_cd;
            document.getElementById('txt_mgmt_com_nm').value = json.j_txt_mgmt_com_nm;
            document.getElementById('txt_manage_dept_id').value = json.j_txt_manage_dept_id;
            document.getElementById('txt_manage_dept_nm').value = json.j_txt_manage_dept_nm;
            document.getElementById('txt_purchase_co').value = json.j_txt_purchase_co;
            document.getElementById('txt_purchase_co_nm').value = json.j_txt_purchase_co_nm;
            document.getElementById('txt_manager_emp_id').value = json.j_txt_manager_emp_id;
            document.getElementById('txt_manager_emp_nm').value = json.j_txt_manager_emp_nm;
            document.getElementById('txt_manager_emp_id_b').value= json.j_txt_manager_emp_id_b;
            document.getElementById('txt_manager_emp_nm_b').value = json.j_txt_manager_emp_nm_b;
            document.getElementById('txt_asset_nm').focus();

            $.ajax({
                url: "menu02_proc.asp",
                data:
                {
                    txt_hidden_flag: "asscd_cd_search_hidden",
                    txt_asset_cd: $("#txt_asset_cd").val(),
                    txt_datepicker: $("#datepicker1").val()
                },
                method: "POST",
                dataType: "json"
            })
            .done(function(json) {
                if (json.cnt != 0) {
                    if (!confirm("이미 실사 등록된 자산입니다. 그래도 계속 하시겠습니까?")) {
                        $("#txt_asset_cd").val("");
                        $("#txt_asset_nm").val("");
                        $("#txt_asset_status_nm").val("");
                        $("#txt_mgmt_com_cd").val("");
                        $("#txt_mgmt_com_nm").val("");
                        $("#txt_manage_dept_id").val("");
                        $("#txt_manage_dept_nm").val("");
                        $("#txt_purchase_co").val("");
                        $("#txt_purchase_co_nm").val("");
                        $("#txt_manager_emp_id").val("");
                        $("#txt_manager_emp_nm").val("");
                        $("#txt_manager_emp_id_b").val("");
                        $("#txt_manager_emp_nm_b").val("");
                        $("#txt_vary_specialty").val("");
                        $(".middle-box").css({"background-color":"#f4f8fb"});
                        $("#lbl_result").text("");
                        $("#txt_asset_cd").focus();
                        return;
                    }
                }
            }).fail(function(xhr, status, errorThrown) {}).always(function(xhr, status) {});
        }).fail(function(xhr, status, errorThrown) {}).always(function(xhr, status) {});
    }

    $("#txt_asset_cd").keyup(function(event) {
        if (event.keyCode == 13) {
            lfn_asset_search();
        }
    });

    $("#btn_insert").click(function() {
        var f = document.frm;

        if (f.datepicker1.value == "") {
            alert("실사일은 필수 입력 항목입니다.");
            f.datepicker1.focus();
            return;
        } else if (f.cbo_mgmt_com.value == "") {
            alert("관리회사는 필수 입력 항목입니다.");
            f.cbo_mgmt_com.focus();
            return;
        } else if (f.cbo_manage_dept.value == "") {
            alert("사용부서는 필수 입력 항목입니다.");
            f.cbo_manage_dept.focus();
            return;
        } else if (f.txt_asset_cd.value == "") {
            alert("자산번호는 필수 입력 항목입니다.");
            f.txt_asset_cd.focus();
            return;
        }

        f.txt_hidden_flag.value="btn_insert";
        f.action = "menu02_proc.asp"
        f.method = "post";
        f.target = "tmp";
        f.submit();

        $("#txt_asset_cd").focus();
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
        buttonText: "Select date",
        dateFormat: "yy/mm/dd"
    });
    $("#datepicker1").val(lfn_getday());
    $("#datepicker2").datepicker({
        showOn: "button",
        buttonImage: "../img/btn_calendar.png",
        buttonImageOnly: true,
        buttonText: "Select date",
        dateFormat: "yy/mm/dd"
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
        <h3>재고 실사</h3>
    </div>
</div>
<!-- //header -->

<form name="frm" id="frm" onkeydown="if (event.keyCode == 13) return false;">
<!-- container -->
<div class="container">

    <div class="sub-inquiry">
        <div class="title-area">
            <h4>실사 정보</h4>
        </div>
        <div class="inquiry-area">
            <table class="table table-condensed">
                <colgroup>
                    <col style="width:79px;">
                    <col>
                </colgroup>
                <tbody>
                <tr>
                    <th>실사일</th>
                    <td><input type="text" id="datepicker1" readonly="readonly" name="txt_inv_cnt_dt"></td>
                </tr>
                <tr>
                    <th>관리회사</th>
                    <td>
                        <select id="cbo_mgmt_com" name="cbo_mgmt_com">
                            <option selected>회로/모듈</option>
                            <option>옵션</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <th>사용부서</th>
                    <td>
                        <select id="cbo_manage_dept" name="cbo_manage_dept">
                            <option selected>생산 1</option>
                            <option>옵션</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <th>구매법인</th>
                    <td>
                        <select id="cbo_purchase_co" name="cbo_purchase_co">
                            <option selected>생산 1</option>
                            <option>옵션</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <th>관리자(정)</th>
                    <td>
                        <div class="input-group">
                            <span><input type="text" class="input-mini" readonly="readonly" id="txt_a_manager_emp_id" name="txt_a_manager_emp_id"></span>
                            <span><input type="text" class="alone" readonly="readonly" id="txt_a_manager_emp_nm" name="txt_a_manager_emp_nm"></span>
                            <span><button type="button" class="btn" id="myBtn2">검색</button></span>
                        </div>

                        <!-- 관리자(정) 모달팝업 -->
                        <div class="modal fade" id="myModal2" role="dialog">
                            <div class="modal-dialog">

                                <!-- Modal content-->
                                <div class="modal-content">

                                    <div class="modal-header">
                                        <h4 class="modal-title">사원 조회</h4>
                                    </div>

                                    <div class="modal-body">

                                        <!-- 조회 테이블 -->
                                        <div class="inquiry">
                                            <div class="inquiry-area">

                                                <table class="table table-condensed">
                                                    <tbody>
                                                    <tr>
                                                        <th>사번</th>
                                                        <td><input type="text" class="alone" id="txt_modal_emp_id_a" name="txt_modal_emp_id_a"></td>
                                                    </tr>
                                                    <tr>
                                                        <th>사원</th>
                                                        <td><input type="text" class="alone" id="txt_modal_emp_nm_a" name="txt_modal_emp_nm_a"></td>
                                                    </tr>
                                                    </tbody>
                                                </table>

                                            </div>
                                            <div class="mt-15 tcenter">
                                                <span><button class="btn btn-danger" type="button" id="btn_modal_emp_search_a" name="btn_modal_emp_search_a">조회</button></span>
                                            </div>
                                        </div>
                                        <!-- //조회 테이블 -->

                                        <!-- 데이터 테이블 -->
                                        <div class="sub-inquiry grid-table">
                                            <div class="inquiry-area">
                                                <table class="table table-condensed" id="modal_emp_search_a">
                                                    <colgroup>
                                                        <col style="width:33.3%;">
                                                        <col style="width:33.3%;">
                                                        <col style="width:33.3%;">
                                                    <thead>
                                                    <tr>
                                                        <th>사번</th>
                                                        <th>사원</th>
                                                        <th>부서명</th>
                                                    </tr>
                                                    </thead>
                                                    <tbody id="modal_emp_search_a_tbody">
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                        <!-- //데이터 테이블 -->
                                    </div>

                                    <div class="modal-footer">
                                        <button type="button" class="btn btn-default" data-dismiss="modal">닫기</button>
                                    </div>

                                </div>
                                <!-- //Modal content-->

                            </div>
                        </div>
                        <!-- //관리자(정) 모달팝업 -->

                        <script>
                            $(document).ready(function(){
                                $("#myBtn2").click(function(){
                                    document.getElementById("txt_modal_emp_id_a").value = "";
                                    document.getElementById("txt_modal_emp_nm_a").value = "";
                                    $("#modal_emp_search_a_tbody").empty();
                                    $("#myModal2").modal();
                                });
                            });
                        </script>

                    </td>
                </tr>
                <tr>
                    <th>관리자(부)</th>
                    <td>
                        <div class="input-group">
                            <span><input type="text" class="input-mini" readonly="readonly" id="txt_b_manager_emp_id" name="txt_b_manager_emp_id"></span>
                            <span><input type="text" class="alone" readonly="readonly" id="txt_b_manager_emp_nm" name="txt_b_manager_emp_nm"></span>
                            <span><button type="button" class="btn" id="myBtn3">검색</button></span>
                        </div>

                        <!-- 관리자(부) 모달팝업 -->
                        <div class="modal fade" id="myModal3" role="dialog">
                            <div class="modal-dialog">

                                <!-- Modal content-->
                                <div class="modal-content">

                                    <div class="modal-header">
                                        <h4 class="modal-title">사원 조회</h4>
                                    </div>

                                    <div class="modal-body">

                                        <!-- 조회 테이블 -->
                                        <div class="inquiry">
                                            <div class="inquiry-area">

                                                <table class="table table-condensed">
                                                    <tbody>
                                                    <tr>
                                                        <th>사번</th>
                                                        <td><input type="text" class="alone" id="txt_modal_emp_id_b" name="txt_modal_emp_id_b"></td>
                                                    </tr>
                                                    <tr>
                                                        <th>사원</th>
                                                        <td><input type="text" class="alone" id="txt_modal_emp_nm_b" name="txt_modal_emp_nm_b"></td>
                                                    </tr>
                                                    </tbody>
                                                </table>

                                            </div>
                                            <div class="mt-15 tcenter">
                                                <span><button class="btn btn-danger" type="button" id="btn_modal_emp_search_b" name="btn_modal_emp_search_b">조회</button></span>
                                            </div>
                                        </div>
                                        <!-- //조회 테이블 -->

                                        <!-- 데이터 테이블 -->
                                        <div class="sub-inquiry grid-table">
                                            <div class="inquiry-area">
                                                <table class="table table-condensed" id="modal_emp_search_b">
                                                    <colgroup>
                                                        <col style="width:33.3%;">
                                                        <col style="width:33.3%;">
                                                        <col style="width:33.3%;">
                                                    <thead>
                                                    <tr>
                                                        <th>사번</th>
                                                        <th>사원</th>
                                                        <th>부서명</th>
                                                    </tr>
                                                    </thead>
                                                    <tbody id="modal_emp_search_b_tbody">
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    <tr>
                                                        <td>KC2003-137</td>
                                                        <td>홍길동</td>
                                                        <td>품질관리팀</td>
                                                    </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                        <!-- //데이터 테이블 -->
                                    </div>

                                    <div class="modal-footer">
                                        <button type="button" class="btn btn-default" data-dismiss="modal">닫기</button>
                                    </div>

                                </div>
                                <!-- //Modal content-->

                            </div>
                        </div>
                        <!-- //관리자(부) 모달팝업 -->

                        <script>
                            $(document).ready(function(){
                                $("#myBtn3").click(function(){
                                    document.getElementById("txt_modal_emp_id_b").value = "";
                                    document.getElementById("txt_modal_emp_nm_b").value = "";
                                    $("#modal_emp_search_b_tbody").empty();
                                    $("#myModal3").modal();
                                });
                            });
                        </script>

                    </td>
                </tr>
                </tbody>
            </table>
        </div>
    </div>

    <div align="center"><div class="middle-box" style="width:60%; height:55px; font-size:30pt; text-align:center; line-height:55px; background-color:#f4f8fb"><span id="lbl_result"></span></div></div>

    <div class="sub-inquiry">
        <div class="title-area">
            <h4>현재 자산 정보</h4>
        </div>
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
                    <th>자산명</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_asset_nm" name="txt_asset_nm"></td>
                </tr>
                <tr>
                    <th>상태</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_asset_status_nm" name="txt_asset_status_nm"></td>
                </tr>
                <tr>
                    <th>관리회사</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_mgmt_com_nm" name="txt_mgmt_com_nm"><input type="hidden" id="txt_mgmt_com_cd" name="txt_mgmt_com_cd"></td>
                </tr>
                <tr>
                    <th>사용부서</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_manage_dept_nm" name="txt_manage_dept_nm"><input type="hidden" id="txt_manage_dept_id" name="txt_manage_dept_id"></td>
                </tr>
                <tr>
                    <th>구매법인</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_purchase_co_nm" name="txt_purchase_co_nm"><input type="hidden" id="txt_purchase_co" name="txt_purchase_co"></td>
                </tr>
                <tr>
                    <th>관리자(정)</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_manager_emp_nm" name="txt_manager_emp_nm"><input type="hidden" id="txt_manager_emp_id" name="txt_manager_emp_id"></td>
                </tr>
                <tr>
                    <th>관리자(부)</th>
                    <td><input type="text" value="" readonly="readonly" class="alone" id="txt_manager_emp_nm_b" name="txt_manager_emp_nm_b"><input type="hidden" id="txt_manager_emp_id_b" name="txt_manager_emp_id_b"></td>
                </tr>
                <tr>
                    <th>비고</th>
                    <td><input type="text" class="alone" id="txt_vary_specialty" name="txt_vary_specialty" tabindex="-1"></td>
                </tr>
                </tbody>
            </table>
        </div>
    </div>

    <div class="btn-bottom two-button">
        <div class="button-left">
            <button class="btn btn-danger" type="button" id="btn_insert" name="btn_insert">저장</button>
        </div>
        <div class="button-right">
            <button class="btn btn-info" type="button" onclick="javascript:location.replace('../main.asp')">닫기</button>
        </div>
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
    strSQL = strSQL & "    AND cd_class = N'ASPUCO' "
    strSQL = strSQL & "    AND cd_data <> N'000000' "
    strSQL = strSQL & "    AND use_flag = N'1' "

    objRS.open strSQL, objDBConn
    If Not (objRS.eof Or objRS.bof) Then
        Response.write "<script>"
        Response.write "var v_cbo_purchase_co = document.frm.cbo_purchase_co;"
        Response.write "v_cbo_purchase_co.options.length=0;"
        Response.write "v_cbo_purchase_co.options[0] = new Option('', '');"
        cnt = 1
        While Not objRS.eof
            Response.write "v_cbo_purchase_co.options[" & cnt & "] = new Option('" & objRS("cdname") & "', '" & objRS("cdcode") & "');"
            objRS.movenext
            cnt = cnt + 1
        Wend
        Response.write "</script>"
    End If
    objRS.Close: objDBConn.Close
%>
