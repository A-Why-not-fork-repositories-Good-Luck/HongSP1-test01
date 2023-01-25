<% @Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="UTF-8"
Session.CodePage="65001"
Response.CodePage="65001"
%>
<!-- #include file="db_conn.asp" -->
<%
Dim uid, pwd, com_cd

uid = request("edt_user_id")
pwd = request("edt_user_pw")
com_cd = "11"

strSQL = ""
strSQL = strSQL & " SELECT zu.com_cd "
strSQL = strSQL & "      , hc.com_loc_nm "
strSQL = strSQL & "      , zu.emp_id "
strSQL = strSQL & "      , zu.user_id "
strSQL = strSQL & "      , dbo.gfn_compare_crypto(N'" & pwd & "', zu.user_password) AS pwd_chk "
strSQL = strSQL & "      , ISNULL(zu.system_admin_flag, '0') AS system_admin_flag "
strSQL = strSQL & "      , zu.notice_admin_flag "
strSQL = strSQL & "      , ISNULL(hem.mgmt_com_cd, '') AS mgmt_com_cd "
strSQL = strSQL & "   FROM zsy_user AS zu "
strSQL = strSQL & "  INNER "
strSQL = strSQL & "   JOIN hor_company AS hc "
strSQL = strSQL & "     ON zu.com_cd = hc.com_cd "
strSQL = strSQL & "   LEFT "
strSQL = strSQL & "   JOIN hhm_emp_master AS hem "
strSQL = strSQL & "     ON hem.com_cd = zu.com_cd "
strSQL = strSQL & "    AND hem.emp_id = zu.emp_id "
strSQL = strSQL & "  WHERE zu.com_cd    = N'" & com_cd & "'"
strSQL = strSQL & "    AND zu.user_id   = N'" & uid & "'"
strSQL = strSQL & "    AND zu.user_type = '0' "

objRS.open strSQL, objDBConn
If objRS.eof Or objRS.bof Then
%>
    <script>
        alert("시스템에 존재하지 않는 아이디 입니다.");
        parent.document.frm.edt_user_id.focus();
    </script>
<%
    objRS.Close: objDBConn.Close
    Response.End

ElseIf IsNull(objRS("pwd_chk")) Then
%>
    <script>
        alert("비밀번호를 확인하시기 바랍니다.");
        parent.document.frm.edt_user_pw.focus();
    </script>
<%
    objRS.Close: objDBConn.Close
    Response.End
End If

session("uid") = objRS("user_id")
session("emp_id") = objRS("emp_id")
session("com_cd") = objRS("com_cd")
session("div_cd") = objRS("com_cd")
session("mgmt_com_cd") = objRS("mgmt_com_cd")
session("system_admin_flag") = objRS("system_admin_flag")
session("notice_admin_flag") = objRS("notice_admin_flag")

objRS.Close: objDBConn.Close
%>
<script>
    parent.location.replace("http://211.48.148.52:8023/xone_mgmt_web/main.asp");
</script>
