<%If session("uid") = "" Then%>
    <script>
        alert("TRF 자산관리 시스템을 이용하기 위해서는 로그인 하셔야 합니다.");
        location.replace("http://211.48.148.52:8023/xone_mgmt_web");
    </script>
<%
    Response.End
End If
%>
