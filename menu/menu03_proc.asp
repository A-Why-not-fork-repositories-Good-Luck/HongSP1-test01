<% @Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="UTF-8"
Session.CodePage="65001"
Response.CodePage="65001"
%>
<!-- #include file="../common/auth_chk.asp" -->
<!-- #include file="../common/db_conn.asp" -->
<%
    Dim hidden_flag, strSQL
    hidden_flag = request("txt_hidden_flag")

    If (hidden_flag = "asscd_cd_search") Then
        strSQL = ""
        strSQL = strSQL & " SELECT com_cd "
        strSQL = strSQL & "      , div_cd "
        strSQL = strSQL & "      , asset_cd "
        strSQL = strSQL & "      , asset_nm "
        strSQL = strSQL & "      , spec "
        strSQL = strSQL & "      , asset_class "
        strSQL = strSQL & "      , asset_class_nm "
        strSQL = strSQL & "      , mgmt_class "
        strSQL = strSQL & "      , mgmt_class_nm "
        strSQL = strSQL & "      , s_no "
        strSQL = strSQL & "      , CASE WHEN (make_dt IS NULL OR make_dt = '') THEN '' ELSE CONVERT(VARCHAR, CONVERT(DATETIME, make_dt), 111) END AS make_dt "
        strSQL = strSQL & "      , make_vend_nm "
        strSQL = strSQL & "      , puchase_vend_cd "
        strSQL = strSQL & "      , puchase_vend_nm "
        strSQL = strSQL & "      , person_in_charge "
        strSQL = strSQL & "      , purchase_vend_tel "
        strSQL = strSQL & "      , CASE WHEN (buy_dt IS NULL OR buy_dt = '') THEN '' ELSE CONVERT(VARCHAR, CONVERT(DATETIME, buy_dt), 111) END AS buy_dt "
        strSQL = strSQL & "      , FORMAT(buy_amount, '#,#') AS buy_amount "
        strSQL = strSQL & "      , specialty "
        strSQL = strSQL & "      , attach_file "
        strSQL = strSQL & "      , vary_dt "
        strSQL = strSQL & "      , vary_div "
        strSQL = strSQL & "      , vary_div_nm "
        strSQL = strSQL & "      , asset_status "
        strSQL = strSQL & "      , asset_status_nm "
        strSQL = strSQL & "      , mgmt_com_cd "
        strSQL = strSQL & "      , mgmt_com_nm "
        strSQL = strSQL & "      , manage_dept_id "
        strSQL = strSQL & "      , manage_dept_nm "
        strSQL = strSQL & "      , purchase_co "
        strSQL = strSQL & "      , purchase_co_nm "
        strSQL = strSQL & "      , manager_emp_id "
        strSQL = strSQL & "      , manager_emp_nm "
        strSQL = strSQL & "      , manager_emp_id_b "
        strSQL = strSQL & "      , manager_emp_nm_b "
        strSQL = strSQL & "      , vary_specialty "
        strSQL = strSQL & "      , create_class "
        strSQL = strSQL & "      , create_class_nm "
        strSQL = strSQL & "   FROM v_oas_asset "
        strSQL = strSQL & "  WHERE com_cd = N'" & session("com_cd") + "'"
        strSQL = strSQL & "    AND div_cd = N'" & session("div_cd") + "'"
        strSQL = strSQL & "    AND asset_cd = N'" & request("txt_asset_cd") & "'"

        objRS.open strSQL, objDBConn
        If (objRS.eof Or objRS.bof) Then
            objRS.Close: objDBConn.Close
            Response.write "<script>"
            Response.write "    alert(""자산번호를 확인하시기 바랍니다."");"
            Response.write "    parent.document.getElementById('txt_asset_cd').value = '';"
            Response.write "    parent.document.getElementById('txt_asset_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_spec').value = '';"
            Response.write "    parent.document.getElementById('txt_asset_class_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_mgmt_class_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_make_vend_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_make_dt').value = '';"
            Response.write "    parent.document.getElementById('txt_puchase_vend_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_s_no').value = '';"
            Response.write "    parent.document.getElementById('txt_buy_dt').value = '';"
            Response.write "    parent.document.getElementById('txt_purchase_vend_tel').value = '';"
            Response.write "    parent.document.getElementById('txt_person_in_charge').value = '';"
            Response.write "    parent.document.getElementById('txt_buy_amount').value = '';"
            Response.write "    parent.document.getElementById('txt_specialty').value = '';"
            Response.write "    parent.document.getElementById('txt_asset_status_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_mgmt_com_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_manage_dept_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_purchase_co_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_manager_emp_nm').value = '';"
            Response.write "    parent.document.getElementById('txt_manager_emp_nm_b').value = '';"
            Response.write "    parent.$('#grd01_body').empty();"
            Response.write "</script>"
            Response.End
        End If
        Response.write "<script>"
        Response.write "    parent.document.getElementById('txt_asset_nm').value = '" & Replace(objRS("asset_nm"), "'", "\'") & "';"
        Response.write "    parent.document.getElementById('txt_spec').value = '" & objRS("spec") & "';"
        Response.write "    parent.document.getElementById('txt_asset_class_nm').value = '" & objRS("asset_class_nm") & "';"
        Response.write "    parent.document.getElementById('txt_mgmt_class_nm').value = '" & objRS("mgmt_class_nm") & "';"
        Response.write "    parent.document.getElementById('txt_make_vend_nm').value = '" & objRS("make_vend_nm") & "';"
        Response.write "    parent.document.getElementById('txt_make_dt').value = '" & objRS("make_dt") & "';"
        Response.write "    parent.document.getElementById('txt_puchase_vend_nm').value = '" & objRS("puchase_vend_nm") & "';"
        Response.write "    parent.document.getElementById('txt_s_no').value = '" & objRS("s_no") & "';"
        Response.write "    parent.document.getElementById('txt_buy_dt').value = '" & objRS("buy_dt") & "';"
        Response.write "    parent.document.getElementById('txt_purchase_vend_tel').value = '" & objRS("purchase_vend_tel") & "';"
        Response.write "    parent.document.getElementById('txt_person_in_charge').value = '" & objRS("person_in_charge") & "';"
        Response.write "    parent.document.getElementById('txt_buy_amount').value = '" & objRS("buy_amount") & "';"
        Response.write "    parent.document.getElementById('txt_specialty').value = '" & objRS("specialty") & "';"
        Response.write "    parent.document.getElementById('txt_asset_status_nm').value = '" & objRS("asset_status_nm") & "';"
        Response.write "    parent.document.getElementById('txt_mgmt_com_nm').value = '" & objRS("mgmt_com_nm") & "';"
        Response.write "    parent.document.getElementById('txt_manage_dept_nm').value = '" & objRS("manage_dept_nm") & "';"
        Response.write "    parent.document.getElementById('txt_purchase_co_nm').value = '" & objRS("purchase_co_nm") & "';"
        Response.write "    parent.document.getElementById('txt_manager_emp_nm').value = '" & objRS("manager_emp_nm") & "';"
        Response.write "    parent.document.getElementById('txt_manager_emp_nm_b').value = '" & objRS("manager_emp_nm_b") & "';"
        Response.write "    parent.document.getElementById('txt_asset_nm').focus();"
        Response.write "</script>"
        objRS.Close

        strSQL = ""
        strSQL = strSQL & " SELECT CONVERT(VARCHAR, CONVERT(DATETIME, asv.vary_dt), 111) AS vary_dt                            /* 변동 일 */ "
        strSQL = strSQL & "      , asv.seq "
        strSQL = strSQL & "      , asv.vary_div                           /* 최종 변동 구분 코드 */ "
        strSQL = strSQL & "      , mvr.cd_nm_loc    AS vary_div_nm        /* 최종 변동 구분 명 */ "
        strSQL = strSQL & "      , asv.mgmt_com_cd                        /* 최종 관리 사 코드 */ "
        strSQL = strSQL & "      , mcm.cd_nm_loc    AS mgmt_com_nm        /* 최종 관리 회사 명 */ "
        strSQL = strSQL & "      , asv.manage_dept_id                     /* 최종 관리부서 코드 */ "
        strSQL = strSQL & "      , hdm.dept_loc_nm  AS manage_dept_nm     /* 최종 관리부서 명 */ "
        strSQL = strSQL & "      , asv.purchase_co                        /* 최종 구매법인코드 */ "
        strSQL = strSQL & "      , puc.cd_nm_loc      AS purchase_co_nm   /* 최종 구매법인 명 */ "
        strSQL = strSQL & "      , asv.manager_emp_id                     /* 최종 관리자(정) 코드 */ "
        strSQL = strSQL & "      , hem.emp_loc_nm   AS manager_emp_nm     /* 최종 관리자(정) 명 */ "
        strSQL = strSQL & "      , asv.manager_emp_id_b                   /* 최종 관리자(부) 코드 */ "
        strSQL = strSQL & "      , hem_b.emp_loc_nm AS manager_emp_nm_b   /* 최종 관리자(부) 명 */ "
        strSQL = strSQL & "      , asv.specialty    AS vary_specialty     /* 최종 자산변동 비고 */ "
        strSQL = strSQL & "      , asv.create_class                       /* 자산변동 생성구분 */ "
        strSQL = strSQL & "      , acc.cd_nm_loc    AS create_class_nm    /* 자산변동 생성구분 명 */ "
        strSQL = strSQL & "   FROM oas_asset_vary asv "
        strSQL = strSQL & "   LEFT "
        strSQL = strSQL & "   JOIN mco_code mvr "
        strSQL = strSQL & "     ON asv.com_cd       = mvr.com_cd "
        strSQL = strSQL & "    AND asv.vary_div     = mvr.cd_data "
        strSQL = strSQL & "    AND mvr.cd_class     = 'ASDVCL' "
        strSQL = strSQL & "    AND mvr.use_flag     = '1' "
        strSQL = strSQL & "   LEFT "
        strSQL = strSQL & "   JOIN mco_code mcm "
        strSQL = strSQL & "     ON asv.com_cd       = mcm.com_cd "
        strSQL = strSQL & "    AND asv.mgmt_com_cd  = mcm.cd_data "
        strSQL = strSQL & "    AND mcm.cd_class     = 'ASMGCM' "
        strSQL = strSQL & "    AND mcm.use_flag     = '1' "
        strSQL = strSQL & "   LEFT "
        strSQL = strSQL & "   JOIN mco_code acc "
        strSQL = strSQL & "     ON asv.com_cd       = acc.com_cd "
        strSQL = strSQL & "    AND asv.create_class = acc.cd_data "
        strSQL = strSQL & "    AND acc.cd_class     = 'MMIOCR' "
        strSQL = strSQL & "    AND acc.use_flag     = '1' "
        strSQL = strSQL & "  LEFT "
        strSQL = strSQL & "  JOIN mco_code puc "
        strSQL = strSQL & "    ON asv.com_cd       = puc.com_cd "
        strSQL = strSQL & "   AND asv.purchase_co  = puc.cd_data "
        strSQL = strSQL & "   AND puc.cd_class     = 'ASPUCO'  /* 구매법인 */ "
        strSQL = strSQL & "   AND puc.use_flag     = '1' "
        strSQL = strSQL & "   LEFT "
        strSQL = strSQL & "   JOIN hor_department   hdm "
        strSQL = strSQL & "     ON asv.com_cd           = hdm.com_cd "
        strSQL = strSQL & "    AND asv.div_cd           = hdm.div_cd "
        strSQL = strSQL & "    AND asv.manage_dept_id   = hdm.dept_id "
        strSQL = strSQL & "    AND CONVERT(VARCHAR(8), getdate(), 112) between hdm.appnt_dt and hdm.abolish_dt "
        strSQL = strSQL & "   LEFT "
        strSQL = strSQL & "   JOIN hhm_emp_master   hem "
        strSQL = strSQL & "     ON asv.com_cd           = hem.com_cd "
        strSQL = strSQL & "    AND asv.div_cd           = hem.div_cd "
        strSQL = strSQL & "    AND asv.manager_emp_id   = hem.emp_id "
        strSQL = strSQL & "   LEFT "
        strSQL = strSQL & "   JOIN hhm_emp_master   hem_b "
        strSQL = strSQL & "     ON asv.com_cd           = hem_b.com_cd "
        strSQL = strSQL & "    AND asv.div_cd           = hem_b.div_cd "
        strSQL = strSQL & "    AND asv.manager_emp_id_b = hem_b.emp_id "
        strSQL = strSQL & "  WHERE asv.com_cd       = N'" & session("com_cd") & "'"
        strSQL = strSQL & "    AND asv.div_cd       = N'" & session("div_cd") & "'"
        strSQL = strSQL & "    AND asv.asset_cd     = N'" & request("txt_asset_cd") & "'"
        strSQL = strSQL & "  ORDER "
        strSQL = strSQL & "     BY asv.vary_dt "
        strSQL = strSQL & "      , asv.seq "

        Response.write "<script> "
        Response.write "    var trCnt = parent.$('#grd01 tr').length; "
        Response.write "    var innerHtml = ''; "
        Response.write "    parent.$('#grd01_body').empty(); "

        objRS.open strSQL, objDBConn
        If Not (objRS.eof Or objRS.bof) Then
            While Not objRS.eof
                Response.write "    innerHtml = ''; "
                Response.write "    innerHtml += '<tr>'; "
                Response.write "    innerHtml += '    <td class=tcenter>" & objRS("vary_dt") & "</td>'; "
                Response.write "    innerHtml += '    <td class=tcenter>" & objRS("vary_div_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("mgmt_com_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("manage_dept_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("purchase_co_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("manager_emp_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("manager_emp_nm_b") & "</td>'; "
                Response.write "    innerHtml += '    <td class=tcenter>" & objRS("create_class_Nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("vary_specialty") & "</td>'; "
                Response.write "    innerHtml += '</tr>'; "
                Response.write "    parent.$('#grd01 > #grd01_body:last').append(innerHtml); "
                objRS.movenext
            Wend
        End If
        Response.write "</script>"
       objRS.Close: objDBConn.Close

    End If
%>