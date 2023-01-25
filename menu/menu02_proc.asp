<% @Language="VBScript" CODEPAGE="65001" %>
<%
Response.CharSet="UTF-8"
Session.CodePage="65001"
Response.CodePage="65001"
%>
<!-- #include file="../common/auth_chk.asp" -->
<!-- #include file="../common/db_conn.asp" -->
<%
    Dim hidden_flag, strSQL, v_err, cnt
    hidden_flag = request("txt_hidden_flag")

    If (hidden_flag = "mgmt_com_change") then
        strSQL = ""
        strSQL = strSQL & " SELECT distinct A.dept_id AS cdcode "
        strSQL = strSQL & "      , A.dept_loc_nm      AS cdname "
        strSQL = strSQL & "   FROM hor_department AS A "
        strSQL = strSQL & "  INNER "
        strSQL = strSQL & "   JOIN oas_com_dept_link AS cdl "
        strSQL = strSQL & "     ON A.com_cd  = cdl.com_cd "
        strSQL = strSQL & "    AND A.div_cd  = cdl.div_cd "
        strSQL = strSQL & "    AND A.dept_id = cdl.dept_id "
        strSQL = strSQL & "    AND cdl.mgmt_com_cd   = N'" & request("cbo_mgmt_com") & "'"
        strSQL = strSQL & "  WHERE A.com_cd          = N'" & session("com_cd") & "'"
        strSQL = strSQL & "    AND A.use_flag        = N'Y' "
        strSQL = strSQL & "    AND CONVERT(VARCHAR(8), GETDATE(), 112) BETWEEN A.appnt_dt and A.abolish_dt "
        strSQL = strSQL & "  ORDER "
        strSQL = strSQL & "     BY dept_loc_nm "

        Response.write "<script>"
        Response.write "    var v_cbo_manage_dept = parent.document.frm.cbo_manage_dept;"
        Response.write "    v_cbo_manage_dept.options.length=0;"
        Response.write "    v_cbo_manage_dept.options[0] = new Option('선택하세요', '');"

        objRS.open strSQL, objDBConn
        If Not (objRS.eof Or objRS.bof) Then
            cnt = 1
            While Not objRS.eof
                Response.write "    v_cbo_manage_dept.options[" & cnt & "] = new Option('" & objRS("cdname") & "', '" & objRS("cdcode") & "');"
                objRS.movenext
                cnt = cnt + 1
            Wend
        End If
        Response.write "</script>"
        objRS.Close: objDBConn.Close


    ElseIf (hidden_flag = "modal_emp_search_a") Then
        strSQL = ""
        strSQL = strSQL & " SELECT A.emp_id "
        strSQL = strSQL & "      , A.emp_cd      emp_no "
        strSQL = strSQL & "      , A.emp_loc_nm  emp_nm "
        strSQL = strSQL & "      , A.dept_id "
        strSQL = strSQL & "      , B.dept_loc_nm "
        strSQL = strSQL & "   FROM hhm_emp_master A "
        strSQL = strSQL & "  INNER JOIN hor_department B ON A.com_cd   = B.com_cd "
        strSQL = strSQL & "                             AND A.div_cd   = B.div_cd "
        strSQL = strSQL & "                             AND A.dept_id  = B.dept_id "
        strSQL = strSQL & "                             AND B.use_flag = N'Y' "
        strSQL = strSQL & "  WHERE A.com_cd = N'" & session("com_cd") & "'"
        strSQL = strSQL & "    AND A.div_cd = N'" & session("div_cd") & "'"
        strSQL = strSQL & "    AND A.emp_cd LIKE N'%" & request("txt_modal_emp_id_a") & "%'"
        strSQL = strSQL & "    AND A.emp_loc_nm LIKE N'%" & request("txt_modal_emp_nm_a") & "%'"

        Response.write "<script> "
        Response.write "    var trCnt = parent.$('#modal_emp_search_a tr').length; "
        Response.write "    var innerHtml = ''; "
        Response.write "    parent.$('#modal_emp_search_a_tbody').empty(); "

        objRS.open strSQL, objDBConn
        If Not (objRS.eof Or objRS.bof) Then
            While Not objRS.eof
                Response.write "    innerHtml = ''; "
                Response.write "    innerHtml += '<tr onclick=lfn_modal_emp_search_a_rtn(""" & objRS("emp_no") & """,""" & objRS("emp_nm") & """);>'; "
                Response.write "    innerHtml += '    <td>" & objRS("emp_no") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("emp_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("dept_loc_nm") & "</td>'; "
                Response.write "    innerHtml += '</tr>'; "
                Response.write "    parent.$('#modal_emp_search_a > #modal_emp_search_a_tbody:last').append(innerHtml); "
                objRS.movenext
            Wend
        End If
        Response.write "</script>"
       objRS.Close: objDBConn.Close


    ElseIf (hidden_flag = "modal_emp_search_b") Then
        strSQL = ""
        strSQL = strSQL & " SELECT A.emp_id "
        strSQL = strSQL & "      , A.emp_cd      emp_no "
        strSQL = strSQL & "      , A.emp_loc_nm  emp_nm "
        strSQL = strSQL & "      , A.dept_id "
        strSQL = strSQL & "      , B.dept_loc_nm "
        strSQL = strSQL & "   FROM hhm_emp_master A "
        strSQL = strSQL & "  INNER JOIN hor_department B ON A.com_cd   = B.com_cd "
        strSQL = strSQL & "                             AND A.div_cd   = B.div_cd "
        strSQL = strSQL & "                             AND A.dept_id  = B.dept_id "
        strSQL = strSQL & "                             AND B.use_flag = N'Y' "
        strSQL = strSQL & "  WHERE A.com_cd = N'" & session("com_cd") & "'"
        strSQL = strSQL & "    AND A.div_cd = N'" & session("div_cd") & "'"
        strSQL = strSQL & "    AND A.emp_cd LIKE N'%" & request("txt_modal_emp_id_b") & "%'"
        strSQL = strSQL & "    AND A.emp_loc_nm LIKE N'%" & request("txt_modal_emp_nm_b") & "%'"

        Response.write "<script> "
        Response.write "    var trCnt = parent.$('#modal_emp_search_b tr').length; "
        Response.write "    var innerHtml = ''; "
        Response.write "    parent.$('#modal_emp_search_b_tbody').empty(); "

        objRS.open strSQL, objDBConn
        If Not (objRS.eof Or objRS.bof) Then
            While Not objRS.eof
                Response.write "    innerHtml = ''; "
                Response.write "    innerHtml += '<tr onclick=lfn_modal_emp_search_b_rtn(""" & objRS("emp_no") & """,""" & objRS("emp_nm") & """);>'; "
                Response.write "    innerHtml += '    <td>" & objRS("emp_no") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("emp_nm") & "</td>'; "
                Response.write "    innerHtml += '    <td>" & objRS("dept_loc_nm") & "</td>'; "
                Response.write "    innerHtml += '</tr>'; "
                Response.write "    parent.$('#modal_emp_search_b > #modal_emp_search_b_tbody:last').append(innerHtml); "
                objRS.movenext
            Wend
        End If
        Response.write "</script>"
       objRS.Close: objDBConn.Close


    ElseIf (hidden_flag = "asscd_cd_search_hidden") Then
        Dim v_cnt

        strSQL = ""
        strSQL = strSQL & " SELECT COUNT(*) AS cnt "
        strSQL = strSQL & "   FROM oas_inv_cnt "
        strSQL = strSQL & "  WHERE com_cd = N'" & session("com_cd") & "'"
        strSQL = strSQL & "    AND div_cd = N'" & session("div_cd") & "'"
        strSQL = strSQL & "    AND asset_cd = N'" & request("txt_asset_cd") & "'"
        strSQL = strSQL & "    AND inv_cnt_dt = '" & Replace(request("txt_datepicker"), "/", "") & "'"
        strSQL = strSQL & "    AND seq = 2 "
        strSQL = strSQL & "    AND class = '2' "

        objRS.open strSQL, objDBConn
        v_cnt = objRS("cnt")
        objRS.Close: objDBConn.Close

        Response.write "{"
        Response.write  """cnt" & "" & """:""" & v_cnt & """"
        Response.write "}"


    ElseIf (hidden_flag = "asscd_cd_search") Then
        Dim v_mgmt_com_a, v_manage_dept_a, v_a_manager_emp_id_a, v_b_manager_emp_id_a, v_purchase_co_a
        Dim v_mgmt_com_b, v_manage_dept_b, v_a_manager_emp_id_b, v_b_manager_emp_id_b, v_purchase_co_b

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
            Response.write  "{""cnt" & "" & """:""" & "0" & """}"
            Response.End
        End If

        Response.write "{"
        Response.write  """cnt" & "" & """:""" & "1" & """"

        v_mgmt_com_a = request("cbo_mgmt_com")
        v_mgmt_com_b = objRS("mgmt_com_cd")
        v_manage_dept_a = request("cbo_manage_dept")
        v_manage_dept_b = objRS("manage_dept_id")
        v_purchase_co_a = request("cbo_purchase_co")
        v_purchase_co_b = objRS("purchase_co")
        v_a_manager_emp_id_a = request("txt_a_manager_emp_id")
        v_a_manager_emp_id_b = objRS("manager_emp_id")
        v_b_manager_emp_id_a = request("txt_b_manager_emp_id")
        v_b_manager_emp_id_b = objRS("manager_emp_id_b")

        If (v_mgmt_com_a = v_mgmt_com_b And v_manage_dept_a = v_manage_dept_b And v_a_manager_emp_id_a = v_a_manager_emp_id_b And v_b_manager_emp_id_a = v_b_manager_emp_id_b And v_purchase_co_a = v_purchase_co_b) Then
            Response.write  ",""result" & "" & """:""" & "1" & """"
        Else
            Response.write  ",""result" & "" & """:""" & "0" & """"
        End If

        Response.write  ",""j_txt_asset_nm" & "" & """:""" & Replace(objRS("asset_nm"), "'", "\'") & """"
        Response.write  ",""j_txt_asset_status_nm" & "" & """:""" & objRS("asset_status_nm") & """"
        Response.write  ",""j_txt_mgmt_com_cd" & "" & """:""" & objRS("mgmt_com_cd") & """"
        Response.write  ",""j_txt_mgmt_com_nm" & "" & """:""" & objRS("mgmt_com_nm") & """"
        Response.write  ",""j_txt_manage_dept_id" & "" & """:""" & objRS("manage_dept_id") & """"
        Response.write  ",""j_txt_manage_dept_nm" & "" & """:""" & objRS("manage_dept_nm") & """"
        Response.write  ",""j_txt_purchase_co" & "" & """:""" & objRS("purchase_co") & """"
        Response.write  ",""j_txt_purchase_co_nm" & "" & """:""" & objRS("purchase_co_nm") & """"
        Response.write  ",""j_txt_manager_emp_id" & "" & """:""" & objRS("manager_emp_id") & """"
        Response.write  ",""j_txt_manager_emp_nm" & "" & """:""" & objRS("manager_emp_nm") & """"
        Response.write  ",""j_txt_manager_emp_id_b" & "" & """:""" & objRS("manager_emp_id_b") & """"
        Response.write  ",""j_txt_manager_emp_nm_b" & "" & """:""" & objRS("manager_emp_nm_b") & """"
        Response.write "}"
        objRS.Close: objDBConn.Close


    ElseIf (hidden_flag = "btn_insert") Then
        On Error Resume Next
        objDBConn.BeginTrans

        strSQL = ""

        strSQL = strSQL & " DELETE oas_inv_cnt "
        strSQL = strSQL & "  WHERE com_cd = N'" & session("com_cd") & "'"
        strSQL = strSQL & "    AND div_cd = N'" & session("div_cd") & "'"
        strSQL = strSQL & "    AND asset_cd = N'" & request("txt_asset_cd") & "'"
        strSQL = strSQL & "    AND inv_cnt_dt = '" & Replace(request("txt_inv_cnt_dt"), "/", "") & "'"
        strSQL = strSQL & "    AND seq = 2 "

        strSQL = strSQL & " INSERT "
        strSQL = strSQL & "   INTO oas_inv_cnt "
        strSQL = strSQL & "      ( com_cd "
        strSQL = strSQL & "      , div_cd "
        strSQL = strSQL & "      , asset_cd "
        strSQL = strSQL & "      , inv_cnt_dt "
        strSQL = strSQL & "      , seq "
        strSQL = strSQL & "      , mgmt_com_cd "
        strSQL = strSQL & "      , manage_dept_id "
        strSQL = strSQL & "      , manager_emp_id "
        strSQL = strSQL & "      , manager_emp_id_b "
        strSQL = strSQL & "      , specialty "
        strSQL = strSQL & "      , class "
        strSQL = strSQL & "      , create_dt "
        strSQL = strSQL & "      , create_by "
        strSQL = strSQL & "      , update_dt "
        strSQL = strSQL & "      , update_by "
        strSQL = strSQL & "      , purchase_co "
        strSQL = strSQL & "      ) "
        strSQL = strSQL & " VALUES "
        strSQL = strSQL & "      ( N'" & session("com_cd") & "'"
        strSQL = strSQL & "      , N'" & session("div_cd") & "'"
        strSQL = strSQL & "      , N'" & request("txt_asset_cd") & "'"
        strSQL = strSQL & "      , '" & Replace(request("txt_inv_cnt_dt"), "/", "") & "'"
        strSQL = strSQL & "      , 2 "
        strSQL = strSQL & "      , N'" & request("cbo_mgmt_com") & "'"
        strSQL = strSQL & "      , N'" & request("cbo_manage_dept") & "'"
        strSQL = strSQL & "      , N'" & request("txt_a_manager_emp_id") & "'"
        strSQL = strSQL & "      , N'" & request("txt_b_manager_emp_id") & "'"
        strSQL = strSQL & "      , N'" & request("txt_vary_specialty") & "'"
        strSQL = strSQL & "      , '2' "
        strSQL = strSQL & "      , GETDATE() "
        strSQL = strSQL & "      , N'" + session("uid") & "'"
        strSQL = strSQL & "      , GETDATE() "
        strSQL = strSQL & "      , N'" + session("emp_id") & "'"
        strSQL = strSQL & "      , N'" & request("cbo_purchase_co") & "'"
        strSQL = strSQL & "      ) "

        objDBConn.execute strSQL

        v_err = objDBConn.errors.count
        If (v_err > 0 Or Err.number > 0) Then
            objDBConn.RollbackTrans
            objDBConn.Close
            Response.write "<script>alert(""" & Err.Description & """);</script>"
            Response.End
        Else
            objDBConn.CommitTrans
            objDBConn.Close
        End if

        Response.write "<script>"
        Response.write "    parent.document.getElementById('txt_asset_cd').value = '';"
        Response.write "    parent.document.getElementById('txt_asset_nm').value = '';"
        Response.write "    parent.document.getElementById('txt_asset_status_nm').value = '';"
        Response.write "    parent.document.getElementById('txt_mgmt_com_cd').value = '';"
        Response.write "    parent.document.getElementById('txt_mgmt_com_nm').value = '';"
        Response.write "    parent.document.getElementById('txt_manage_dept_id').value = '';"
        Response.write "    parent.document.getElementById('txt_manage_dept_nm').value = '';"
        Response.write "    parent.document.getElementById('txt_purchase_co').value = '';"
        Response.write "    parent.document.getElementById('txt_purchase_co_nm').value = '';"
        Response.write "    parent.document.getElementById('txt_manager_emp_id').value = '';"
        Response.write "    parent.document.getElementById('txt_manager_emp_nm').value = '';"
        Response.write "    parent.document.getElementById('txt_manager_emp_id_b').value = '';"
        Response.write "    parent.document.getElementById('txt_manager_emp_nm_b').value = '';"
        Response.write "    parent.document.getElementById('txt_vary_specialty').value = '';"
        Response.write "    parent.$('.middle-box').css({'background-color':'#f4f8fb'});"
        Response.write "    parent.$('#lbl_result').text('');"
        Response.write "    alert(""저장이 완료되었습니다."");"
        Response.write "</script>"

    End If
%>
