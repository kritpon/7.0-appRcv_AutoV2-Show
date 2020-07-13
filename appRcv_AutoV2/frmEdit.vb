Public Class frmEdit
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        'Dim strAns As String = ""

        txtSQL = "Select Trh_Type,Trh_No,Trh_DateSection,Trh_KeyType,Dtl_n_trade,"
        txtSQL = txtSQL & "sum(Dtl_Num)as sumQty,sum(Dtl_Num_2)as sumW "
        txtSQL = txtSQL & "From TranDataH left Join TranDataD "
        txtSQL = txtSQL & "On (TranDataH.Trh_Type=TranDataD.Dtl_Type And TranDataH.Trh_NO=TranDataD.Dtl_No ) "
        txtSQL = txtSQL & "Where Trh_type='M' "
        txtSQL = txtSQL & "And Trh_KeyType ='B' "
        txtSQL = txtSQL & "And Trh_DateSection='20-03-2019' "
        txtSQL = txtSQL & "group by Trh_type ,Trh_No,Trh_DateSection,Trh_KeyType,Dtl_n_trade "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dataList")

        For i = 0 To subDS.Tables("dataList").Rows.Count - 1
            Dim subds1 As New DataSet
            Dim subda1 As SqlClient.SqlDataAdapter

            Dim strDocNO As String = subDS.Tables("dataList").Rows(i).Item("Trh_No")

            txtSQL = "Select BOM_No "
            txtSQL = txtSQL & "From BOMmastF "
            txtSQL = txtSQL & "Where  year(BOM_RM_Updatetime)='2019' "
            txtSQL = txtSQL & "And day(BOM_RM_Updatetime)='20' "
            txtSQL = txtSQL & "And month(BOM_RM_Updatetime)='03' "
            txtSQL = txtSQL & "And BOM_RM_Scales='5' "
            txtSQL = txtSQL & "And BOM_No='" & strDocNO & "' "

            subda1 = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subda1.Fill(subds1, "dataList1")

            If subds1.Tables("dataList1").Rows.Count > 0 Then
            Else
                txtSQL = "Delete TranDataH "
                txtSQL = txtSQL & "Where Trh_Type='M' "
                txtSQL = txtSQL & "And Trh_No='" & strDocNO & "' "
                DBtools.dbSaveSQLsrv(txtSQL, "")

                txtSQL = "Delete TranDataD "
                txtSQL = txtSQL & "Where Dtl_Type='M' "
                txtSQL = txtSQL & "And Dtl_No='" & strDocNO & "' "
                DBtools.dbSaveSQLsrv(txtSQL, "")

            End If

        Next

    End Sub
End Class