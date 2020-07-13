Public Class frmConnectDB
    Dim strConn1 As String = ""
    Dim strConn2 As String = ""
    Dim strConn3 As String = ""
    Dim strConn4 As String = ""
    Dim strConn5 As String = ""

    Private Sub frmConnectDB_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Dim dTB As New DataTable()
        Dim dRow As DataRow
        'Dim i As Integer

        strConn1 = DBConnString.strConn1
        strConn2 = DBConnString.strConn2
        strConn3 = DBConnString.strConn3
        strConn4 = DBConnString.strConn4

        strConn5 = DBConnString.strConn5

        dTB.Columns.Add(New DataColumn("DB_Name", GetType(String)))
        dTB.Columns.Add(New DataColumn("DB_Connect", GetType(String)))

        dRow = dTB.NewRow
        dRow(0) = "บริษัท แพนเอเซีย-IP"
        dRow(1) = strConn1
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "บ.แพน เอเซีย-EDP"
        dRow(1) = strConn2
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "ผ่าน Internet"
        dRow(1) = strConn3
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "ข้อมูลทดสอบ"
        dRow(1) = strConn4
        dTB.Rows.Add(dRow)

        dRow = dTB.NewRow
        dRow(0) = "ทดสอบ Inter"
        dRow(1) = strConn5
        dTB.Rows.Add(dRow)



        With cboDBlist
            .DataSource = dTB
            .DisplayMember = "DB_Name"
            .ValueMember = "DB_Connect"
        End With

        cboDBlist.SelectedIndex = 0
        strConn = cboDBlist.SelectedValue ' DBConnString.strConn2
        lbDBconnect.Text = strConn
        lbTimer1.Text = 0
        Timer1.Enabled = True
        'With cboDBlist


        '    .Items.Add(strConn2)
        '    .Items.Add(strConn1)
        '    .SelectedItem = 1

        'End With
    End Sub

    Private Sub cboDBlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboDBlist.SelectedIndexChanged
        Try
            strConn = cboDBlist.SelectedValue ' DBConnString.strConn2
            lbDBconnect.Text = strConn
            'lbTimer1.Text = 0
            'Timer1.Enabled = True
        Catch ex As Exception

        End Try

    End Sub

    Private Sub cmdConnect_Click(sender As Object, e As EventArgs) Handles cmdConnect.Click
        Call RunDB()
    End Sub
    Sub RunDB()
        ' Dim frmBegin As New frmBegin
        Dim frmBegin As New frmBegin

        DBtools.openDB()
        With Conn
            If .State = ConnectionState.Open Then
                Me.Hide()
                frmBegin.ShowDialog()

                ' frmBegin.Show()

            Else
                '  MsgBox("ไม่สามารถติดต่อฐานข้อมูลได้")
            End If

        End With
        'End

    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        lbTimer1.Text = lbTimer1.Text + 1
        lbTime.Text = Now
        If lbTimer1.Text = 1 Then
            'Call selectOK()
            Call RunDB()
            Me.Close()
        End If
    End Sub
    'Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


    'End Sub
End Class