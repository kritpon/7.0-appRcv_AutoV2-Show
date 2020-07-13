


Module DBtools
    'Public Declare Function ActivateKeyboardLayout Lib "user32" (ByVal HKL As Long, ByVal flags As Long) As Long
    Public chkUserUpdate As Boolean



    '=========   ตัวแปรเรื่องการ บันทึกข้อมุล
    Public chkSaveOK As Boolean = False

    Public strFactor As String = ""
    Public strCon As String
    'Public SLid As String

    Public pathDB As String
    Public pathDB02 As String


    Public Conn As SqlClient.SqlConnection = New SqlClient.SqlConnection()
    Public subCom As SqlClient.SqlCommand = New SqlClient.SqlCommand
    Public txtSQL As String
    Public txtSQL1 As String
    Public txtSQL2 As String

    Public docType As String = "S"

    Public Str01 As String
    Public gDA As SqlClient.SqlDataAdapter
    Public gDs As DataSet = New DataSet()
    Public selWH As String = ""  'เลือกคลังสินค้่า

    Public selCusName As String
    Public selCountry As String
    Public selAmphoe As String

    'Public selZip As String
    Public selCusiD As String

    Public selSale As String
    Public selSaleName As String

    Public strConn As String

    Public CusId As String
    Public CusName As String
    Public CusCreTerm As Integer
    Public CusSaleName As String
    Public CusSaleID As String
    Public CusDiscount As String 'Double
    Public CusLimit As Double
    Public ramDateCost As String = ""
    Public chkNew As Boolean = False
    Public chkEditDoc As Boolean = False
    Public EditNo As String
    Public EditCus As String

    Public PId As String = "" 'เก็บรหัสลูกค้า
    Public PName As String = "" 'เก็บชื่อลูกค้า
    Public Pdate As String = ""
    Public Pdate2 As String = ""
    Public SelectDate As Date
    Public SelectNo As String = "" 'เก็บเลขที่ Order

    Public SelectCode As String = "" 'เก็บรหัสสินค้า
    Public SelectName As String = "" 'เก็บชื่อสินค้า
    Public SelectNum As Integer = 0 'เก็บจำนวน
    Public SelectPrice As Double = 0 '
    Public SelectDisc As String = ""

    Public SelectWeight As Double = 0 'เก็บน้ำหนักแผ่น
    Public SelectStock As Double = 0 'เก็บStock
    Public SelectPNo As String
    Public Stock As Integer = 0  'stock

    Public SelectNo2 As String 'เก็บเลขที่ Order
    Public SelectName2 As String = "" 'เก็บชื่อสินค้า
    Public SelectNum2 As Integer = 0 'เก็บจำนวน
    Public SelectPrice2 As Double = 0 'เก็บราคา
    Public SelectCode2 As String = "" 'เก็บรหัสสินค้า
    Public SelectWeight2 As Double = 0  'เก็บน้ำหนักแผ่น
    Public SelectStock2 As Double = 0 'เก็บStock
    Public SelectPNo2 As String = ""
    Public SelectRevNo2 As String = ""  ' เก็บข้อมูลใบจอง
    Public SelectTypeP_No2 As String = ""  ' เก็บใบเบิก
    Public selTypeVAT As String = ""

    Public selectDetail As String = ""
    Public selectCondition As String = ""


    Public CodeT As String = ""
    Public CodeG As String = ""
    Public CodeColor As String = ""
    Public CodeTh As String = ""
    Public CodeSize As String = ""
    Public CodePaper As String = ""
    Public CodeGrade As String = ""

    Public Ddate As String = ""
    Public Dno As String = ""
    Public orderNum As String = ""
    Public Dvat As String = ""
    Public DPNo As String = ""
    Public Dcus As String = ""
    Public Dpro As String = ""
    Public Dnum As Integer = 0
    Public Dprice As Single = 0
    Public Dweight As Single = 0
    Public Dproduct As String = ""
    Public Dcusname As String = ""
    Public Ditem As String = ""
    Public DOrder As String = ""
    Public DDetail2 As String = ""
    Public DDisc As String = ""
    Public NoDoc As String = ""

    Public LEdit As ListViewItem
    Public selStkID As String   ' ตัวแปลร่วมสำหรับเก็บค่า รหัสสินค้า

    Public trh_RunID As Integer = 0

    Public subDs As DataSet = New DataSet
    Public subDa As SqlClient.SqlDataAdapter

    'Public CustomerId As String
    Public ThaiBaht01 As String
    Public Num As Integer
    Public chkID As Boolean
    Public NumberDoc As String
    'Public TypeDoc As String

    Public ChkDClick As Boolean = True
    Public chkItem As Boolean = False
    Public chkLoad As Boolean = False
    Public chkSelStk As Boolean = False

    Declare Function GetUserName Lib "advapi32.dll" Alias _
                  "GetUserNameA" (ByVal lpBuffer As String, _
                  ByRef nSize As Integer) As Integer
    Sub setYear()
        ' CultureInfo("th-TH")
        System.Threading.Thread.CurrentThread.CurrentCulture = System.Globalization.CultureInfo.CreateSpecificCulture("en-US")
        System.Threading.Thread.CurrentThread.CurrentUICulture = System.Threading.Thread.CurrentThread.CurrentCulture
    End Sub


    Public Function GetComputerName() As String
        Dim ComputerName As String
        ComputerName = Environment.MachineName ' System.Net.Dns.GetHostName
        Return ComputerName
    End Function



    Public Function GetUserName() As String 'เก็บUsername Passwordของเครื่องคนๆนั้น

        Dim iReturn As Integer
        Dim userName As String
        userName = New String(CChar(" "), 50)
        iReturn = GetUserName(userName, 50)
        GetUserName = userName.Substring(0, userName.IndexOf(Chr(0)))

    End Function
    Function getUserID(ByVal comName As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        Try
            If Trim(comName) = "" Then
                subDS = Nothing
                subDA = Nothing
                Return ""
            Else
                txtSQL = "Select * "
                txtSQL = txtSQL & "From csMast "

                txtSQL = txtSQL & "WHERE cs_comName='" & Trim(comName) & "'"

                subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
                subDA.Fill(subDS, "csCode")

                ans = subDS.Tables("csCode").Rows(0).Item("cs_code")
                subDS = Nothing
                subDA = Nothing
                Return ans

            End If
        Catch ex As Exception
            Return ""
        End Try



    End Function

    Sub openDB()

        'strConn = DBConnString.strConn2
        Try
            With Conn
                If .State = ConnectionState.Open Then .Close()
                .ConnectionString = strConn
                .Open()
            End With
        Catch ex As Exception
            MsgBox("ไม่สามารถติดต่อฐานข้อมูลได้")
        End Try


    End Sub

    Sub closeDB()
        Conn.Close()
    End Sub

    Function getTimeSer() As Date

        Dim Ans As Date
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "select getdate()as aDate ;"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "TimeList")

        Ans = subDS.Tables("TimeList").Rows(0).Item("aDate")
        subDS = Nothing
        subDA = Nothing
        Return Ans

    End Function
    'Sub DBConnection()
    '    'pathDB = "i:\center\PC50.mdb"
    '    'pathDB = "i:\center\PC50.mdb"
    '    'pathDB = "i:\test49\db\test49.mdb"
    '    'pathDB02 = "i:\center\acct50.mdb"
    '    'strConn = "server=(local);database=TestDB2006;Trusted_Connection=yes"
    '    'DB01 = DAODBEngine_definst.OpenDatabase(pathDB)
    '    ''DB02 = DAODBEngine_definst.OpenDatabase(pathDB02)
    '    strConn = "server=ENJOY-PC\sqlexpress;database=TestDB2006;Trusted_Connection=yes;" 'connectเครื่องตัวเอง
    '    'strConn = "server=EDP;database=DB2006;Trusted_Connection=yes;"
    '    'strConn = "Data Source=EDP2\SQLEXPRESS;Initial Catalog=db2006;User ID=sa;Password=sys0500" 'connect server
    '    'strConn = "Data Source=Time\SQLEXPRESS;Initial Catalog=backup_db2006;User ID=sa;Password=sys0500" 'connect server
    '    'strConn = "server=TIME\sqlexpress;database=backup_db2006;Trusted_Connection=yes;" 'connectเครื่องตัวเอง
    '    With Conn
    '        If .State = ConnectionState.Open Then .Close()
    '        .ConnectionString = strConn
    '        .Open()
    '    End With
    'End Sub

    Sub dbDelDATA(ByVal txtSQL As String, ByVal txtDisy As String)
        Try
            'If MessageBox.Show("ต้องการลบข้อมูล ' " & txtDisy & " ' ที่ระบุหรือไม่", "คำยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
            'DB01.Execute(txtSQL) ' บันทึกข้อมูลลง Business sc50
            'DB02.Execute(txtSQL) ' บันทึกข้อมูลลง Business acct50
            If Conn.State = ConnectionState.Closed Then
                DBtools.openDB()
            End If

            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With
            'End If
        Catch errprocess As Exception
            MessageBox.Show("ไม่สามารถลบข้อมูลได้เนื่องจาก " & errprocess.Message, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Sub dbDelSQLsrv(ByVal txtSQL As String, ByVal txtDisy As String)

        Try
            If MessageBox.Show("ต้องการลบข้อมูล ' " & txtDisy & " ' ที่ระบุหรือไม่", "คำยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                'DB01.Execute(txtSQL) ' บันทึกข้อมูลลง Business sc50
                'DB02.Execute(txtSQL) ' บันทึกข้อมูลลง Business acct50
                If Conn.State = ConnectionState.Closed Then
                    DBtools.openDB()
                End If
                With subCom
                    .CommandType = CommandType.Text
                    .CommandText = txtSQL
                    .Connection = Conn
                    .ExecuteNonQuery()
                End With
            End If
        Catch errprocess As Exception
            MessageBox.Show("ไม่สามารถลบข้อมูลได้เนื่องจาก " & errprocess.Message, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub
    Sub dbSaveSQLsrv(ByVal tSQL As String, ByVal txtDisy As String)
        DBtools.openDB()
        Try
            With subCom
                .CommandType = CommandType.Text
                .CommandText = tSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With
            DBtools.closeDB()

        Catch errprocess As Exception
            MessageBox.Show("ไม่สามารถเพิ่มข้อมูลได้เนื่องจาก " & errprocess.Message, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            'MessageBox.Show(tSQL)
            DBtools.closeDB()
            Exit Sub
        End Try

    End Sub

    Sub dbSaveDATA(ByVal txtSQL As String, ByVal txtNo As String)

        Try
            'If MessageBox.Show("ต้องการบันทึกข้อมูล ' " & txtDisy & " ' ที่ระบุหรือไม่", "คำยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

            If Conn.State = ConnectionState.Closed Then
                DBtools.openDB()
            End If

            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
                SqlGetUserSave(txtNo, txtSQL)

            End With
            'DB01.Execute(txtSQL) ' บันทึกข้อมูลลง Business 
            'DB02.Execute(txtSQL) ' บันทึกข้อมูลลง Business acct50
            'End If
        Catch errprocess As Exception
            MessageBox.Show("ไม่สามารถเพิ่มข้อมูลได้เนื่องจาก " & errprocess.Message, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Sub SqlGetUserSave(txtNo As String, strSQL As String)

        txtSQL = "Insert into userhistory("
        txtSQL = txtSQL & "history_user,history_Status,history_date,history_time,History_head,History_SQL"
        txtSQL = txtSQL & ")"
        txtSQL = txtSQL & "Values("
        txtSQL = txtSQL & "'" & GetUserName() & "','บันทึกเอกสาร-" & txtNo & "-','"
        'txtSQL = txtSQL & (Format(DateAdd(DateInterval.Year, -543, Now), "MM/dd/yyyy")) & "','"
        'txtSQL = txtSQL & (Format(DateAdd(DateInterval.Year, -543, Now), "HH:mm:ss")) & "','"

        txtSQL = txtSQL & (((Month(Now) & "-" & DateAndTime.Day(Now) & "-" & Year(Now) - 43))) & "','"
        txtSQL = txtSQL & (Format((Now), "HH:mm:ss")) & "','"
        txtSQL = txtSQL & txtNo & "','$"
        txtSQL = txtSQL & "" & "$' "
        txtSQL = txtSQL & ")"

        With subCom

            .CommandType = CommandType.Text
            .CommandText = txtSQL
            .Connection = Conn
            .ExecuteNonQuery()

        End With

    End Sub


    'Sub dbSaveSQLsrv(ByVal txtSQL As String, ByVal txtDisy As String)

    '    Try
    '        'If MessageBox.Show("ต้องการบันทึกข้อมูล ' " & txtDisy & " ' ที่ระบุหรือไม่", "คำยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then

    '        If Conn.State = ConnectionState.Closed Then
    '            DBtools.openDB()
    '        End If

    '        With subCom
    '            .CommandType = CommandType.Text
    '            .CommandText = txtSQL
    '            .Connection = Conn
    '            .ExecuteNonQuery()
    '        End With
    '        'DB01.Execute(txtSQL) ' บันทึกข้อมูลลง Business 
    '        'DB02.Execute(txtSQL) ' บันทึกข้อมูลลง Business acct50
    '        'End If
    '    Catch errprocess As Exception
    '        MessageBox.Show("ไม่สามารถเพิ่มข้อมูลได้เนื่องจาก " & errprocess.Message, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '        Exit Sub
    '    End Try
    'End Sub

    'Sub readDB()
    '    Dim ansTB As String
    '    Dim ansF As String
    '    Dim sizeF As String
    '    Dim countTB As Integer
    '    Dim countF As Integer
    '    With DB01
    '        For countTB = 0 To DB01.TableDefs.Count - 1
    '            ansTB = .TableDefs(countTB).Name
    '            For countF = 0 To .TableDefs(countTB).Fields.Count - 1
    '                ansF = .TableDefs(countTB).Fields(countF).Name
    '                sizeF = Convert.ToString(.TableDefs(countTB).Fields(countF).Size)

    '            Next
    '        Next
    '    End With
    '    With Conn

    '    End With
    'End Sub
    Sub dbSaveUser(ByVal txtSQL As String, ByVal txtDisy As String)


        Try

            If Conn.State = ConnectionState.Closed Then
                DBtools.openDB()
            End If
            With subCom
                .CommandType = CommandType.Text
                .CommandText = txtSQL
                .Connection = Conn
                .ExecuteNonQuery()
            End With



        Catch errprocess As Exception
            MessageBox.Show("ไม่สามารถเพิ่มข้อมูลได้เนื่องจาก " & errprocess.Message, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End Try
    End Sub

    Function getSaleName(ByVal saleId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From salesman "

        txtSQL = txtSQL & "WHERE SL_id='" & saleId & "'"


        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "slList")

        If subDS.Tables("slList").Rows.Count - 1 < 0 Then
            ans = ""
        Else
            ans = subDS.Tables("slList").Rows(0).Item("sl_Name")
        End If

        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Public Function ThaiBaht(ByVal pamt As Double) As String
        Dim i As Integer, j As Integer
        'Dim v As Integer
        Dim Valstr As String, Vlen As String, Vno As String
        'Dim syslge As String
        Dim wnumber(10) As String, wdigit(10) As String, spcdg(5) As String
        Dim vword(20) As String
        If pamt <= 0 Then
            ThaiBaht = ""
            Exit Function
        End If
        Valstr = Trim(Format$(pamt, "##########0.00"))
        Vlen = Len(Valstr) - 3
        For i = 1 To 20
            vword(i) = ""
        Next i
        wnumber(1) = "หนึ่ง" : wnumber(2) = "สอง" : wnumber(3) = "สาม" : wnumber(4) = "สี่" : wnumber(5) = "ห้า"
        wnumber(6) = "หก" : wnumber(7) = "เจ็ด" : wnumber(8) = "แปด" : wnumber(9) = "เก้า" : wdigit(1) = "บาท"
        wdigit(2) = "สิบ" : wdigit(3) = "ร้อย" : wdigit(4) = "พัน" : wdigit(5) = "หมื่น" : wdigit(6) = "แสน" : wdigit(7) = "ล้าน"
        spcdg(1) = "สตางค์" : spcdg(2) = "เอ็ด" : spcdg(3) = "ยี่" : spcdg(4) = "ถ้วน"
        For i = 1 To Vlen
            Vno = Int(Val(Mid$(Valstr, i, 1)))
            If Vno = 0 Then
                vword(i) = ""

                If (Vlen - i + 1) = 7 Then
                    vword(i) = wdigit(7) 'ล้าน
                End If
            Else
                If (Vlen - i + 1) > 7 Then
                    j = Vlen - i - 5 ' เกินหลักล้าน
                Else
                    j = Vlen - i + 1 'หลักแสน
                End If
                vword(i) = wnumber(Vno) + wdigit(j)  '30-90
            End If

            If Vno = 1 And j = 2 Then
                vword(i) = wdigit(2) 'สิบ
            End If
            If Vno = 2 And j = 2 Then
                vword(i) = spcdg(3) + wdigit(j) 'ยี่สิบ
            End If
            If j = 1 Then
                vword(i) = wnumber(Vno)
                If Vno = 1 And Vlen > 1 Then
                    If Mid$(Valstr, i - 1, 1) <> "0" Then
                        vword(i) = spcdg(2)
                    End If
                End If
            End If
            If j = 7 Then
                vword(i) = wnumber(Vno) + wdigit(j) 'ล้าน
                If Vno = 1 And Vlen > 7 Then
                    If Mid$(Valstr, i - 1, 1) <> "0" Then
                        vword(i) = spcdg(2) + wdigit(j)
                    End If
                End If
            End If
        Next i
        If Int(pamt) > 0 Then
            vword(Vlen) = vword(Vlen) + wdigit(1)
        End If


        Valstr = Mid$(Valstr, Vlen + 2, 2) 'ทศนิยม
        Vlen = Len(Valstr)
        For i = 1 To Vlen
            Vno = Int(Val(Mid$(Valstr, i, 1)))
            If Vno = 0 Then
                vword(i + 10) = ""
            Else
                j = Vlen - i + 1
                vword(i + 10) = wnumber(Vno) + wdigit(j)
                If Vno = 1 And j = 2 Then
                    vword(i + 10) = wdigit(2)
                End If
                If Vno = 2 And j = 2 Then
                    vword(i + 10) = spcdg(3) + wdigit(j)
                End If
                If j = 1 Then
                    If Vno = 1 And Int(Val(Mid$(Valstr, i - 1, 1))) <> 0 Then
                        vword(i + 10) = spcdg(2)
                    Else
                        vword(i + 10) = wnumber(Vno)
                    End If
                End If
            End If

        Next i
        If pamt <> 0 Then
            If Val(Valstr) = 0 Then
                vword(13) = spcdg(4)
            Else
                vword(13) = spcdg(1)
            End If
        End If
        Valstr = ""
        For i = 1 To 20
            Valstr = Valstr + vword(i)
        Next i
        ThaiBaht = (Valstr)
    End Function

    '=====================   Function  เสริม ใช้สอบถามค่าต่างๆ ใน DataBase ============================
    Function getArVat(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_VAT")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getCusLimit(ByVal ar_Code As String) As Integer

        Dim ans As Integer = 0
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Term")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getArAddr1(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getArAddr2(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr_1")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDocPO(orderNo As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")
        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")) Then
                ans = ""
            Else

                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")
            End If
        Else
            ans = ""
        End If
       

        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDocPO(orderNo As String, stkCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And TranDataD.Dtl_idTrade='" & stkCode & "' "
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")
        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")) Then
                ans = ""
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_PO")
            End If
        Else
            ans = ""
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getOrComment(orderNo As String, stkCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And TranDataD.Dtl_idTrade='" & stkCode & "' "
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")
        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_detail")) Then
                ans = ""
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_Detail")
            End If
        Else
            ans = ""
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDiscTrH(orderNo As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataH "

        txtSQL = txtSQL & "WHERE (Trh_Type='B')"
        txtSQL = txtSQL & "And TranDataH.Trh_No='" & orderNo & "' "


        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")

        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Trh_Disc")) Then
                ans = 0
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Trh_Disc")
            End If
        Else
            ans = 0
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getDiscProd(orderNo As String, stkCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataD "

        txtSQL = txtSQL & "WHERE (dtl_Type='B')"
        txtSQL = txtSQL & "And TranDataD.Dtl_idTrade='" & stkCode & "' "
        txtSQL = txtSQL & "And  (((TranDataD.Dtl_con_id) = '" & orderNo & "')) "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "OrderList")

        If subDS.Tables("OrderList").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("OrderList").Rows(0).Item("Dtl_T_Disc")) Then
                ans = 0
            Else
                ans = subDS.Tables("OrderList").Rows(0).Item("Dtl_T_Disc")
            End If
        Else
            ans = 0
        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getArAddr3(ByVal ar_Code As String) As String
        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & CusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr_2")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getArAddrFull(ByVal ar_Code As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & ar_Code & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Addr") & " " & subDS.Tables("ARList").Rows(0).Item("Ar_Addr_1") & " " & subDS.Tables("ARList").Rows(0).Item("Ar_Addr_2")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getTypeVATbyT(RevNo As String) As String

        Dim txtSQL As String = ""
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataH "
        txtSQL = txtSQL & "Where trh_Type='T' "
        txtSQL = txtSQL & "And trh_NO='" & RevNo & "' "



        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dbOrder")


        If subDS.Tables("dbOrder").Rows.Count = 0 Then
            Return "N"
        Else
            If IsDBNull(subDS.Tables("dbOrder").Rows(0).Item("trh_NoType")) Then
                Return "V"
            Else

                Return subDS.Tables("dbOrder").Rows(0).Item("trh_NoType")
            End If


        End If

    End Function
    Function getCusName(ByVal cusId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_C_Term,Ar_Target,Ar_Cre_Lim "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Name")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function
    Function getSalesCode(ByVal cusId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_Sales,Ar_C_Term,Ar_Target,Ar_Cre_Lim "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("Ar_Sales")
        subDS = Nothing
        subDA = Nothing

        Return ans

    End Function

    Function getArDisc(ByVal cusId As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_Sales,AR_DISC,Ar_Target,Ar_Cre_Lim "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusId & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")

        ans = subDS.Tables("ARList").Rows(0).Item("AR_DISC")
        subDS = Nothing
        subDA = Nothing

        Return ans

    End Function

    Function getStkWight(ByVal stkCode As String) As Double
        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If stkCode = "" Then
            ans = 0

        Else

            txtSQL = "Select * "
            txtSQL = txtSQL & "From BaseMast "

            txtSQL = txtSQL & "WHERE (((Stk_Code)='" & stkCode & "'))"

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "command")


            If IsNumeric(subDS.Tables("command").Rows(0).Item("Stk_Factor")) Then
                ans = subDS.Tables("command").Rows(0).Item("Stk_Factor")
            Else
                ans = 0
            End If

        End If


        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function

    Function getDocNumber(ByVal DocNo As String, ByVal DocType As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(DocNo) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            DBtools.openDB()
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataH "

            txtSQL = txtSQL & "WHERE ((Trh_Type='" & DocType & "') "
            txtSQL = txtSQL & "And (Trh_No='" & DocNo & "')) "
            'txtSQL=txtSQL & "And Trh_Wh='" & "'"  ' ใส่ข้อมูล คลังสินค้า

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            DBtools.closeDB()
            Return ans
        End If

    End Function

    'Public Function GetExcelVersion() As String
    '    Dim excel As Object = Nothing
    '    Dim ver As Integer = 0
    '    Dim build As Integer
    '    Try
    '        excel = CreateObject("Excel.Application")
    '        ver = excel.Version
    '        build = excel.Build
    '    Catch ex As Exception
    '        'Continue to finally sttmt
    '    Finally
    '        Try
    '            Marshal.ReleaseComObject(excel)
    '        Catch
    '        End Try
    '        GC.Collect()
    '    End Try
    '    Return ver
    'End Function

    Function getDocNumberD(ByVal DocNo As String, ByVal DocType As String, ByVal stkCode As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(DocNo) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            DBtools.openDB()
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataD "

            txtSQL = txtSQL & "WHERE ((dtl_Type='" & DocType & "') "
            txtSQL = txtSQL & "And (dtl_No='" & DocNo & "')"
            txtSQL = txtSQL & "And (dtl_idTrade='" & stkCode & "') "
            txtSQL = txtSQL & ") "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            DBtools.closeDB()
            Return ans
        End If

    End Function

    Function getDocType(ByVal docType As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TypeDocMast "

        txtSQL = txtSQL & "WHERE (((Type_ID) Like '%" & docType & "%'))"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "command")

        ans = subDS.Tables("command").Rows(0).Item("Type_Name")
        subDS = Nothing
        subDA = Nothing
        Return ans

    End Function


    Function getWhName(ByVal WhCode As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(WhCode) = "" Then
            Return ""
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From WareHouse "

            txtSQL = txtSQL & "WHERE Wh_id='" & WhCode & "' "
            txtSQL = txtSQL & "And Wh_Type='DC' "


            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "Ans")

            If subDS.Tables("Ans").Rows.Count > 0 Then
                ans = subDS.Tables("Ans").Rows(0).Item("Wh_Name")
            Else
                MsgBox("ไม่พบข้อมูล DC ที่ต้องการ")
                ans = ""
            End If

            subDS = Nothing
            subDA = Nothing
            Return ans

        End If


    End Function

    Function getStkNoDoc(ByVal DocType As String, ByVal DocNo As String, ByVal stkCode As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        If Trim(DocNo) = "" Then
            Return False
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From TranDataD "

            txtSQL = txtSQL & "WHERE Dtl_Type='" & DocType & "' "
            txtSQL = txtSQL & "And Dtl_No='" & DocNo & "' "
            txtSQL = txtSQL & "And Dtl_Stk='" & stkCode & "' "


            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            Return ans
        End If

    End Function

    Function getCostByStk(ByVal stkCode As String, ByVal CSDate As Date, strDate As String, ByVal chkRun As Integer) As Double

        Dim txtSQL As String = ""
        Dim subDa As SqlClient.SqlDataAdapter
        Dim subDs As DataSet = New DataSet
        Dim Ans As Double
        Dim strDatePr As String = ""

        txtSQL = "Select * "
        txtSQL = txtSQL & "From CostMast "
        txtSQL = txtSQL & "Where CS_Stk_Code='" & stkCode & "' "

        If chkRun = 0 Then
            chkRun = chkRun + 1
            strDatePr = Microsoft.VisualBasic.Right(Year(CSDate) - 543, 2) & "/" & Format(Month(CSDate), "00")
            txtSQL = txtSQL & "And CS_Date='" & strDatePr & "' "
            ramDateCost = strDatePr  ' เก็บค่า Date ของต้นทุน เป็น public เพื่อนำไปใส่ในรายงานให้รู้ว่าดึงต้นทุนจากวันไหน
        ElseIf chkRun = 24 Then

            Ans = 0
            Return Ans

        Else

            ' strDatePr = Microsoft.VisualBasic.Right(Year(CSDate) - 543, 2) & "/" & (Format(Month(CSDate), "00")) - 1 & "' "
            chkRun = chkRun + 1
            Dim strYY As String = Left(strDate, 2)
            Dim strMM As String = Format(Int(Right(strDate, 2)) - 1, "00").ToString

            If strYY = "12" Then
                Ans = 0
                Return Ans

            Else

                If strMM = "00" Then
                    strMM = "12"
                    strYY = Format((Int(strYY) - 1), "00").ToString
                    strDatePr = strYY & "/" & strMM
                Else
                    strDatePr = Left(strDate, 2) & "/" & Format(Int(Right(strDate, 2)) - 1, "00").ToString

                End If
            End If

            txtSQL = txtSQL & "And CS_Date='" & strDatePr & "' "
            ramDateCost = strDatePr  ' เก็บค่า Date ของต้นทุน เป็น public เพื่อนำไปใส่ในรายงานให้รู้ว่าดึงต้นทุนจากวันไหน
        End If
        '================================================== old


        'If CSDate = "01/01/2013" Then
        'Else

        '    txtSQL = txtSQL & "And CS_Date='" & Microsoft.VisualBasic.Right(Year(CSDate) - 543, 2) & "/" & Format(Month(CSDate), "00") & "' "
        'End If
        txtSQL = txtSQL & "Order by CS_Date desc "

        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "master")

        If chkRun = 0 Then
            If subDs.Tables("master").Rows.Count > 0 Then
                Ans = subDs.Tables("master").Rows(0).Item("CS_Cost")
                Return Ans
            Else
                Ans = getCostByStk(stkCode, "01/01/2558", strDatePr, chkRun)
                Return Ans '100 'getCostByStk(stkCode, "")
            End If
            '===============================================
        Else
            If subDs.Tables("master").Rows.Count > 0 Then
                Ans = subDs.Tables("master").Rows(0).Item("CS_Cost")
                Return Ans
            Else
                Ans = getCostByStk(stkCode, "01/01/2558", strDatePr, chkRun)
                Return Ans
                'Ans = 0
                'Return Ans '100 'getCostByStk(stkCode, "")
            End If

        End If
        'Catch ex As Exception
        '    MsgBox(ex.ToString)
        'End Try

    End Function

    Function getCusTaxCode(ByVal cusCode As String) As String

        Dim ans As String
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Ar_Cus_ID,Ar_Name,Ar_C_Term,Ar_Target,Ar_Cre_Lim,Ar_Tax_Code "
        txtSQL = txtSQL & "From ArFile "

        txtSQL = txtSQL & "WHERE (((ArFile.AR_Cus_ID) Like '%" & cusCode & "%'))"
        txtSQL = txtSQL & "And (Ar_Type='AR') "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "ARList")
        If IsDBNull(subDS.Tables("ARList").Rows(0).Item("Ar_Tax_Code")) Then
            ans = ""
        Else
            ans = subDS.Tables("ARList").Rows(0).Item("Ar_Tax_Code")
        End If

        subDS = Nothing
        subDA = Nothing
        Return ans


    End Function

    Function getCostType(ByVal stkCode As String) As Integer

        Dim txtSQL As String = ""
        Dim subDa As SqlClient.SqlDataAdapter
        Dim subDs As DataSet = New DataSet
        Dim Ans As Integer

        txtSQL = "Select * "
        txtSQL = txtSQL & "From BaseMast "
        txtSQL = txtSQL & "Where Stk_Code='" & stkCode & "'"


        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "master")

        If subDs.Tables("master").Rows.Count > 0 Then

            If IsDBNull(subDs.Tables("master").Rows(0).Item("Stk_Type")) Then

                If IsDBNull(subDs.Tables("master").Rows(0).Item("Stk_Equ")) Then
                    Ans = 0 'subDs.Tables("master").Rows(0).Item("Stk_Equ")
                Else
                    Ans = subDs.Tables("master").Rows(0).Item("Stk_Equ")

                End If

            Else

                Ans = subDs.Tables("master").Rows(0).Item("Stk_Type")
            End If


            Return Ans
        Else
            Return 9 '  ไม่มีค่า

        End If

    End Function


    Function getChkStkDetl(ByVal StrCode As String, ByVal StkCode As String) As Boolean
        Dim ans As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(StkCode) = "" Or Trim(StrCode) = "" Then
            Return False
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From StkDetl "

            txtSQL = txtSQL & "WHERE ((Dtl_Store='" & StrCode & "') "
            txtSQL = txtSQL & "And (Dtl_Code='" & StkCode & "')) "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                ans = True
            Else
                ans = False
            End If

            subDS = Nothing
            subDA = Nothing
            Return ans
        End If

    End Function

    Function getProDCode(ByVal StkCode As String) As String
        'Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(StkCode) = "" Then
            Return ""
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From BaseMast "

            txtSQL = txtSQL & "WHERE  "
            txtSQL = txtSQL & "(Stk_Code='" & StkCode & "') "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                Return subDS.Tables("StkList").Rows(0).Item("Stk_ProD").ToString
            Else
                Return ""
            End If

        End If

    End Function

    Function getChkBaseMast(ByVal StkCode As String) As Boolean

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(StkCode) = "" Or Trim(StkCode) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From BaseMast "
            txtSQL = txtSQL & "WHERE (Stk_Code='" & StkCode & "')"

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "stkList")

            If subDS.Tables("stkList").Rows.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End If

    End Function

    Function getStkName(ByVal stkId As String) As String

        Dim ans As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        Try
            If Trim(stkId) = "" Then
                subDS = Nothing
                subDA = Nothing
                Return ""
            Else
                txtSQL = "Select * "
                txtSQL = txtSQL & "From BaseMast "

                txtSQL = txtSQL & "WHERE Stk_Code='" & stkId & "'"

                subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
                subDA.Fill(subDS, "stkList")

                ans = subDS.Tables("stkList").Rows(0).Item("stk_Name_1")
                subDS = Nothing
                subDA = Nothing
                Return ans

            End If
        Catch ex As Exception
            Return ""
        End Try



    End Function
    Sub updatePrint(docNo As String)

        '  Trh_chk_print เท่ากับ  1  หมายถึง  พิมพ์ไปแล้ว

        txtSQL = "Update TranDataH "
        txtSQL = txtSQL & "Set Trh_Chk_Print =1 "
        txtSQL = txtSQL & "Where Trh_Type='S' "
        txtSQL = txtSQL & "And Trh_NO='" & docNo & "' "

        DBtools.dbSaveDATA(txtSQL, "")

    End Sub

    'Function getStock(ByVal stkId As String, ByVal strID As String, ByVal stkWH As String) As Double 'สำหรับหา Stock

    '    Dim ans As String

    '    Dim subDA As SqlClient.SqlDataAdapter
    '    Dim subDS As New DataSet

    '    If Trim(stkId) = "" Then
    '        Return 0 '""
    '        Exit Function
    '    Else
    '        txtSQL = "Select * "
    '        txtSQL = txtSQL & "From StkDetl "

    '        txtSQL = txtSQL & "WHERE (((Dtl_Code) Like '%" & stkId & "%')) "
    '        txtSQL = txtSQL & "And (((Dtl_Store) Like '%" & strID & "%')) "
    '        txtSQL = txtSQL & "And (Dtl_WH='" & stkWH & "')"

    '        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
    '        subDA.Fill(subDS, "stkList")

    '        If subDS.Tables("stkList").Rows.Count > 0 Then
    '            ans = subDS.Tables("stkList").Rows(0).Item("Dtl_Bal_Q1")
    '        Else
    '            ans = 0
    '        End If

    '        subDS = Nothing
    '        subDA = Nothing
    '        Return ans

    '    End If


    'End Function

    Function getStock(ByVal stkId As String, ByVal strID As String, ByVal whCode As String) As Integer
        Dim ans As Integer = 0
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet
        Try
            If Trim(stkId) = "" Then
                subDS = Nothing
                subDA = Nothing
                Return ""
            Else
                txtSQL = "Select Dtl_Wh,Dtl_Code,Dtl_Bal_Q1 "
                txtSQL = txtSQL & "From StkDetl "

                txtSQL = txtSQL & "WHERE Dtl_Code='" & stkId & "' "
                txtSQL = txtSQL & "And Dtl_Store='110098' "
                txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "' "

                'txtSQL = txtSQL & "group by Dtl_Wh,Dtl_Code "

                subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
                subDA.Fill(subDS, "StkStock")

                If subDS.Tables("StkStock").Rows.Count > 0 Then
                    ans = subDS.Tables("StkStock").Rows(0).Item("Dtl_Bal_Q1")
                Else
                    ans = 0
                End If

                subDS = Nothing
                subDA = Nothing

                Return ans
            End If
        Catch ex As Exception

            Return 0
        End Try

    End Function

    '=====================   Function  เสริม ใช้สอบถามค่าต่างๆ ใน DataBase ============================
    Public Sub rightPrint(ByVal strNum As Double, ByVal PositionX As Integer, ByVal PositionY As Integer)
        Dim txtAns1, txtAns2, txt As String
        Dim i As Integer
        txtAns1 = Str(Format(strNum, "#,###,###.00"))
        txtAns1 = Format(txtAns1, "#,###,###.00")

        For i = 0 To Len(txtAns1) - 1
            txt = (Right(txtAns1, 1))
            If txt = "," Then
                PositionX = PositionX + 50
            End If
            'Printer.CurrentX = PositionX
            'Printer.CurrentY = PositionY
            'Printer.Print(txt)
            txtAns2 = Left(txtAns1, Len(txtAns1) - 1) 'ทำเพื่อ ตัดตัวที่พิมพ์ไปแล้ว
            txtAns1 = txtAns2
            PositionX = PositionX - 98
        Next i
    End Sub

    Function getPending(ByVal cusCode As String, ByVal stkCode As String) As Double

        Dim ans As Double

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(stkCode) = "" Or Trim(cusCode) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return False
        Else
            txtSQL = "Select Dtl_idCus,Dtl_idTrade,sum(Dtl_Num-Dtl_Num_2)as penDing "
            txtSQL = txtSQL & "From TranDataD "

            txtSQL = txtSQL & "Where Dtl_idCus='" & cusCode & "'"
            txtSQL = txtSQL & "And Dtl_idTrade='" & stkCode & "'"
            txtSQL = txtSQL & "And Dtl_Type='B' "
            txtSQL = txtSQL & "Group by Dtl_idCus,Dtl_idTrade "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "PendingTB")
            If subDS.Tables("PendingTB").Rows.Count > 0 Then
                ans = subDS.Tables("PendingTB").Rows(0).Item("penDing")
                Return ans
            Else
                subDS = Nothing
                subDA = Nothing
                Return 0
            End If


        End If
    End Function

    ' ใช้ Update  Stkdetl  แบบสำเร็จ
    Sub updateStock(ByVal CusID As String, ByVal whCode As String, ByVal StkCode As String, ByVal OrderQ1 As Double, ByVal RcvQ1 As Double, ByVal IssQ1 As Double, ByVal PenDingUpdate As Boolean)

        Dim subDA3 As New SqlClient.SqlDataAdapter
        Dim subDS3 As New DataSet

        Dim wStk As Double = DBtools.getStkWight(StkCode)

        Dim Dtl_Bal_q2 As Double = 0
        Dim Dtl_Bal_q1 As Double = 0
        Dim Dtl_LS_Q1 As Double = 0
        Dim Dtl_LS_Q2 As Double = 0

        txtSQL = "Select * From StkDetl "
        txtSQL = txtSQL & "Where Dtl_Code='" & StkCode & "' "
        txtSQL = txtSQL & "And  Dtl_Store='" & CusID & "' "
        txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "'"

        subDA3 = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA3.Fill(subDS3, "chkStk")


        If subDS3.Tables("chkStk").Rows.Count > 0 Then
            With subDS3.Tables("chkStk").Rows(0)

                If IsDBNull(.Item("Dtl_Bal_q1")) Then
                    Dtl_Bal_q1 = 0
                Else
                    Dtl_Bal_q1 = .Item("Dtl_Bal_q1")
                End If

                If IsDBNull(.Item("Dtl_Bal_q2")) Then
                    Dtl_Bal_q2 = 0
                Else
                    Dtl_Bal_q2 = .Item("Dtl_Bal_q2")
                End If

                If IsDBNull(.Item("Dtl_LS_Q1")) Then
                    Dtl_LS_Q1 = 0
                Else
                    Dtl_LS_Q1 = .Item("Dtl_LS_Q1")
                End If

                If IsDBNull(.Item("Dtl_LS_Q2")) Then
                    Dtl_LS_Q2 = 0
                Else
                    Dtl_LS_Q2 = .Item("Dtl_LS_Q2")
                End If

                txtSQL = "Update StkDetl Set "
                If IsDBNull(.Item("Dtl_Or_Q1")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((0 + OrderQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((.Item("Dtl_Or_Q1") + OrderQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_Or_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((0 + (OrderQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((.Item("Dtl_Or_Q1") + OrderQ1) * wStk) & ","
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q1")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((0 + RcvQ1)) & ","
                    RcvQ1 = (0 + RcvQ1)                    '   แก้ไข 2-1-57  kritpon
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((.Item("Dtl_Rcv_Q1") + RcvQ1)) & ","
                    RcvQ1 = (.Item("Dtl_Rcv_Q1") + RcvQ1)  '   แก้ไข 2-1-57  kritpon
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & ((0 + (RcvQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & (.Item("Dtl_Rcv_Q1") + RcvQ1) * wStk & ","
                End If

                If IsDBNull(.Item("Dtl_iss_Q1")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((0 + IssQ1)) & ","
                    IssQ1 = ((0 + IssQ1))        '   แก้ไข 2-1-57  kritpon
                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((.Item("Dtl_iss_Q1") + IssQ1)) & ","
                    IssQ1 = ((.Item("Dtl_iss_Q1") + IssQ1))  '   แก้ไข 2-1-57  kritpon
                End If

                If IsDBNull(.Item("Dtl_iss_Q2")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((0 + (IssQ1 * wStk))) & ","

                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((.Item("Dtl_iss_Q1") + IssQ1)) * wStk & ","

                End If
                '  ไม่เกี่ยวกับยอดยกมา  ยอดยกมาห้าม update 
                'txtSQL = txtSQL & "Dtl_LS_Q1=" & Dtl_LS_Q1 & ","
                'txtSQL = txtSQL & "Dtl_LS_Q2=" & (Dtl_LS_Q1 * wStk) & ","
                'txtSQL = txtSQL & "Dtl_LS_Q1=0,"   '& (((Dtl_LS_Q1 + dtl_Bal_Q1 + RcvQ1) - IssQ1) * -1) & ","
                'txtSQL = txtSQL & "Dtl_LS_Q2=0,"   '& (((Dtl_Bal_q1 + RcvQ1) - IssQ1) * -1) * wStk & ","

                '==================================================================
                If ((Dtl_LS_Q1 + Dtl_Bal_q1 + RcvQ1) - IssQ1) > 0 Then

                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) * wStk & " "

                Else
                    MsgBox("ข้อมูลสต๊อกมีปัญหา โปรดแจ้ง ICT ด่วน ", MsgBoxStyle.Critical, "แจ้งเตีอน")

                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & (((Dtl_LS_Q1 + RcvQ1) - IssQ1)) * wStk & " "

                End If

                If PenDingUpdate = True Then
                    txtSQL = txtSQL & ",Dtl_Pen_Q1=" & ((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & ","
                    txtSQL = txtSQL & "Dtl_Pen_Q2=" & (((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "
                End If

                txtSQL = txtSQL & "Where Dtl_Store='110098' "
                txtSQL = txtSQL & "And Dtl_Code='" & StkCode & "' "
                txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "' "

            End With

        Else

            txtSQL = "Insert Into StkDetl "
            txtSQL = txtSQL & "("
            txtSQL = txtSQL & "Dtl_Wh,Dtl_Store,Dtl_Code,"

            txtSQL = txtSQL & "Dtl_Or_Q1,Dtl_Or_Q2,"    ' จอง
            txtSQL = txtSQL & "Dtl_Rcv_Q1,Dtl_Rcv_Q2,"  ' ผลิต
            txtSQL = txtSQL & "Dtl_ISS_Q1,Dtl_ISS_Q2,"  ' ขาย

            txtSQL = txtSQL & "Dtl_LS_Q1,Dtl_LS_Q2, "   ' ยกมา
            txtSQL = txtSQL & "Dtl_Bal_Q1,Dtl_Bal_Q2 "

            'txtSQL = txtSQL & "Dtl_Pen_Q1,Dtl_Pen_Q2 "
            txtSQL = txtSQL & ")"
            txtSQL = txtSQL & " Values"
            txtSQL = txtSQL & "('" & whCode & "',"
            txtSQL = txtSQL & "'110098','" & StkCode & "',"

            txtSQL = txtSQL & (OrderQ1) & "," & (((OrderQ1 * wStk))) & ","   '  จอง
            txtSQL = txtSQL & (RcvQ1) & "," & (RcvQ1 * wStk) & ","   '  ผลิต
            txtSQL = txtSQL & (IssQ1) & "," & (IssQ1 * wStk) & ","   '  ขาย

            txtSQL = txtSQL & 0 & "," & 0 & ","     '  ยกมา
            txtSQL = txtSQL & (RcvQ1 - IssQ1) & "," & ((RcvQ1 - IssQ1)) * wStk & " " '  Stock

            'If (RcvQ1 - IssQ1) < 0 Then

            'Else
            '    txtSQL = txtSQL & 0 & "," & 0 & ","  '  ยกมา
            '    txtSQL = txtSQL & (RcvQ1 - IssQ1) & "," & ((RcvQ1 - IssQ1) * wStk) & ","     '  Stock
            'End If
            'txtSQL = txtSQL & ((dbTools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & "," 'pen_q1
            'txtSQL = txtSQL & (((dbTools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "  '  pen

            txtSQL = txtSQL & ")"

        End If
        Call DBtools.dbSaveDATA(txtSQL, "")
        subDS3 = Nothing
        subDA3 = Nothing

    End Sub

    Sub updateStock2(ByVal CusID As String, ByVal whCode As String, ByVal StkCode As String, ByVal OrderQ1 As Double, ByVal RcvQ1 As Double, ByVal IssQ1 As Double, ByVal PenDingUpdate As Boolean)

        Dim subDA3 As New SqlClient.SqlDataAdapter
        Dim subDS3 As New DataSet

        Dim wStk As Double = DBtools.getStkWight(StkCode)

        Dim Dtl_Bal_q2 As Double = 0
        Dim Dtl_Bal_q1 As Double = 0
        Dim Dtl_LS_Q1 As Double = 0
        Dim Dtl_LS_Q2 As Double = 0



        txtSQL = "Select * From StkDetl "
        txtSQL = txtSQL & "Where Dtl_Code='" & StkCode & "' "
        txtSQL = txtSQL & "And  Dtl_Store='" & CusID & "' "
        txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "'"

        subDA3 = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA3.Fill(subDS3, "chkStk")


        If subDS3.Tables("chkStk").Rows.Count > 0 Then
            With subDS3.Tables("chkStk").Rows(0)

                If IsDBNull(.Item("Dtl_Bal_q1")) Then
                    Dtl_Bal_q1 = 0
                Else
                    Dtl_Bal_q1 = .Item("Dtl_Bal_q1")
                End If
                If IsDBNull(.Item("Dtl_Bal_q2")) Then
                    Dtl_Bal_q2 = 0
                Else
                    Dtl_Bal_q2 = .Item("Dtl_Bal_q2")
                End If

                If IsDBNull(.Item("Dtl_LS_Q1")) Then
                    Dtl_LS_Q1 = 0
                Else
                    Dtl_LS_Q1 = .Item("Dtl_LS_Q1")
                End If
                If IsDBNull(.Item("Dtl_LS_Q2")) Then
                    Dtl_LS_Q2 = 0
                Else
                    Dtl_LS_Q2 = .Item("Dtl_LS_Q2")
                End If

                txtSQL = "Update StkDetl Set "
                If IsDBNull(.Item("Dtl_Or_Q1")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((0 + OrderQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & ((.Item("Dtl_Or_Q1") + OrderQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_Or_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((0 + (OrderQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Or_Q2=" & ((.Item("Dtl_Or_Q2") + (OrderQ1 * wStk))) & ","
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((0 + RcvQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q1=" & ((.Item("Dtl_Rcv_Q2") + RcvQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_Rcv_Q2")) Then
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & ((0 + (RcvQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_Rcv_Q2=" & ((.Item("Dtl_Rcv_Q2") + (RcvQ1 * wStk))) & ","
                End If

                If IsDBNull(.Item("Dtl_iss_Q1")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((0 + IssQ1)) & ","
                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q1=" & ((.Item("Dtl_iss_Q1") + IssQ1)) & ","
                End If

                If IsDBNull(.Item("Dtl_iss_Q2")) Then
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((0 + (IssQ1 * wStk))) & ","
                Else
                    txtSQL = txtSQL & "Dtl_iSS_Q2=" & ((.Item("Dtl_iss_Q2") + (IssQ1 * wStk))) & ","
                End If

                If ((Dtl_LS_Q1 + Dtl_Bal_q1 + RcvQ1) - IssQ1) > 0 Then
                    txtSQL = txtSQL & "Dtl_LS_Q1=0" & ","
                    txtSQL = txtSQL & "Dtl_LS_Q2=0" & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & (Dtl_LS_Q1 + Dtl_Bal_q1 + RcvQ1) - IssQ1 & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & (((Dtl_LS_Q2 + Dtl_Bal_q2 + RcvQ1) - IssQ1) * wStk) & " "
                Else
                    txtSQL = txtSQL & "Dtl_LS_Q1=" & (((.Item("Dtl_LS_Q1") + .Item("Dtl_Bal_q1") + RcvQ1) - IssQ1) * -1) & ","
                    txtSQL = txtSQL & "Dtl_LS_Q2=0" & (((.Item("Dtl_LS_Q1") + .Item("Dtl_Bal_q1") + RcvQ1) - IssQ1) * -1) * wStk & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & ((.Item("Dtl_LS_Q1") + .Item("Dtl_Bal_q1") + RcvQ1) - IssQ1) * -1 & ","
                    txtSQL = txtSQL & "Dtl_Bal_Q2=" & ((((.Item("Dtl_LS_Q2") + .Item("Dtl_Bal_q2") + RcvQ1) - IssQ1)) * wStk) * -1 & " "
                End If

                If PenDingUpdate = True Then
                    txtSQL = txtSQL & ",Dtl_Pen_Q1=" & ((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & ","
                    txtSQL = txtSQL & "Dtl_Pen_Q2=" & (((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "
                End If
                txtSQL = txtSQL & "Where Dtl_Store='" & CusID & "' "
                txtSQL = txtSQL & "And Dtl_Code='" & StkCode & "' "
                txtSQL = txtSQL & "And Dtl_Wh='" & whCode & "' "

            End With

        Else
            txtSQL = "Insert Into StkDetl "
            txtSQL = txtSQL & "("
            txtSQL = txtSQL & "Dtl_Wh,Dtl_Store,Dtl_Code,"

            txtSQL = txtSQL & "Dtl_Or_Q1,Dtl_Or_Q2,"    ' จอง
            txtSQL = txtSQL & "Dtl_Rcv_Q1,Dtl_Rcv_Q2,"  ' ผลิต
            txtSQL = txtSQL & "Dtl_ISS_Q1,Dtl_ISS_Q2,"  ' ขาย

            txtSQL = txtSQL & "Dtl_LS_Q1,Dtl_LS_Q2, "   ' ยกมา
            txtSQL = txtSQL & "Dtl_Bal_Q1,Dtl_Bal_Q2,"

            txtSQL = txtSQL & "Dtl_Pen_Q1,Dtl_Pen_Q2 "
            txtSQL = txtSQL & ")"
            txtSQL = txtSQL & " Values"
            txtSQL = txtSQL & "('" & whCode & "',"
            txtSQL = txtSQL & "'" & CusID & "','" & StkCode & "',"

            txtSQL = txtSQL & (OrderQ1) & "," & (((OrderQ1 * wStk))) & ","   '  จอง
            txtSQL = txtSQL & (RcvQ1) & "," & (RcvQ1 * wStk) & ","   '  ผลิต
            txtSQL = txtSQL & (IssQ1) & "," & (IssQ1 * wStk) & ","   '  ขาย

            If (RcvQ1 - IssQ1) < 0 Then
                txtSQL = txtSQL & (RcvQ1 - IssQ1) * -1 & "," & ((RcvQ1 - IssQ1) * -1) * wStk & ","  '  ยกมา
                txtSQL = txtSQL & 0 & "," & 0 & ","     '  Stock
            Else
                txtSQL = txtSQL & 0 & "," & 0 & ","  '  ยกมา
                txtSQL = txtSQL & (RcvQ1 - IssQ1) & "," & ((RcvQ1 - IssQ1) * wStk) & ","     '  Stock
            End If
            txtSQL = txtSQL & ((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) & "," 'pen_q1
            txtSQL = txtSQL & (((DBtools.getPending(CusID, StkCode) + OrderQ1) - IssQ1) * wStk) & " "  '  pen

            txtSQL = txtSQL & ")"

        End If
        Call DBtools.dbSaveDATA(txtSQL, "")
        subDS3 = Nothing
        subDA3 = Nothing

    End Sub

    Sub updateStkdetl(ByVal dtlType As String, ByVal dtlStr As String, ByVal dtlWH As String, ByVal dtlCode As String, ByVal dtlQty As Integer)

        Dim f As String = ""
        'Dim stkCode As String = "010103001230184000001011"
        Dim subDaZ As SqlClient.SqlDataAdapter
        Dim subDsZ As DataSet = New DataSet
        Dim iLoop As Integer = 0
        'Dim strcode As String = "110098"
        'f = dbTools.chkDtlFlag(stkcode, strcode)
        DBtools.openDB()
        txtSQL = "Select * from StkDetl "
        txtSQL = txtSQL & "where Dtl_Code='" & dtlCode & "' "
        txtSQL = txtSQL & "And Dtl_Store='" & dtlStr & "' "
        txtSQL = txtSQL & "And Dtl_Wh='" & dtlWH & "' "

        subDaZ = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDaZ.Fill(subDsZ, "DtlSet")

        Dim issQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_iSS_Q1")
        Dim rcvQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_RCV_Q1")
        Dim lsQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_LS_Q1")
        Dim orQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_Or_Q1")
        Dim balQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_Bal_Q1")
        Dim penQ1 As Integer = subDsZ.Tables("DtlSet").Rows(iLoop).Item("Dtl_Pen_Q1")


        If subDsZ.Tables("DtlSet").Rows.Count > 0 Then

            txtSQL = "Update StkDetl Set "
            Select Case dtlType
                Case "S"
                    txtSQL = txtSQL & "Dtl_iss_Q1=" & issQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Pen_Q1=" & orQ1 + dtlQty - issQ1 & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 - dtlQty & " "
                Case "D"
                    txtSQL = txtSQL & "Dtl_RCV_Q1=" & rcvQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 + dtlQty & " "
                Case "G"
                    txtSQL = txtSQL & "Dtl_RCV_Q1=" & rcvQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 + dtlQty & " "
                Case "B"
                    txtSQL = txtSQL & "Dtl_Or_Q1=" & orQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Pen_Q1=" & orQ1 + dtlQty - issQ1 & " "
                Case "F"
                    txtSQL = txtSQL & "Dtl_RCV_Q1=" & rcvQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 + dtlQty & " "
                Case "Z"
                    txtSQL = txtSQL & "Dtl_iss_Q1=" & issQ1 + dtlQty & " "
                    txtSQL = txtSQL & "Dtl_Bal_Q1=" & lsQ1 + rcvQ1 - issQ1 - dtlQty & " "

            End Select

            txtSQL = txtSQL & "Where Dtl_Store='" & dtlStr & "' "
            txtSQL = txtSQL & "And Dtl_Code='" & dtlCode & "' "
            txtSQL = txtSQL & "And Dtl_Wh='" & dtlWH & "' "

            DBtools.dbSaveDATA(txtSQL, "")
        End If
        DBtools.closeDB()

    End Sub

    Function getRunID(trhNo As String) As Integer   '  เก็บข้อมูล ของหัวบิล  ลงตัวแปร เพื่อจะใช้ในการบันทึก 

        'CusId = Trim(lbCusID.Text)  '  รหัสลูกค้า
        'DocID = Trim(txtNo.Text)    '  เลขที่เอกสารขาย
        'whCode = lbWHcode.Text      '  รหัสคลังสินค้า
        'VatType = cboTypeDoc.SelectedValue    '  ประเภท VAT

        If Len(trhNo) > 0 Then
            trh_RunID = Microsoft.VisualBasic.Right(trhNo, 4)
        ElseIf Len(trhNo) = 9 Then
            'strDoc = Microsoft.VisualBasic.Right(txtNo.Text, 4)
            trh_RunID = Microsoft.VisualBasic.Right(trhNo, 4)
        End If
        If IsNumeric(trh_RunID) = False Then
            trh_RunID = 0
        End If

        Return trh_RunID

    End Function
    Function chkDataAmphoe(cusID As String) As Boolean

        Dim subDa As SqlClient.SqlDataAdapter
        Dim subDs As DataSet = New DataSet
        Dim Ans As Boolean = False

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ArFile "
        txtSQL = txtSQL & "Where Ar_Cus_ID='" & cusID & "' "

        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "arList")

        If IsDBNull(subDs.Tables("arList").Rows(0).Item("Ar_Amphoe_Code")) Then
            Return False
        Else
            Return True

        End If

    End Function

    Function getChkPrint(docID As String) As Boolean

        Dim subDa As SqlClient.SqlDataAdapter
        Dim subDs As DataSet = New DataSet
        Dim Ans As Boolean = False

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataH "

        txtSQL = txtSQL & "Where Trh_Type='S' "
        txtSQL = txtSQL & "And Trh_No='" & docID & "' "

        subDa = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDa.Fill(subDs, "TrhList")



        If IsDBNull(subDs.Tables("TRhList").Rows(0).Item("Trh_chk_Print")) Then
            Return False
        Else
            Return subDs.Tables("TRhList").Rows(0).Item("Trh_chk_Print")

        End If

    End Function


    Function getSaleByArFile(ByVal CusCode As String) As String

        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        If Trim(CusCode) = "" Or Trim(CusCode) = "" Then
            subDS = Nothing
            subDA = Nothing
            Return ""
        Else
            txtSQL = "Select * "
            txtSQL = txtSQL & "From ArFile "
            txtSQL = txtSQL & "WHERE Ar_Cus_ID='" & CusCode & "' "

            subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
            subDA.Fill(subDS, "ArList")

            If subDS.Tables("ArList").Rows.Count > 0 Then
                Return subDS.Tables("ArList").Rows(0).Item("Ar_Sales")
            Else
                Return ""

            End If
        End If

    End Function


    'Public Function ceiling(ByVal strvat As Decimal) As Decimal
    '    Math.Ceiling(strvat)
    '    Return strvat
    'End Function
End Module
