Public Class frmBegin

    Dim strDocNo_1 As String = ""
    Dim strDocNo_2 As String = ""
    Dim strClockUpdate As String = ""

    'Dim dateStart As DateTime
    Dim intQty1 As Integer
    Dim intQty2 As Integer

    Dim chkLoad As Boolean = False
    Dim intCounter As Integer = 0
    Dim intCounterX As Integer = 0
    Dim chkCounter As Boolean = False
    Dim EndTime As DateTime
    Dim StartTime As DateTime
    ' Dim countQty As Integer = 0
    Dim time_S As Double = 0
    Dim time_M As Double = 0
    Dim intStdCounter0 As Integer = 0
    Dim intStdCounter1 As Integer = 0
    Dim intStdCounter2 As Integer = 0
    Dim timeClockMast As DateTime
    Private _isPause As Boolean
    Private _isCountBack As Boolean
    Private _timeAll As Integer
    Private _timeAll0 As Integer
    Dim txtTime01 As String
    Dim txtTime02 As String
    Dim txtTime03 As String
    Dim txtTime04 As String
    Dim strDate01 As Date
    Dim strDate02 As Date
    Sub clsText()
        lbPCName.Text = ""
        lbDocNo02.Text = ""
        lbStartTime.Text = "00:00:00"
        lbEndTime.Text = "00:00:00"
        lbTotalTime.Text = "00:00:00"
        lbTimer1.Text = "00:00:00"
        lbTime_Use.Text = "00"
        lbTime_Counter.Text = "00"
        lbCryCleTime.Text = "00:00:00"
        lbTime_AVG.Text = "00:00:00"
        lbQtyCount.Text = "00"

        lbWeight.Text = "00"
        lbWeight2.Text = "00"
        lbQty.Text = "00"
        lbQty2.Text = "00"
        lbQty3.Text = "00"
        lbWeight_True.Text = "00"

    End Sub


    Sub chkTimeText()

        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ClockMast "
        txtSQL = txtSQL & "Order by Clock_Update desc "
        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "clock")


        Dim overTime As Integer = subDS.Tables("clock").Rows(0).Item("Clock_OverTime")
        If overTime = 0 Then
            txtTime01 = "08:00:00"
            txtTime02 = "17:00:00"
            txtTime03 = "17:01:00"
            txtTime04 = "02:00:00"
        ElseIf overTime = 1 Then
            txtTime01 = "08:01:00"
            txtTime02 = "18:00:00"
            txtTime03 = "18:01:00"
            txtTime04 = "04:00:00"
        ElseIf overTime = 2 Then
            txtTime01 = "08:00:00"
            txtTime02 = "19:00:00"
            txtTime03 = "19:01:00"
            txtTime04 = "06:00:00"
        ElseIf overTime = 3 Then
            txtTime01 = "08:01:00"
            txtTime02 = "20:00:00"
            txtTime03 = "20:01:00"
            txtTime04 = "08:00:00"
        End If

    End Sub

    Sub showData(strDocNo As String)
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter

        Dim intQty_Work2 As Integer = 0
        Dim intQty_Work1 As Integer = 0
        Dim dblTime_Page As Double = 0
        Dim dblTimeProcess As Double = 0

        Dim tmpTrhNO As String = "" '.Rows(0).Item("Trh_No")
        Dim tmpStkCode As String = "" '.Rows(0).Item("dtl_idTrade")
        Dim tmpPCname As String = "" '.Rows(0).Item("Dtl_n_Trade")
        ' Dim tmpTimer As DateTime 'cdate(Format(.Rows(0).Item("Trh_Cre_Lim"), "00:00") & ":00.0")
        Dim docDate2 As String
        docDate2 = Format(Now, "dd/MM/yyyy HH:mm:ss").ToString
        docDate2 = Format(DateAdd(DateInterval.Year, -543, CDate(docDate2)), "MM/dd/yyyy HH:mm:ss")

        txtSQL = "Select * "
        txtSQL = txtSQL & "From ClockMast "
        txtSQL = txtSQL & "Order by Clock_Update desc "
        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "clock")

        Dim intNight As Integer = subDS.Tables("clock").Rows(0).Item("Clock_Night")

        lbLock.Text = subDS.Tables("clock").Rows(0).Item("Clock_Lock")
        lbSection.Text = subDS.Tables("clock").Rows(0).Item("Clock_Section")
        'lbDate.Text = Format(DateAdd(DateInterval.Year, -543, subDS.Tables("clock").Rows(0).Item("Clock_Update")), "dd-MM-yyyy")
        lbDate.Text = Format(DateAdd(DateInterval.Year, -543, CDate(strClockUpdate)), "dd-MM-yyyy")

        If lbLock.Text = "1" Then
            lbTextStatus.Text = "stop ปรับแผนผลิต"
            lbTextStatus.ForeColor = Color.Red
            Exit Sub
            ' Timer1.Enabled = False
        Else
            lbTextStatus.Text = "สถานะ ทำงานปกติ "
            lbTextStatus.ForeColor = Color.GreenYellow
        End If

        txtSQL = "Select  * "
        txtSQL = txtSQL & "From TranDataH_E "
        txtSQL = txtSQL & "left Join TranDataD_E "
        txtSQL = txtSQL & "On (TranDataH_E.Trh_Type=TranDataD_E.Dtl_Type "
        txtSQL = txtSQL & "And TranDataH_E.Trh_NO=TranDataD_E.Dtl_No )"

        txtSQL = txtSQL & "Where Trh_type='E' "
        txtSQL = txtSQL & "And Trh_no='" & strDocNo & "'"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dataList")
        If subDS.Tables("dataList").Rows(0).Item("Dtl_Num") = 0 Then
            Exit Sub
        End If

        txtSQL = "select  BOM_RM_Updatetime,BOM_RM_Values,Dtl_num_2/Dtl_Num,(Dtl_num_2/Dtl_num)-(((Dtl_num_2/dtl_Num) *10)/100) as tt "
        txtSQL = txtSQL & "From BOMmastF "
        txtSQL = txtSQL & "Left Join TranDataH_E "
        txtSQL = txtSQL & "On BOM_No=Trh_No "
        txtSQL = txtSQL & "Left Join TranDataD_E "
        txtSQL = txtSQL & "On Trh_No=Dtl_no "

        txtSQL = txtSQL & "Where BOM_RM_Scales=5 "
        txtSQL = txtSQL & "And BOM_No='" & strDocNo & "' "
        txtSQL = txtSQL & "And BOM_RM_Values > (Dtl_num_2/Dtl_num)-(((Dtl_num_2/dtl_Num) *10)/100) "
        txtSQL = txtSQL & "Order by BOM_RM_Updatetime desc "

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "chkQty")
        intQty_Work2 = subDS.Tables("chkQty").Rows.Count


        If subDS.Tables("dataList").Rows.Count > 0 Then

            With subDS.Tables("dataList")

                If chkCounter = False Then

                    Timer1.Enabled = True
                    '  Set ค่าสีตัวอักษรเริ่มต้นนับเวลาผลิต
                    lbTimer1.ForeColor = Color.Lime

                    'เวลาเริ่มในการผลิตชุดงานใบคุมนี้
                    'subDS.Tables("chkQty").Rows(intQty_Work2 - 1).Item("BOM_RM_Updatetime") 'BOM_RM_Updatetime'Now

                    If subDS.Tables("chkQty").Rows.Count > 0 Then  ' จุดนี้คือ ที่คาดว่า Error ทีหาไม่เจอ
                        StartTime = subDS.Tables("chkQty").Rows(intQty_Work2 - 1).Item("BOM_RM_Updatetime") 'BOM_RM_Updatetime'Now
                    Else
                        StartTime = Now
                    End If

                    lbStartTime.Text = StartTime
                    tmpTrhNO = .Rows(0).Item("Trh_No")   '  เลขที่เอกสาร
                    ' If chkTrhWorkPrint(trh_No, "M") = False Then
                    'If tmpTrhNO = "56919CA" Then
                    '    MsgBox("56919CA")

                    'End If

                    txtDocNo.Text = tmpTrhNO
                    tmpStkCode = .Rows(0).Item("dtl_idTrade")  ' รหัสสินค้า  SA
                    tmpPCname = .Rows(0).Item("Dtl_n_Trade")   '  ชื่อชุดงาน PC 

                    If subDS.Tables("chkQty").Rows.Count > 0 Then
                        lbWeight_True.Text = subDS.Tables("chkQty").Rows(0).Item("BOM_RM_Values")   '  น้ำหนักเท ในแผ่นนั้น
                    Else
                        lbWeight_True.Text = 0
                    End If

                    lbDocNo.Text = tmpTrhNO                     '  แสดงผล เลขที่เอกสาร
                    lbPCName.ForeColor = Color.Lime
                    lbPCName.Text = tmpPCname                   '  แสดงผล ชื่อชุดงาน

                    ' ==============  จำนวนที่ผลิตไปแล้ว ====================
                    intQty_Work2 = subDS.Tables("chkQty").Rows.Count    '  นับจำนวนแผ่นในเลขที่ใบคุมนี้ทึ่ผลิตไปแล้ว
                    intQty1 = intQty_Work2   '  เก็บไว้เทียย
                    intQty_Work1 = .Rows(0).Item("Dtl_Num")             '   จำนวนแผ่นที่ต้องผลิตของใบคุมนี้
                    lbWeight.Text = .Rows(0).Item("Dtl_Num_2")
                    lbQty.Text = .Rows(0).Item("Dtl_Num")
                    lbWeight2.Text = Format((.Rows(0).Item("Dtl_Num_2") / .Rows(0).Item("Dtl_Num")), "#,##0.0")
                    '===========================================

                    time_M = .Rows(0).Item("Trh_Cre_Lim") '   Crycle time ในการผลิต
                    lbTrh_Cre_lim.Text = (time_M * 60) / intQty_Work1
                    time_S = time_M * 60  '  เวลา Crycle Time เป็น วินาที
                    lbTotalTime.Text = changeInt2Time(time_S)

                    lbCryCleTime.Text = Format(time_S / intQty_Work1, "#,##0.00")
                    dblTime_Page = time_S / .Rows(0).Item("Dtl_Num")    '   เวลาวินาที-ผลิตต่อแผ่น
                    'dblTimeProcess = (intQty_Work1 - intQty_Work2) * (dblTime_Page / 60)   'ผิดอยู่   เวลาผลิตที่ต้องใช้สำหรับแผ่นที่เหลือ หาร 60 ทำให้เป็น นาที

                    time_S = time_S - ((time_S / intQty_Work1) * (intQty_Work2 - 1))

                    EndTime = Format(DateAdd(DateInterval.Second, time_S, Now), "dd-MM-yyyy HH:mm:ss")
                    lbEndTime.Text = EndTime
                    '==================================================
                    '_timeAll = dblTimeProcess
                    'lbTimeLimit.Text = changeTxtTime(dblTimeProcess)     '  แสดงผลที่หน้าจอ
                    '===================================================
                    lbPCName.Text = tmpPCname '.Rows(0).Item("dtl_n_Trade")

                    'Call saveTranDataD("M", strDocNo, "110098", "แพน", "01", "", 1, tmpStkCode, tmpPCname, intQty_Work1, (intQty_Work1 * lbWeight2.Text), 1, tmpStkCode, lbDate.Text, lbSection.Text)
                    chkCounter = True
                End If   ' ********************  จบ if ที่เป็น loop เริ่มต้นแผ่นแรกของใบคุม *********************

                intQty1 = subDS.Tables("chkQty").Rows.Count  '  แสดงจำนวนที่เพิ่มขึ้น
                lbQty3.Text = lbQty.Text - intQty1
                lbQty2.Text = intQty1

                txtSQL = "select  BOM_RM_Values,Dtl_num_2/Dtl_Num,(Dtl_num_2/Dtl_num)-(((Dtl_num_2/dtl_Num) *10)/100) as tt "
                txtSQL = txtSQL & "From BOMmastF "
                txtSQL = txtSQL & "Left Join TranDataH_E "
                txtSQL = txtSQL & "On BOM_No=Trh_No "
                txtSQL = txtSQL & "Left Join TranDataD_E "
                txtSQL = txtSQL & "On Trh_No=Dtl_no "

                txtSQL = txtSQL & "Where BOM_RM_Scales=5 "
                txtSQL = txtSQL & "And BOM_No='" & strDocNo & "' "
                'txtSQL = txtSQL & "And BOM_RM_Values > (Dtl_num_2/Dtl_num)-(((Dtl_num_2/dtl_Num) *10)/100) "
                txtSQL = txtSQL & "Order by BOM_RM_Updatetime desc "

                subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
                subDA.Fill(subDS, "chkWeight")
                ' intQty_Work2 = subDS.Tables("chkWeight").Rows.Count
                If subDS.Tables("chkWeight").Rows.Count > 0 Then
                    lbWeight_True.Text = subDS.Tables("chkWeight").Rows(0).Item("BOM_RM_Values")   '  น้ำหนักเท ในแผ่นนั้น
                    'intQty1 = subDS.Tables("chkQty").Rows.Count  '  แสดงจำนวนที่เพิ่มขึ้น
                Else
                    lbWeight_True.Text = 0
                End If
                '======================  ทำเพื่อเช็คจำนวนการผลิต เพื่อหยุดการผลิต ถ้ามีการพักการทำงาน ==================

                intCounter = intCounter + 1

                '======================  ทำเพื่อเช็คจำนวนการผลิต เพื่อหยุดการผลิต ถ้ามีการพักการทำงาน ==================

                'lbTimer1.Text = intCounter
                '_timeAll = (_timeAll * 60 * 10) - intCounter
                'If _isCountBack = True Then
                '    _timeAll -= 1
                'Else

                'End If
                '_timeAll += 1

                If intQty1 = intQty2 Then
                    'intCounterX = intCounterX + 1
                    Dim intCounter As Integer = DateDiff(DateInterval.Second, Now, EndTime)
                    lbTime_Use.Text = DateDiff(DateInterval.Second, StartTime, EndTime)
                    lbTime_Counter.Text = intCounter ' DateDiff(DateInterval.Second, dateStart, Now)

                    If intCounter < 0 Then
                        lbTimer1.ForeColor = Color.Red
                        lbTimer1.Text = "00:00:00"
                    Else
                        lbTimer1.Text = changeInt2Time(intCounter) ' changeToTimeShowUp(intCounterX)

                    End If
                    lbQtyCount.Text = lbTime_Use.Text - lbTime_Counter.Text 'DateDiff(DateInterval.Second, StartTime, EndTime)
                    lbTime_AVG.Text = Format(lbQtyCount.Text / lbQty2.Text, "#,##0.00")
                Else


                    'countQty = countQty + 1
                    intQty2 = intQty1
                    'lbQtyCount.Text = countQty
                    ' intCounterX = 0
                End If

                Dim secCounter As Integer
                'Dim diffNow As Double

                Call chkTimeText()

                If intNight = 0 Then '  เช็คกลางคืนหรือไม่  ถ้า 1 เป็นกลางคืน

                    strDate01 = Format(Now, "dd-MM-yyyy HH:mm:ss")
                    strDate02 = Format(Now, "dd-MM-yyyy").ToString & " " & (txtTime02)

                    Dim ifDate01 As DateTime = Format(Now, "dd-MM-yyyy").ToString & " " & txtTime01
                    Dim ifDate02 As DateTime = Format(DateAdd(DateInterval.Day, 1, Now), "dd-MM-yyyy").ToString & " " & txtTime02
                    lbTotalSecTime.Text = Format(DateDiff(DateInterval.Second, ifDate01, ifDate02), "#,##0")

                ElseIf intNight = 1 Then  '  เช็คกลางคืนหรือไม่  ถ้า 1 เป็นกลางคืน

                    strDate01 = Format(Now, "dd-MM-yyyy HH:mm:ss")
                    strDate02 = Format(DateAdd(DateInterval.Day, 1, subDS.Tables("clock").Rows(0).Item("Clock_Update")), "dd-MM-yyyy").ToString & " " & (txtTime04)

                    Dim ifDate01 As DateTime = Format(Now, "dd-MM-yyyy").ToString & " " & txtTime03
                    Dim ifDate02 As DateTime = Format(DateAdd(DateInterval.Day, 1, Now), "dd-MM-yyyy").ToString & " " & txtTime04
                    lbTotalSecTime.Text = Format(DateDiff(DateInterval.Second, ifDate01, ifDate02), "#,##0")

                End If
                'lbTime1.Text = strDate01
                'lbTime1.Text = strDate01
                lbDate02.Text = strDate02

                'secCounter = DateDiff(DateInterval.Second, strDate01, strDate02)

                lbTotalPagePerTime.Text = Format(lbTotalSecTime.Text / lbTotalQty.Text, "#,##0.00")
                lbTotalWeightPerTime.Text = Format(lbTotalSecTime.Text / lbTotalWeight.Text, "#,##0.00")


                secCounter = DateDiff(DateInterval.Second, strDate01, strDate02)
                lbTimeClock.Text = changeInt2Time(secCounter)

            End With


        End If

        Call showSumQty()


    End Sub
    Sub showSumQty()
        Dim subDA As SqlClient.SqlDataAdapter
        Dim subDS As New DataSet

        txtSQL = "Select Trh_Type,sum(Dtl_Num)as sumQty,sum(Dtl_Num_2)as sumW "
        txtSQL = txtSQL & "From TranDataH "
        txtSQL = txtSQL & "left Join TranDataD "
        txtSQL = txtSQL & "On (TranDataH.Trh_Type=TranDataD.Dtl_Type "
        txtSQL = txtSQL & "And TranDataH.Trh_NO=TranDataD.Dtl_No )"

        txtSQL = txtSQL & "Where Trh_type='M' "
        'txtSQL = txtSQL & "And Year(Trh_Date)='" & Year(Now) - 543 & "' "
        'txtSQL = txtSQL & "And month(Trh_Date)='" & Month(Now) & "' "
        'txtSQL = txtSQL & "And DAy(Trh_Date)='" & Format(Now, "dd").ToString & "' "
        txtSQL = txtSQL & "And Trh_KeyType='" & lbSection.Text & "' "
        txtSQL = txtSQL & "And Trh_DateSection='" & lbDate.Text & "' "
        txtSQL = txtSQL & "group by Trh_type "


        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "sumQty")
        If subDS.Tables("sumQty").Rows.Count > 0 Then
            If IsDBNull(subDS.Tables("sumQty").Rows(0).Item("sumQty")) Then

            Else
                lbTotalQty.Text = subDS.Tables("sumQty").Rows(0).Item("sumQty").ToString
                lbTotalWeight.Text = subDS.Tables("sumQty").Rows(0).Item("sumW").ToString

            End If
        Else
            lbTotalQty.Text = 0 'subDS.Tables("sumQty").Rows(0).Item("sumQty").ToString
            lbTotalWeight.Text = 0 'subDS.Tables("sumQty").Rows(0).Item("sumW").ToString


        End If
    End Sub


    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick


        '============================================

        strDocNo_1 = getDocNo()
        lbDocID2.Text = strDocNo_1
        lbDocNo02.Text = strDocNo_1
        If chkType_E(strDocNo_1) = True Then

            'Timer3.Enabled = False
            lbTextStatus.Text = "การทำงานปกติ "
            lbTextStatus.ForeColor = Color.GreenYellow
            lbPCName.ForeColor = Color.Lime
            If strDocNo_1 = strDocNo_2 Then
                'lbDocNo02.Text = strDocNo_2
                _isCountBack = False     '/ถ้าอยากให้เดินหน้าก็กำหนดเป็น False
                ' _timeAll = 0
                _isPause = False
                'Call saveTranDataH("M", strDocNo_1, "110098", "102", "01", "", lbDate.Text, lbSection.Text)
                Call showData(strDocNo_1)
            Else
                Call clsText()
                ' Timer1.Enabled = False
                lbTimer1.Text = "00:00:00.0"
                '=============  บันทึก ===================
                ' txtSQL = "Insert Into TranDataD "
                'Call saveTranDataH("M", strDocNo_1, 110098, "102", "01", "", lbDate.Text, lbSection.Text)
                strDocNo_2 = strDocNo_1
                lbDocNo02.Text = strDocNo_2
                intCounter = 0
                ' Call setTimerStart()
                chkCounter = False
                Call showData(strDocNo_1)

            End If

        Else
            Call clsText()
            lbPCName.ForeColor = Color.Red
            lbTextStatus.Text = "ไม่พบข้อมูลใบคุมเลขที่ " & strDocNo_1 & " "
            lbPCName.Text = "ไม่พบข้อมูลใบคุมเลขที่ " & strDocNo_1 & " "
            '    Timer3.Enabled = True

        End If
        '=============================================================
        'strDocNo_1 = getDocNo()

        'If strDocNo_1 = strDocNo_2 Then
        '    'lbDocNo02.Text = strDocNo_2
        '    _isCountBack = False     '/ถ้าอยากให้เดินหน้าก็กำหนดเป็น False
        '    ' _timeAll = 0
        '    _isPause = False

        '    Call showData(strDocNo_1)

        'Else
        '    timeClockMast = getClockMast()
        '    Timer1.Enabled = False
        '    lbTimer1.Text = "00:00:00.0"
        '    countQty = 0
        '    '=============  บันทึก ===================
        '    ' txtSQL = "Insert Into TranDataD "
        '    Call saveTranDataH("M", strDocNo_1, 110098, "102", "01", "", lbDate.Text, lbSection.Text)
        '    strDocNo_2 = strDocNo_1
        '    lbDocNo02.Text = strDocNo_2
        '    intCounter = 0
        '    ' Call setTimerStart()
        '    chkCounter = False
        '    Call showData(strDocNo_1)

        'End If

        '========================================================================

    End Sub
    Function chkType_E(strDocNo As String) As Boolean
        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter

        txtSQL = "Select * "
        txtSQL = txtSQL & "From TranDataH_E "
        txtSQL = txtSQL & "left Join TranDataD_E "
        txtSQL = txtSQL & "On (TranDataH_E.Trh_Type=TranDataD_E.Dtl_Type "
        txtSQL = txtSQL & "And TranDataH_E.Trh_NO=TranDataD_E.Dtl_No )"

        txtSQL = txtSQL & "Where Trh_type='E' "
        txtSQL = txtSQL & "And Trh_no='" & strDocNo & "'"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dataTypeE")

        If subDS.Tables("dataTypeE").Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If

    End Function
    Function getClockUpdate() As String

        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        txtSQL = "Select * "
        txtSQL = txtSQL & "From ClockMast "
        txtSQL = txtSQL & "Order by Clock_Update desc "
        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "clock")

        Return subDS.Tables("clock").Rows(0).Item("Clock_Update")


    End Function
    Private Sub frmBegin_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        strClockUpdate = getClockUpdate()

        setTimerStart()

        strDocNo_1 = getDocNo()
        lbDocNo.Text = strDocNo_1
        strDocNo_2 = strDocNo_1
        lbDocNo02.Text = strDocNo_2
        Timer1.Enabled = True
        chkLoad = True

    End Sub
    Function getDocNo() As String

        Dim subDS As New DataSet
        Dim subDA As SqlClient.SqlDataAdapter
        Dim strAns As String = ""

        txtSQL = "Select * "
        txtSQL = txtSQL & "From BOMmastF "

        txtSQL = txtSQL & "Where year(BOM_RM_Update)=" & Year(Now) - 543 & " "
        txtSQL = txtSQL & "And BOM_RM_Scales='5' "
        txtSQL = txtSQL & "And month(BOM_RM_Update)=" & Month(Now) & " "

        '   --where (bom_RM_Update)='2019-02-25'
        txtSQL = txtSQL & "Order by BOM_RM_Updatetime desc"

        subDA = New SqlClient.SqlDataAdapter(txtSQL, Conn)
        subDA.Fill(subDS, "dataList")

        If subDS.Tables("dataList").Rows.Count > 0 Then
            With subDS.Tables("dataList").Rows(0)

                strAns = .Item("BOM_No")

            End With
        End If
        Return strAns
    End Function


    Sub saveTranDataD(strDocType As String, ByVal dtlNO As String, ByVal cusCode As String, cusName As String, ByVal whCode As String, ByVal docDate As String, running As Integer, ByVal stkCode As String, stkName As String, ByVal stkQty As Double, ByVal stkQty2 As Double, ByVal stkPrice As Double, ByVal stkCode2 As String, ByVal docDetail As String, ByVal docOrderID As String)

        Dim stkSumPrice As Double
        Dim runNumber As Integer
        stkPrice = 100 'getCostByStk(stkCode, docDate, 0)
        stkSumPrice = (stkQty * stkPrice)

        Dim docDate2 As String
        docDate2 = Format(Now, "dd/MM/yyyy HH:mm:ss").ToString
        docDate2 = Format(DateAdd(DateInterval.Year, -543, CDate(docDate2)), "MM/dd/yyyy HH:mm:ss")

        'Dim docDate2 As String = Format(DateAdd(DateInterval.Year, -543, CDate(docDate)), "MM/dd/yyyy HH:mm:ss")
        'docDate = Format(DateAdd(DateInterval.Year, -543, CDate(docDate)), "dd/MM/yyyy HH:mm:ss")
        Dim strDateS As String = Format(DateAdd(DateInterval.Year, -543, Now), "dd/MM/yyyy HH:mm:ss").ToString
        Dim strDateF As String = Format(DateAdd(DateInterval.Year, -543, Now), "dd/MM/yyyy HH:mm:ss").ToString

        runNumber = running
        If getDocNumberD(dtlNO, strDocType, stkCode) = True And running = 1 Then

            If stkCode = stkCode2 Then  '  check StkCode กับ  stkCode2 ถ้าไม่ตรงกัน แสดงว่ามีการเปลี่ยนสินค้า  stkCode2 เก็บรหัสสินค้าเก่า kritpon

                txtSQL = "Update TranDataD "
                txtSQL = txtSQL & "Set "
                'txtSQL = txtSQL & "Dtl_Date='" & docDate2 & "',"  '  วันที่
                txtSQL = txtSQL & "Dtl_idCus='" & cusCode & "'," '  รหัสลูกค้า
                txtSQL = txtSQL & "Dtl_Con_ID='" & docOrderID & "',"          '  เลขที่เอกสาร
                txtSQL = txtSQL & "Dtl_WH='" & whCode & "',"     '  คลังสินค้า 
                txtSQL = txtSQL & "dtl_idTrade='" & stkCode & "'," ' รหัสสินค้า
                txtSQL = txtSQL & "Dtl_n_Trade='" & stkName & "',"
                txtSQL = txtSQL & "Dtl_dateF='" & strDateS & "',"

                txtSQL = txtSQL & "dtl_Num=" & stkQty & ","         ' จำนวนสินค้า
                txtSQL = txtSQL & "dtl_Num_2=" & stkQty2 & ","         ' จำนวนสินค้า

                txtSQL = txtSQL & "Dtl_Same_Code='" & stkCode2 & "',"  '  รหัสสินค้าตรวจสอบ
                txtSQL = txtSQL & "Dtl_Price='" & stkPrice & "',"
                txtSQL = txtSQL & "Dtl_Sum='" & stkSumPrice & "',"
                txtSQL = txtSQL & "Dtl_Chk='" & 0 & "',"   '   ใช้ดูว่า เข้าผลิต(1) หรีอ  ออกจากการผลิต  (0)
                txtSQL = txtSQL & "Dtl_Detail='" & docDetail & "', "
                txtSQL = txtSQL & "Dtl_Date_F='" & docDate2 & "'  "

                txtSQL = txtSQL & "Where dtl_type='" & strDocType & "' "
                txtSQL = txtSQL & "And dtl_No='" & dtlNO & "'"
                txtSQL = txtSQL & "And dtl_idTrade='" & stkCode2 & "'"

                'DBtools.dbSaveSQLsrv(txtSQL, "")

            End If

        Else

            txtSQL = "Insert into TranDataD(Dtl_Type,Dtl_Date,Dtl_No,Dtl_idCus,"
            txtSQL = txtSQL & "Dtl_idTrade,DtL_N_Trade,dtl_num,dtl_num_2,Dtl_Price,Dtl_Sum,"
            txtSQL = txtSQL & "Dtl_con_id,Dtl_runnum,Dtl_cus_Name,"
            txtSQL = txtSQL & "Dtl_Same_Code,Dtl_DateS,Dtl_WH,Dtl_Chk,Dtl_Detail,Dtl_Date_S,Dtl_Date_F) "

            txtSQL = txtSQL & "Values("
            txtSQL = txtSQL & "'" & strDocType & "','" & docDate2 & "','" & dtlNO & "','" & cusCode & "',"
            txtSQL = txtSQL & "'" & stkCode & "','" & stkName & "','" & stkQty & "','" & stkQty2 & "','" & stkPrice & "','" & stkSumPrice & "',"
            txtSQL = txtSQL & "'" & docOrderID & "','" & runNumber & "','" & cusName & "',"
            txtSQL = txtSQL & "'" & stkCode2 & "','" & strDateS & "','" & whCode & "'," & 1 & ",'" & docDetail & "', "
            txtSQL = txtSQL & "'" & docDate2 & "','" & docDate2 & "' "
            txtSQL = txtSQL & ")"

            'DBtools.dbSaveSQLsrv(txtSQL, "")

        End If
        Dim storeCode As String = "103"
        ' เอกสารรับลงผลิต Auto

        'txtSQL = "Update TranDataD "
        'txtSQL = txtSQL & "Set Dtl_DateF='" & Now & "' "
        'txtSQL = txtSQL & "Where Dtl_type='M' and Dtl_DueDate='" & dtlNO & "' "
        'dbTools.dbSaveSQLsrv(txtSQL, "")

        If running = 1 Then
            Call saveTranDataH(strDocType, dtlNO, cusCode, storeCode, whCode, docDate2, lbDate.Text, lbSection.Text)
        End If


    End Sub
    Sub saveTranDataH(DocType As String, ByVal trhNO As String, ByVal cusCode As String, storeCode As String, ByVal whCode As String, ByVal docDate As String, strDetail As String, strSection As String)

        Dim strDate As String

        strDate = Format(Now, "dd/MM/yyyy HH:mm:ss").ToString
        strDate = Format(DateAdd(DateInterval.Year, -543, CDate(strDate)), "MM/dd/yyyy HH:mm:ss")
        storeCode = "102"
        If getDocNumber(trhNO, DocType) = True Then

            txtSQL = "Update TranDataH "
            txtSQL = txtSQL & "Set Trh_Cus='" & cusCode & "',"
            txtSQL = txtSQL & "Trh_Date='" & strDate & "',"
            txtSQL = txtSQL & "Trh_Sale='" & getSaleByArFile(cusCode) & "',"
            txtSQL = txtSQL & "Trh_wh='" & whCode & "',"
            txtSQL = txtSQL & "Trh_Chk_Print='0',"
            txtSQL = txtSQL & "Trh_KeyType='" & strSection & "',"
            txtSQL = txtSQL & "Trh_DateSection='" & strDetail & "',"
            txtSQL = txtSQL & "Trh_Store='" & storeCode & "' "

            txtSQL = txtSQL & "Where trh_type='" & DocType & "' "
            txtSQL = txtSQL & "And trh_No='" & trhNO & "' "
            txtSQL = txtSQL & "And Trh_Store='" & storeCode & "' "  ' 102  แผนก น้ำยา

            'DBtools.dbSaveSQLsrv(txtSQL, "")
        Else
            '  ใช้เพิ่มข้อมูลส่วนหัวเอกสารรับผลิต
            txtSQL = "Insert into TranDataH(Trh_Type,Trh_NO,Trh_KeyType,"
            txtSQL = txtSQL & "Trh_Cus,Trh_Store,Trh_Date,"
            txtSQL = txtSQL & "Trh_Amt,Trh_Full_Amt,"
            txtSQL = txtSQL & "Trh_Sale,Trh_Chk_Print,"
            txtSQL = txtSQL & "Trh_Wh,Trh_DateSection) "
            txtSQL = txtSQL & "Values('" & DocType & "','" & trhNO & "','" & strSection & "','" & cusCode & "','" & storeCode & "','"
            txtSQL = txtSQL & strDate & "'," & 0 & ",0,'" & getSaleByArFile(cusCode) & "','0','" & whCode & "','" & strDetail & "') "
            'DateAdd(DateInterval.Year, -543, Now)
            'DBtools.dbSaveSQLsrv(txtSQL, "")

        End If

    End Sub


    Private Sub txtDocNo_KeyDown(sender As Object, e As KeyEventArgs) Handles txtDocNo.KeyDown
        If e.KeyCode = Keys.Enter Then
            Call showData(txtDocNo.Text)
        End If
    End Sub

    Sub setTimerStart()

        'lbStatus1.Text = txtDocNo.Text & " เริ่มผลิต เวลา " & txtTime.Text
        'lbStatus1.Text = Format(CDate(txtTime.Text), "dd-MM-yyyy HH:mm:ss")

        'Timer1.Interval = 100
        '_isCountBack = False     '/ถ้าอยากให้เดินหน้าก็กำหนดเป็น False
        ''_isPause = False
        '_timeAll = 0
        'lbTimer1.Text = "00:00:00.0"
        'Timer1.Enabled = True
        '_isPause = False

    End Sub
    Function changeInt2Time(setTime As Integer) As String

        Dim ans As String = ""
        Dim total_millisec As Integer = setTime * 1000
        Dim millisecNow As Integer = total_millisec
        ' lbCounter.Text = intCounter
        Dim hour As Integer = Math.Floor(millisecNow / (60 * 60 * 1000)) 'ได้ชั่วโมง
        Dim min As Integer = Math.Floor(((millisecNow / (60 * 1000)) - (hour * 60)))
        Dim sec As Integer = Math.Floor((millisecNow / 1000) - (hour * 60 * 60) - (min * 60))
        'Dim mil As Integer '= 100 * (((millisecNow / 1000) - (hour * 60 * 60) - (min * 60)) - sec)
        'mil = Math.Floor(millisecNow / 100 - (hour * 60 * 60) - (min * 60) - (sec * 100))
        ' lbMilSec.Text = mil
        ans = hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00") '+ "." + mil.ToString("00")
        ' Application.DoEvents()
        Return ans

    End Function


    Function changeTxtTime(setTime As Integer) As String

        Dim ans As String = ""
        Dim total_millisec As Integer = setTime * 60 * 1000
        Dim millisecNow As Integer = total_millisec - intCounter * 100
        ' lbCounter.Text = intCounter
        Dim hour As Integer = Math.Floor(millisecNow / (60 * 60 * 1000)) 'ได้ชั่วโมง
        Dim min As Integer = Math.Floor((millisecNow / (60 * 1000) - (hour * 60)))
        Dim sec As Integer = Math.Floor((millisecNow / 1000) - (hour * 60) - (min * 60))

        'Dim hour As Integer = Math.Floor(millisecNow / (60 * 60 * 1000)) 'ได้ชั่วโมง
        'Dim min As Integer = Math.Floor((millisecNow / (60 * 1000) - (hour * 60)))
        'Dim sec As Integer = Math.Floor((millisecNow / 1000) - (hour * 60 * 60) - (min * 60))
        'Dim mil As Integer = 100 * (((millisecNow / 1000) - (hour * 60 * 60) - (min * 60)) - sec)
        Dim mil As Integer = 100 * (((millisecNow / 1000) - (hour * 60) - (min * 60)) - sec)
        ' lbMilSec.Text = mil
        ans = hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00") + "." + mil.ToString("00")
        ' Application.DoEvents()
        Return ans

    End Function
    Function changeToTimeUse(qty1 As Integer) As String

        Dim ans As String = ""
        Dim total_millisec As Integer = _timeAll * 60 '* 1000

        Dim millisecNow As Integer = total_millisec + intCounter * 100
        Dim use_millisec As Integer = (millisecNow) / qty1

        ' lbCounter.Text = intCounter
        'Dim hour As Integer = Math.Floor(use_millisec / (60 * 60 * 1000)) 'ได้ชั่วโมง
        'Dim min As Integer = Math.Floor((use_millisec / (60 * 1000) - (hour * 60)))
        Dim sec As Integer = Math.Floor((use_millisec / 1000))
        Dim mil As Integer = 100 * ((use_millisec / 1000) - sec)
        ' lbMilSec.Text = mil
        ans = sec.ToString("00") + "." + mil.ToString("00")
        ' Application.DoEvents()
        Return ans
        '=======================================================================


    End Function

    Function changeToTimeShowUp(intCounter0 As Integer) As String

        Dim ans As String = ""
        Dim total_millisec As Integer = 0 * 60 * 1000
        Dim millisecNow As Integer = total_millisec + intCounter0 * 100
        ' lbCounter.Text = intCounter
        'Dim hour As Integer = Math.Floor(millisecNow / (60 * 60 * 1000)) 'ได้ชั่วโมง
        ' Dim min As Integer = Math.Floor((millisecNow / (60 * 1000) - (hour * 60)))
        Dim sec As Integer = Math.Floor((millisecNow / 1000)) '- (hour * 60 * 60) - (min * 60)
        Dim mil As Integer = 100 * (((millisecNow / 1000) - sec)) '- (hour * 60 * 60) - (min * 60))
        ' lbMilSec.Text = mil
        'hour.ToString("00") + ":" + min.ToString("00") + ":" +
        ans = sec.ToString("00") + "." + mil.ToString("00")
        ' Application.DoEvents()
        Return ans
        '=======================================================================


    End Function

    Function changeToTimeShow() As String

        Dim ans As String = ""
        Dim total_millisec As Integer = _timeAll * 60 * 1000
        Dim millisecNow As Integer = total_millisec - intCounter * 100
        ' lbCounter.Text = intCounter
        Dim hour As Integer = Math.Floor(millisecNow / (60 * 60 * 1000)) 'ได้ชั่วโมง
        Dim min As Integer = Math.Floor((millisecNow / (60 * 1000) - (hour * 60)))
        Dim sec As Integer = Math.Floor((millisecNow / 1000) - (hour * 60 * 60) - (min * 60))
        Dim mil As Integer = 100 * (((millisecNow / 1000) - (hour * 60 * 60) - (min * 60)) - sec)
        ' lbMilSec.Text = mil
        ans = hour.ToString("00") + ":" + min.ToString("00") + ":" + sec.ToString("00") + "." + mil.ToString("00")
        ' Application.DoEvents()
        Return ans
        '=======================================================================


    End Function
    Dim chkTimer01 As Integer
    Private Sub Timer2_Tick(sender As Object, e As EventArgs) Handles Timer2.Tick
        lbTimeRuning.Text = Now
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles lbDate.Click

    End Sub

    Private Sub txtDocNo_TextChanged(sender As Object, e As EventArgs) Handles txtDocNo.TextChanged

    End Sub

    Private Sub cmbEdit_Click(sender As Object, e As EventArgs) Handles cmbEdit.Click
        Dim frmEdit As New frmEdit
        frmEdit.ShowDialog()

    End Sub
End Class
