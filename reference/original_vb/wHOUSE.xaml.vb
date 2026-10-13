Public Class wHOUSE
    Public HOUSE_CODE As String '라인코드 H1R1E

    Private HOUSE_HEADER As String 'H1, H2

    Private source_ip As String
    Private source_port As Integer
    Private target_ip As String
    Private target_port As Integer
    Private source_manual As Boolean
    Private manual_rpm As Integer
    Private source_line As String
    Private source_id As Integer

    Private Sub PreviewKeyDown_PreventHanguel(sender As Object, e As KeyEventArgs)
        If (e.Key = Key.Space Or e.Key = Key.ImeProcessed) Then e.Handled = True
    End Sub

    Private Sub PreviewTextInput_OnlyInteger(sender As Object, e As TextCompositionEventArgs)
        Dim regex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^0-9]+") ' [^0-9.-]+ : 점과빼기추가
        e.Handled = regex.IsMatch(e.Text)
    End Sub
    Private Sub PreviewTextInput_IP(sender As Object, e As TextCompositionEventArgs)
        Dim regex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^0-9.]+") ' [^0-9.-]+ : 점과빼기추가
        e.Handled = regex.IsMatch(e.Text)
    End Sub

    Private Sub wHOUSE_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        '라인명
        lbl_house.Content = Get_House_Name(HOUSE_CODE)
        HOUSE_HEADER = HOUSE_CODE.Substring(0, 2)

        '설정값
        source_ip = MCONFIG(HOUSE_CODE).SOURCE_IP
        source_port = MCONFIG(HOUSE_CODE).SOURCE_PORT
        target_ip = MCONFIG(HOUSE_CODE).TARGET_IP
        target_port = MCONFIG(HOUSE_CODE).TARGET_PORT
        source_manual = MCONFIG(HOUSE_CODE).SOURCE_MANUAL
        manual_rpm = MCONFIG(HOUSE_CODE).MANUAL_RPM
        source_line = MCONFIG(HOUSE_CODE).SOURCE_LINE
        source_id = MCONFIG(HOUSE_CODE).SOURCE_ID

        txtSIP.Text = source_ip
        txtSPORT.Text = source_port
        txtTIP.Text = target_ip
        txtTPORT.Text = target_port
        txtMRPM.Text = manual_rpm

        If source_line = "" Then source_line = HOUSE_CODE

        If (source_line.Length = 5) Then '설정값에 라인값이 있으면
            Dim tline_code As String = source_line.Substring(2, 3)
            cb_lines.SelectedValue = tline_code
            Load_OLD_IDS(HOUSE_HEADER & tline_code, source_id.ToString)
        Else
            cb_lines.SelectedValue = ""
        End If

        If (source_manual = True) Then
            rb_source_manual.IsChecked = True
            rb_source_ref.IsChecked = False
        Else '다른모터 참조
            rb_source_manual.IsChecked = False
            rb_source_ref.IsChecked = True
        End If

        lbl_rpm_range.Content = "정수(0~" & MAX_RPM & ")"

    End Sub

    Private Sub Load_OLD_IDS(lcode As String, Optional defaultvalue As String = "")

        cb_ids.Items.Clear()

        Dim ci As New ComboBoxItem
        ci.Tag = ""
        ci.Content = "ID"
        ci.IsSelected = True
        cb_ids.Items.Add(ci)

        Dim selidx As Integer = 0
        Dim tlist As List(Of String) = HOUSES(lcode).Get_OLDMOTOR_List
        For Each mid As String In tlist
            Dim ti As New ComboBoxItem
            ti.Tag = mid
            ti.Content = mid
            If (mid = defaultvalue) Then
                ti.IsSelected = True
            Else
                ti.IsSelected = False
            End If
            cb_ids.Items.Add(ti)
        Next


    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        '기존모터 접속 정보
        Dim sip As String = txtSIP.Text
        If (sip = "") Then
            MsgBox("기존모터 접속 IP를 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        If (Check_Valid_IP(sip) = False) Then
            MsgBox("기존모터 접속 IP가 올바른 형식이 아닙니다..", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        Dim sport As Integer
        If (txtSPORT.Text = "") Then
            MsgBox("기존모터 접속 PORT를 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        If (Check_Integer(txtSPORT.Text) = False) Then
            MsgBox("기존모터 접속 PORT를 숫자로 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        Integer.TryParse(txtSPORT.Text, sport)
        If (sport < 1 Or sport > 65535) Then
            MsgBox("기존모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        '신규모터 접속 정보
        Dim tip As String = txtTIP.Text
        If (tip = "") Then
            MsgBox("신규 모터 접속 IP를 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        If (Check_Valid_IP(tip) = False) Then
            MsgBox("신규모터 접속 IP가 올바른 형식이 아닙니다..", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        Dim tport As Integer
        If (txtTPORT.Text = "") Then
            MsgBox("신규모터 접속 PORT를 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        If (Check_Integer(txtTPORT.Text) = False) Then
            MsgBox("신규모터 접속 PORT를 숫자로 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        Integer.TryParse(txtTPORT.Text, tport)
        If (tport < 1 Or tport > 65535) Then
            MsgBox("신규모터 접속 PORT를 숫자로 1~65535범위로 입력해 주세요.", MsgBoxStyle.OkOnly, "에러")
            Exit Sub
        End If
        '원본모드
        Dim smanual As Boolean = False
        Dim mrpm As Integer = -1
        Dim sline As String = ""
        Dim sid As Integer = -1
        If (rb_source_manual.IsChecked = True) Then
            smanual = True
            If (Integer.TryParse(txtMRPM.Text, mrpm) = False) Then
                MsgBox("수동 RPM속도를 숫자로 입력해주세요.", MsgBoxStyle.OkOnly, "에러")
                Exit Sub
            End If
            If (mrpm < 0 Or mrpm > MAX_RPM) Then
                MsgBox("수동 RPM속도를 0~" & MAX_RPM & "범위내로 입력해주세요.", MsgBoxStyle.OkOnly, "에러")
                Exit Sub
            End If
            sline = cb_lines.SelectedValue
            If (cb_ids.SelectedValue <> "ID" And cb_ids.SelectedValue <> "") Then sid = cb_ids.SelectedValue
        ElseIf (rb_source_ref.IsChecked = True) Then
            Integer.TryParse(txtMRPM.Text, mrpm)
            smanual = False
            sline = cb_lines.SelectedValue
            If (sline = "") Then
                MsgBox("참조할 모터가 있는 라인을 선택해주세요.", MsgBoxStyle.OkOnly, "에러")
                Exit Sub
            End If
            If (cb_ids.SelectedValue = "ID" Or Integer.TryParse(cb_ids.SelectedValue, sid) = False) Then
                MsgBox("참조할 모터 ID를 선택해주세요.", MsgBoxStyle.OkOnly, "에러")
                Exit Sub
            End If
            If (sid < 0) Then
                MsgBox("참조할 모터 ID를 선택해주세요.", MsgBoxStyle.OkOnly, "에러")
                Exit Sub
            End If
        Else
            MsgBox("RPM제어 소스를 선택해주세요.")
            Exit Sub
        End If

        sline = HOUSE_HEADER & sline

        'MsgBox(sip & "/" & sport & "=" & tip & "/" & tport & " " & smanual & "/" & mrpm & "/" & sline & "/" & sid)

        Dim tmp_config As Struct_Config

        tmp_config.SOURCE_IP = sip
        tmp_config.SOURCE_PORT = sport
        tmp_config.TARGET_IP = tip
        tmp_config.TARGET_PORT = tport
        tmp_config.SOURCE_MANUAL = smanual
        tmp_config.MANUAL_RPM = mrpm
        tmp_config.SOURCE_LINE = sline
        tmp_config.SOURCE_ID = sid

        MCONFIG(HOUSE_CODE) = tmp_config

        Put_Local_Config()

        '부모창 갱신 함수 실행
        Dim parentWindow As MainWindow = Me.Owner
        parentWindow.Update_HeaderBox()

        Me.Close()
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As RoutedEventArgs)
        Me.Close()
    End Sub

    Private Sub rb_source_manual_Checked(sender As Object, e As RoutedEventArgs)
        manual_line.Visibility = Visibility.Visible
        ref_line.Visibility = Visibility.Collapsed
    End Sub

    Private Sub rb_source_ref_Checked(sender As Object, e As RoutedEventArgs)
        manual_line.Visibility = Visibility.Collapsed
        ref_line.Visibility = Visibility.Visible

    End Sub

    Private Sub cb_lines_SelectionChanged(sender As Object, e As SelectionChangedEventArgs)

        Debug.Print(cb_lines.SelectedValue)

        If (cb_lines.SelectedValue <> "") Then
            Dim hcode As String = HOUSE_HEADER & cb_lines.SelectedValue
            Load_OLD_IDS(hcode)
        Else
            cb_ids.Items.Clear()
            Dim ci As New ComboBoxItem
            ci.Tag = ""
            ci.Content = "ID"
            ci.IsSelected = True
            cb_ids.Items.Add(ci)
        End If



    End Sub

    Private Sub Border_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If (e.ChangedButton = MouseButton.Left) Then
            Me.DragMove()
        End If
    End Sub
End Class
