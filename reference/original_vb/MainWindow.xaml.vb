Imports System.ComponentModel
Imports System.Windows.Threading

Class MainWindow
    Private TMR_UPDATE As New System.Windows.Threading.DispatcherTimer

    Private H1_NETWORK_RUN As Boolean = False
    Private H2_NETWORK_RUN As Boolean = False

    Private Sub MainWindow_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        '이미 실행중인지 체크
        Dim strModName As String = My.Application.Info.ProductName
        Dim approc = Process.GetProcessesByName(strModName)
        If approc.Length > 1 Then
            Me.Hide() '창 안보이게
            MsgBox("모니터가 이미 실행중입니다.")
            End '프로그램 종료
        End If

        '시작/종료 버튼 추기화
        H1_StartBtn.Visibility = Visibility.Visible
        H1_StopBtn.Visibility = Visibility.Collapsed
        H2_StartBtn.Visibility = Visibility.Visible
        H2_StopBtn.Visibility = Visibility.Collapsed


        '설정 읽기
        Get_Local_Config()

        Load_Data()


        '타이머
        TMR_UPDATE = New DispatcherTimer()
        TMR_UPDATE.Interval = New TimeSpan(0, 0, 5) '5초
        AddHandler TMR_UPDATE.Tick, AddressOf Me.TMR_UPDATE_TICK
        '타이머 시작
        TMR_UPDATE.Start()

        '헤더정보 표시
        Update_HeaderBox()



    End Sub

    Private Sub MainWindow_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing

        If (MsgBox("종료하시겠습니까? 종료시 제어기능이 동작하지 않습니다.", MsgBoxStyle.YesNo, "종료") = MsgBoxResult.Yes) Then
            '설정 저장
            Put_Local_Config()
            '서비스 종료
            For Each m As String In HOUSE_LINES
                HOUSES(m).Service_Stop()
            Next
        Else
            e.Cancel = True
        End If

    End Sub

    Private Sub Load_Data()

        '리스트 뷰에 데이터 Bind 연결
        For Each m As String In HOUSE_LINES
            Dim lv As ListView = CType(Me.FindName("LV_" & m), ListView)
            If lv IsNot Nothing Then
                lv.ItemsSource = HOUSES(m).GetList
            End If
        Next


    End Sub

    Public Sub TMR_UPDATE_TICK(sender As Object, e As EventArgs)
        Update_Network_Status()
    End Sub

    Private Sub Update_Network_Status()
        Dim hcode As String
        Dim OK_COLOR As Brush = New SolidColorBrush(Color.FromRgb(41, 128, 185))
        Dim ERR_COLOR As Brush = New SolidColorBrush(Color.FromRgb(192, 57, 43))
        Dim STOP_COLOR As Brush = New SolidColorBrush(Color.FromRgb(128, 128, 128)) 'gray

        '통신상태 표시 업데이트
        For Each s As String In HOUSE_LINES
            hcode = s

            Dim s_rect As Rectangle = CType(Me.FindName("rec_" & hcode & "_SOURCE_NETSTATUS"), Rectangle)
            Dim t_rect As Rectangle = CType(Me.FindName("rec_" & hcode & "_TARGET_NETSTATUS"), Rectangle)
            Dim id_rect As Rectangle = CType(Me.FindName("rec_" & hcode & "_SID_NETSTATUS"), Rectangle)

            Dim network_run As Boolean = False
            If (hcode.Substring(0, 2) = "H1") Then
                network_run = H1_NETWORK_RUN
            ElseIf (hcode.Substring(0, 2) = "H2") Then
                network_run = H2_NETWORK_RUN
            End If

            If (s_rect IsNot Nothing) Then
                If (network_run = True) Then
                    If (HOUSES(hcode).SOURCE_NET_STATUS = True) Then s_rect.Fill = OK_COLOR Else s_rect.Fill = ERR_COLOR
                Else
                    s_rect.Fill = STOP_COLOR
                End If
            End If
            If (t_rect IsNot Nothing) Then
                If (network_run = True) Then
                    If (HOUSES(hcode).TARGET_NET_STATUS = True) Then t_rect.Fill = OK_COLOR Else t_rect.Fill = ERR_COLOR
                Else
                    t_rect.Fill = STOP_COLOR
                End If
            End If
            If (id_rect IsNot Nothing) Then

                If (network_run = True) Then
                    '해당 설정이 수동이면 무조건 정상, 다른 모터 참조이면 해당 모터 상태 구해서 여부 표시
                    Dim id_status As Boolean = False
                    If (MCONFIG(hcode).SOURCE_MANUAL = True) Then '수동이면
                        id_status = True
                    Else '다른모터 참조이면
                        Dim sm As C_HOUSE.STRUCT_MOTOR = HOUSES(MCONFIG(hcode).SOURCE_LINE).Get_Motor(C_HOUSE.MOTOR_TYPE.OLD_MOTOR & "_" & MCONFIG(hcode).SOURCE_ID)
                        If (sm.STATUS = C_HOUSE.MOTOR_STATUS.OK) Then
                            id_status = True
                        Else
                            id_status = False
                        End If
                    End If
                    If (id_status = True) Then id_rect.Fill = OK_COLOR Else id_rect.Fill = ERR_COLOR
                Else
                    id_rect.Fill = STOP_COLOR
                End If

            End If


        Next
    End Sub


    Private Sub ListView_SizeChanged(sender As Object, e As SizeChangedEventArgs)
        Dim listview As ListView = CType(sender, ListView)
        Dim workingWidth As Integer = listview.ActualWidth - SystemParameters.VerticalScrollBarWidth - 15
        Dim gview As GridView = listview.View
        gview.Columns(0).Width = 35
        gview.Columns(1).Width = 60
        gview.Columns(2).Width = workingWidth - 35 - 60
    End Sub


    Private Sub lvHeader_CLICK(sender As Object, e As RoutedEventArgs)
        'MsgBox("header click" + e.Source.ToString + " " + sender.Name)
        '클릭한 리스트뷰이름
        Dim lv As ListView = CType(sender, ListView)
        Dim lv_name As String = lv.Name
        Dim header As GridViewColumnHeader = CType(e.OriginalSource, GridViewColumnHeader)
        If (header.Column Is Nothing) Then Exit Sub '빈 헤더 영역 클릭시 나가기
        '클릭한 열 글자
        Dim header_text As String = header.Column.Header.ToString

        'MsgBox(lv_name + " > " + header_text)

        Dim hname As String = lv_name.Replace("LV_", "")
        Debug.Print(hname & " " & header_text)
        Select Case header_text
            Case ""
                HOUSES(hname).Sort("MTYPE")
            Case "ID"
                HOUSES(hname).Sort("ID")
            'Case "RPM"
            '    HOUSES(hname).Sort("RPM")
            Case "상태"
                HOUSES(hname).Sort("STATUS")
        End Select


    End Sub

    Private Sub Lv_DBClick(sender As Object, e As MouseButtonEventArgs)

        Dim lv As ListView = CType(sender, ListView)
        Dim lv_name As String = lv.Name
        'MsgBox(lv_name)
        If (lv.SelectedItems.Count < 1) Then Exit Sub
        Dim sdata As C_LIST_BIND_Data = lv.SelectedItem
        'MsgBox(lv_name + " > " + sdata.ID)

        '수정창 띄우기
        Dim wm As New wMotor
        wm.Owner = Me
        wm.WMODE = "MOD"
        wm.MOD_DKEY = sdata.DIC_KEY
        wm.HOUSE_CODE = lv_name.Replace("LV_", "")
        wm.ShowDialog()

    End Sub

    Private Sub Btn_Setting_Click(sender As Object, e As RoutedEventArgs)
        Dim hcode As String = sender.name.Replace("Btn_Settings_", "")
        'MsgBox(sender.Name & " " & hcode)
        Dim sw As New wHOUSE
        sw.Owner = Me
        sw.HOUSE_CODE = hcode
        sw.ShowDialog()
    End Sub

    Private Sub Btn_Add_Click(sender As Object, e As RoutedEventArgs)
        'MsgBox(sender.name)
        Dim wm As New wMotor
        wm.Owner = Me
        wm.WMODE = "ADD"
        wm.HOUSE_CODE = sender.Name.Replace("Btn_Add_", "")
        wm.ShowDialog()
    End Sub


    Public Sub Update_HeaderBox()

        Dim hcode As String = ""

        For Each s As String In HOUSE_LINES
            hcode = s
            Dim old_ip As TextBlock = CType(Me.FindName(hcode & "_lbl_old_ip"), TextBlock)
            Dim new_ip As TextBlock = CType(Me.FindName(hcode & "_lbl_new_ip"), TextBlock)
            Dim src_info As TextBlock = CType(Me.FindName(hcode & "_lbl_source"), TextBlock)
            If (old_ip IsNot Nothing) Then
                old_ip.Text = MCONFIG(hcode).SOURCE_IP & "/" & MCONFIG(hcode).SOURCE_PORT
            End If
            If (new_ip IsNot Nothing) Then
                new_ip.Text = MCONFIG(hcode).TARGET_IP & "/" & MCONFIG(hcode).TARGET_PORT
            End If
            If (src_info IsNot Nothing) Then
                If MCONFIG(hcode).SOURCE_MANUAL = True Then
                    src_info.Text = "수동 > " & MCONFIG(hcode).MANUAL_RPM & " RPM"
                Else
                    src_info.Text = Get_House_ShortName(MCONFIG(hcode).SOURCE_LINE, " > ") & " > " & MCONFIG(hcode).SOURCE_ID
                End If
            End If
        Next


    End Sub

    Private Sub H1_StartBtn_Click(sender As Object, e As RoutedEventArgs)
        If (MsgBox("온실1 제어를 시작하시겠습니까?", MsgBoxStyle.YesNo, "제어") = MsgBoxResult.Yes) Then

            '서비스 시작
            For Each m As String In HOUSE_LINES
                If (m.Substring(0, 2) = "H1") Then HOUSES(m).Service_Start()
            Next

            H1_StartBtn.Visibility = Visibility.Collapsed
            H1_StopBtn.Visibility = Visibility.Visible
            H1_NETWORK_RUN = True
            Update_Network_Status()
        End If

    End Sub

    Private Sub H1_StopBtn_Click(sender As Object, e As RoutedEventArgs)
        If (MsgBox("온실1 제어를 종료하시겠습니까?", MsgBoxStyle.YesNo, "제어") = MsgBoxResult.Yes) Then

            '서비스 종료
            For Each m As String In HOUSE_LINES
                If (m.Substring(0, 2) = "H1") Then HOUSES(m).Service_Stop()
            Next


            H1_StartBtn.Visibility = Visibility.Visible
            H1_StopBtn.Visibility = Visibility.Collapsed
            H1_NETWORK_RUN = False
            Update_Network_Status()
        End If

    End Sub

    Private Sub H2_StartBtn_Click(sender As Object, e As RoutedEventArgs)
        Exit Sub
        If (MsgBox("온실2 제어를 시작하시겠습니까?", MsgBoxStyle.YesNo, "제어") = MsgBoxResult.Yes) Then

            '서비스 시작
            For Each m As String In HOUSE_LINES
                If (m.Substring(0, 2) = "H2") Then HOUSES(m).Service_Start()
            Next


            H2_StartBtn.Visibility = Visibility.Collapsed
            H2_StopBtn.Visibility = Visibility.Visible
            H2_NETWORK_RUN = True
            Update_Network_Status()
        End If

    End Sub

    Private Sub H2_StopBtn_Click(sender As Object, e As RoutedEventArgs)
        Exit Sub
        If (MsgBox("온실2 제어를 종료하시겠습니까?", MsgBoxStyle.YesNo, "제어") = MsgBoxResult.Yes) Then

            '서비스 종료
            For Each m As String In HOUSE_LINES
                If (m.Substring(0, 2) = "H2") Then HOUSES(m).Service_Stop()
            Next


            H2_StartBtn.Visibility = Visibility.Visible
            H2_StopBtn.Visibility = Visibility.Collapsed
            H2_NETWORK_RUN = False
            Update_Network_Status()
        End If
    End Sub
End Class
