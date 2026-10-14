Public Class wMotor

    Public WMODE As String '모드 ADD/MOD
    Public MOD_DKEY As String '모터 키
    Public HOUSE_CODE As String '라인코드 H1R1E

    Private MOTOR_ID As Integer '모터 ID
    Private MOTOR_TYPE As C_HOUSE.MOTOR_TYPE

    Public Sub New()

        ' 디자이너에서 이 호출이 필요합니다.
        InitializeComponent()

        ' InitializeComponent() 호출 뒤에 초기화 코드를 추가하세요.
        If (WMODE <> "ADD" And WMODE <> "MOD") Then WMODE = "ADD"
    End Sub

    Private Sub wMotor_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded
        If (WMODE = "ADD") Then
            BtnDelete.Visibility = Visibility.Hidden
            BtnSave.Content = "추가"
            wtitle.Text = "모터 추가"

            mod_type_line.Visibility = Visibility.Collapsed
            add_type_line.Visibility = Visibility.Visible

            mod_id_line.Visibility = Visibility.Collapsed
            add_id_line.Visibility = Visibility.Visible

            txtID.Text = ""

        ElseIf (WMODE = "MOD") Then
            BtnDelete.Visibility = Visibility.Visible
            BtnSave.Content = "수정"
            wtitle.Text = "모터 수정/삭제"

            add_type_line.Visibility = Visibility.Collapsed
            mod_type_line.Visibility = Visibility.Visible

            add_id_line.Visibility = Visibility.Collapsed
            mod_id_line.Visibility = Visibility.Visible

            '해당 모터 정보추출
            Dim m As C_HOUSE.STRUCT_MOTOR = HOUSES(HOUSE_CODE).Get_Motor(MOD_DKEY)

            If m.MTYPE = C_HOUSE.MOTOR_TYPE.OLD_MOTOR Then
                img_type_0.Visibility = Visibility.Visible
                img_type_1.Visibility = Visibility.Collapsed
                lbl_type.Content = "기존모터"
            Else
                img_type_0.Visibility = Visibility.Collapsed
                img_type_1.Visibility = Visibility.Visible
                lbl_type.Content = "신규모터"
            End If
            MOTOR_ID = m.ID
            MOTOR_TYPE = m.MTYPE
            '수정/삭제할 모터 아이디
            lblID.Content = m.ID

            If m.USE = True Then
                rb_use.IsChecked = True
                rb_nouse.IsChecked = False
            Else
                rb_use.IsChecked = False
                rb_nouse.IsChecked = True
            End If
        End If

        '라인명
        lbl_house.Content = Get_House_Name(HOUSE_CODE)

    End Sub

    Private Sub BtnSave_Click(sender As Object, e As RoutedEventArgs)
        Dim msg As String = "추가"
        If (WMODE = "MOD") Then msg = "수정"

        Dim new_id As Integer = 0
        Dim new_type As C_HOUSE.MOTOR_TYPE

        If (WMODE = "MOD") Then
            new_id = MOTOR_ID
            new_type = MOTOR_TYPE
        Else
            'type
            If (rb_oldmotor.IsChecked = False And rb_newmotor.IsChecked = False) Then
                MsgBox("모터 종류을 선택해 주세요.", MsgBoxStyle.OkOnly, "오류")
                Exit Sub
            End If
            If (rb_oldmotor.IsChecked = True) Then new_type = C_HOUSE.MOTOR_TYPE.OLD_MOTOR
            If (rb_newmotor.IsChecked = True) Then new_type = C_HOUSE.MOTOR_TYPE.NEW_MOTOR
            If (txtID.Text = "") Then
                MsgBox("모터 ID값을 숫자로 입력해주세요.", MsgBoxStyle.OkOnly, "오류")
                Exit Sub
            End If
            If (Check_Integer(txtID.Text) = False) Then
                MsgBox("모터 ID값을 숫자로 입력해주세요.", MsgBoxStyle.OkOnly, "오류")
                Exit Sub
            End If
            'id
            Integer.TryParse(txtID.Text, new_id)
            If (new_id < 1 Or new_id > 99) Then
                MsgBox("모터ID값은 숫자로 1-99범위로 입력하셔야 합니다.", MsgBoxStyle.OkOnly, "오류")
                Exit Sub
            End If
        End If

        '사용
        If (rb_use.IsChecked = False And rb_nouse.IsChecked = False) Then
            MsgBox("사용 여부을 선택해 주세요.", MsgBoxStyle.OkOnly, "오류")
            Exit Sub
        End If
        Dim new_use As Boolean = False
        If (rb_use.IsChecked = True) Then new_use = True

        'MsgBox(new_type & "/" & new_id & "/" & new_use)

        If MsgBox("해당 모터정보를 " & msg & "하시겠습니까?", MsgBoxStyle.YesNo, msg) = MsgBoxResult.Yes Then

            If WMODE = "ADD" Then '추가
                Dim dic_key As String = new_type & "_" & new_id
                Dim exist_test As C_HOUSE.STRUCT_MOTOR = HOUSES(HOUSE_CODE).Get_Motor(dic_key)
                If (exist_test.ID = new_id) Then
                    MsgBox("이미 해당 모터 종류와 ID를 갖는 정보가 있습니다. 확인바랍니다.", MsgBoxStyle.OkOnly, "오류")
                    Exit Sub
                End If
                HOUSES(HOUSE_CODE).Add_Motor(new_id, new_type, new_use)
                '변경 저장
                Put_Local_Config()
                Me.Close()
            ElseIf WMODE = "MOD" Then
                Dim dic_key As String = new_type & "_" & new_id
                HOUSES(HOUSE_CODE).Mod_Motor(dic_key, new_use)
                '변경 저장
                Put_Local_Config()
                Me.Close()
            End If
        End If


    End Sub

    Private Sub BtnDelete_Click(sender As Object, e As RoutedEventArgs)
        If MsgBox("해당 모터정보를 삭제하시겠습니까?", MsgBoxStyle.YesNo, "삭제") = MsgBoxResult.Yes Then

            '혹시 해당 모터를 참조로 사용하는지 확인
            If (MOTOR_TYPE = C_HOUSE.MOTOR_TYPE.OLD_MOTOR) Then '기존모터만 확인필요
                For Each m As String In HOUSE_LINES
                    If (MCONFIG(m).SOURCE_MANUAL = False And MCONFIG(m).SOURCE_LINE = HOUSE_CODE And MCONFIG(m).SOURCE_ID = MOTOR_ID) Then
                        MsgBox("해당 모터를 참조로 사용하는 라인이 있습니다. 변경 후, 삭제할 수 있습니다.", MsgBoxStyle.OkOnly, "오류")
                        Exit Sub
                    End If
                Next
            End If



            HOUSES(HOUSE_CODE).Remove_Motor(MOD_DKEY)
            '변경 저장
            Put_Local_Config()
            Me.Close()
        End If
    End Sub


    Private Sub PreviewKeyDown_PreventHanguel(sender As Object, e As KeyEventArgs)
        If (e.Key = Key.Space Or e.Key = Key.ImeProcessed) Then e.Handled = True
    End Sub

    Private Sub PreviewTextInput_OnlyInteger(sender As Object, e As TextCompositionEventArgs)
        Dim regex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^0-9]+") ' [^0-9.-]+ : 점과빼기추가
        e.Handled = regex.IsMatch(e.Text)
    End Sub

    Private Sub Border_MouseDown(sender As Object, e As MouseButtonEventArgs)
        If (e.ChangedButton = MouseButton.Left) Then
            Me.DragMove()
        End If
    End Sub
End Class
