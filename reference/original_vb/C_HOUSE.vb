
Imports System.Collections.ObjectModel
Imports System.Threading
Imports System.Net.Sockets
Public Class C_HOUSE

    Private DEBUG_LINE As String = "H1R1E" '디버그 정보 보고싶은 라인코드를 넣거나 ALL ( 모두 ) 입력
    Private DEBUG_IDS As New List(Of Integer)({5, 50}) '디버그 정보 보고싶은 ID들

    Private TMR_STATUS_UPDATER As New System.Windows.Threading.DispatcherTimer

    Private HOUSE_CODE As String

    Private MOTORS As New Dictionary(Of String, STRUCT_MOTOR) '모터들

    '화면 Bind 용 데이터 개체
    Private _LIST As ObservableCollection(Of C_LIST_BIND_Data) = New ObservableCollection(Of C_LIST_BIND_Data)()

    '정렬
    Private SORT_KEY As String
    Private SORT_DESC As Boolean

    '기존모터용 소켓, 송수신 쓰레드
    Private SOURCE_TCPSOCKET As Socket
    Private SOURCE_SEND_THREAD As Thread
    Private SOURCE_RECEIVE_THREAD As Thread
    Private SOURCE_CMD_IDX As Integer

    '신규모터용 소켓, 송수신 쓰레드
    Private TARGET_TCPSOCKET As Socket
    Private TARGET_SEND_THREAD As Thread
    Private TARGET_RECEIVE_THREAD As Thread
    Private TARGET_CMD_IDX As Integer

    '통신상태
    Public SOURCE_NET_STATUS As Boolean 'true 정상 / false 이상
    Public TARGET_NET_STATUS As Boolean 'true 정상 / false 이상

    'Thread 종료위해 사용하는 변수(Abort로 종료되지만 추가 작업)
    Private MUST_EXIT As Boolean = False

    Private Structure queue_type
        Dim DKEY As String '키
        Dim ID As Integer 'ID
        Dim MSG() As Byte '전송메시지
    End Structure
    Private SOURCE_READ_QUEUE As New Queue(Of queue_type) '기존모터 읽기 명령용 큐
    Private TARGET_READ_QUEUE As New Queue(Of queue_type) '신규모터 읽기 명령용 큐
    Private TARGET_WRITE_QUEUE As New Queue(Of queue_type) '신규모터 쓰기(제어) 명령용 큐

    Public Enum MOTOR_TYPE '모터 종류
        OLD_MOTOR = 0 '기존 모터(RPM읽어오는)
        NEW_MOTOR = 1 '신규 모터(RPM쓰는)
    End Enum

    Public Structure STRUCT_MOTOR
        Dim ID As Integer
        Dim MTYPE As MOTOR_TYPE
        Dim RPM As Integer
        Dim USE As Boolean 'true:사용/false:사용안함
        Dim RDate As DateTime '최근 정보 수신일시
        Dim STATUS As MOTOR_STATUS '상태 ( 0:정상, 1:이상, 2:중지(사용안함) )
        Dim LIST_IDX As Integer 'list의 인덱스번호(정렬된 리스트의 경우 이걸로 위치를 추적)
        Dim MUST_SEND_ZERO As Boolean '사용중지 선택시 한번은 0 보내기 위해 사용
    End Structure

    '모터상태
    Public Enum MOTOR_STATUS
        OK = 0 '정상
        ERR = 1 '이상
        NOTUSE = 2 '중지(사용안함)
        READY = 9 '대기(처음 추가시)
    End Enum

    'Bind용 모터리스트 추출
    Public Function GetList() As ObservableCollection(Of C_LIST_BIND_Data)
        Return _LIST
    End Function

    '등록되어 있는 이전모터 리스트ID 추출
    Public Function Get_OLDMOTOR_List() As List(Of String)
        Dim tmplist As List(Of String)
        Dim rlist As New List(Of String)
        tmplist = MOTORS.Keys.ToList()
        For Each k As String In tmplist
            If MOTORS(k).MTYPE = MOTOR_TYPE.OLD_MOTOR Then rlist.Add(MOTORS(k).ID) '기존모터만
        Next
        Return rlist
    End Function

    ' 해당 ID의 모터 정보 추출
    Public Function Get_Motor(Motor_DKEY As String) As C_HOUSE.STRUCT_MOTOR
        Dim rdata As STRUCT_MOTOR = Nothing
        If (MOTORS.ContainsKey(Motor_DKEY)) Then
            rdata = MOTORS(Motor_DKEY)
        End If
        Return rdata
    End Function

    Public Sub New(hcode As String)
        '코드
        HOUSE_CODE = hcode

        '모터들 상태 업데이트 타이머
        TMR_STATUS_UPDATER = New System.Windows.Threading.DispatcherTimer()
        TMR_STATUS_UPDATER.Interval = New TimeSpan(0, 0, 5) '5초
        AddHandler TMR_STATUS_UPDATER.Tick, AddressOf Me.TMR_STATUS_UPDATER_TICK
        '타이머 시작
        TMR_STATUS_UPDATER.Start()
    End Sub



    Public Sub TMR_STATUS_UPDATER_TICK(sender As Object, e As EventArgs)

        'Dim r As New Random
        ''현재 모터 개수
        'Dim c As Integer = MOTORS.Count
        'If (c < 1) Then Exit Sub
        'Dim idx As Integer = r.Next(0, c)
        ''값 랜덤값
        'Dim rpm As Integer = r.Next(1000, 9999) 'rpm
        ''랜덤값 적용
        'Dim dk As String = _LIST(idx).DIC_KEY
        'Dim t As STRUCT_MOTOR = MOTORS(dk)
        't.RPM = rpm
        't.RDate = Now
        'MOTORS(dk) = t
        ''바인드 리스트에 적용
        'Update_BIND_LIST(dk)


        '상태 갱신
        Dim klist As List(Of String) = MOTORS.Keys.ToList
        For Each k As String In klist
            Dim m As STRUCT_MOTOR = MOTORS(k)
            If m.USE = True Then
                If (m.RDate = Nothing) Then
                    m.STATUS = MOTOR_STATUS.READY
                Else
                    If (Math.Floor(Now.Subtract(m.RDate).TotalSeconds) >= 60) Then '최근 수신일자가 60초를 넘어가면 이상
                        m.STATUS = MOTOR_STATUS.ERR
                    Else
                        m.STATUS = MOTOR_STATUS.OK
                    End If
                End If
            Else
                m.STATUS = MOTOR_STATUS.NOTUSE
            End If
            '상태 변경되면 갱신
            If (m.STATUS <> _LIST(m.LIST_IDX).STATUS) Then
                _LIST(m.LIST_IDX).STATUS = m.STATUS
                MOTORS(k) = m
            End If
        Next

    End Sub
    '정렬
    Public Sub Sort(SKey As String)
        If SORT_KEY <> SKey Then '정렬 항목이 달라지면
            SORT_KEY = SKey
            SORT_DESC = False
        Else '정렬항목이 같으면
            SORT_DESC = Not (SORT_DESC) '오름/내림 반대로
        End If
        Update_BIND_LIST()
    End Sub

    Public Sub Add_Motor(id As Integer, mtype As MOTOR_TYPE, use As Boolean)
        Dim dic_key As String = mtype & "_" & id
        If (MOTORS.ContainsKey(dic_key) = False) Then
            Dim s As MOTOR_STATUS
            If (use = True) Then
                s = MOTOR_STATUS.READY '상태를 대기(9)로 시작
            Else
                s = MOTOR_STATUS.NOTUSE
            End If
            MOTORS.Add(dic_key, New STRUCT_MOTOR With {.ID = id, .MTYPE = mtype, .USE = use, .RPM = 0, .STATUS = s, .MUST_SEND_ZERO = False})
            _LIST.Add(DIC2BIND(dic_key, MOTORS(dic_key)))
            Update_BIND_LIST()
        End If
    End Sub

    Public Sub Remove_Motor(dkey As String)
        If (MOTORS.ContainsKey(dkey) = True) Then
            _LIST.RemoveAt(MOTORS(dkey).LIST_IDX)
            MOTORS.Remove(dkey)
            Update_BIND_LIST()
        End If
    End Sub

    Public Sub Mod_Motor(dkey As String, use As Boolean) '모터정보 사용여부 수정
        If (MOTORS.ContainsKey(dkey) = True) Then
            Dim tm As STRUCT_MOTOR = MOTORS(dkey)
            If (use <> tm.USE) Then '사용여부를 수정했을때
                If (use = True) Then
                    tm.STATUS = MOTOR_STATUS.READY
                    tm.MUST_SEND_ZERO = False
                Else
                    tm.STATUS = MOTOR_STATUS.NOTUSE
                    tm.RDate = Nothing
                    '사용 중지로 변경했을때는 must_send_zero를 사용해 한번은 0 보냄
                    tm.MUST_SEND_ZERO = True
                End If
                tm.USE = use
            End If
            MOTORS(dkey) = tm
            _LIST(tm.LIST_IDX) = DIC2BIND(dkey, tm)
        End If
    End Sub

    Public Sub Update_BIND_LIST(dkey As String) '모터정보를 bind 데이터에 업데이트
        If (MOTORS.ContainsKey(dkey) = True) Then
            _LIST(MOTORS(dkey).LIST_IDX) = DIC2BIND(dkey, MOTORS(dkey))
        End If
    End Sub

    Private Sub Update_BIND_LIST() '모터정보를 bind 데이터에 업데이트
        Dim ML As List(Of String)
        If SORT_KEY = "" Then
            SORT_KEY = "MTYPE"
            SORT_DESC = True
        End If
        '정렬
        Select Case SORT_KEY
            Case "MTYPE"
                'Debug_Print(0,"타입정렬")
                '정렬하기 위해 정렬기준값을 Value로 갖는 Dictionary 생성
                Dim TMP_DIC As New Dictionary(Of String, Integer)
                ML = MOTORS.Keys.ToList()
                For Each k As String In ML
                    Dim t As STRUCT_MOTOR = MOTORS(k)
                    If SORT_DESC = True Then
                        TMP_DIC.Add(k, t.MTYPE * 100 + 100 - t.ID) '타입으로 정렬시 같은 타입은 ID로 보이도록
                    Else
                        TMP_DIC.Add(k, t.MTYPE * 100 + t.ID) '타입으로 정렬시 같은 타입은 ID로 보이도록
                    End If

                Next
                '오름/내림 기준에 따라 정렬
                Dim SORT_DIC As Dictionary(Of String, Integer)
                If (SORT_DESC = True) Then
                    SORT_DIC = (From entry In TMP_DIC Order By entry.Value Descending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
                Else
                    SORT_DIC = (From entry In TMP_DIC Order By entry.Value Ascending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
                End If
                ML = SORT_DIC.Keys.ToList
            Case "ID"
                'Debug_Print(0,"ID정렬")
                '정렬하기 위해 정렬기준값을 Value로 갖는 Dictionary 생성
                Dim TMP_DIC As New Dictionary(Of String, Integer)
                ML = MOTORS.Keys.ToList()
                For Each k As String In ML
                    Dim t As STRUCT_MOTOR = MOTORS(k)
                    TMP_DIC.Add(k, t.ID)
                Next
                '오름/내림 기준에 따라 정렬
                Dim SORT_DIC As New Dictionary(Of String, Integer)
                If (SORT_DESC = True) Then
                    SORT_DIC = (From entry In TMP_DIC Order By entry.Value Descending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
                Else
                    SORT_DIC = (From entry In TMP_DIC Order By entry.Value Ascending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
                End If
                ML = SORT_DIC.Keys.ToList
            'Case "RPM"
            '    'Debug_Print(0,"RPM정렬")
            '    '정렬하기 위해 정렬기준값을 Value로 갖는 Dictionary 생성
            '    Dim TMP_DIC As New Dictionary(Of String, Integer)
            '    ML = MOTORS.Keys.ToList()
            '    For Each k As String In ML
            '        Dim t As STRUCT_MOTOR = MOTORS(k)
            '        TMP_DIC.Add(k, t.RPM)
            '    Next
            '    '오름/내림 기준에 따라 정렬
            '    Dim SORT_DIC As Dictionary(Of String, Integer)
            '    If (SORT_DESC = True) Then
            '        SORT_DIC = (From entry In TMP_DIC Order By entry.Value Descending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
            '    Else
            '        SORT_DIC = (From entry In TMP_DIC Order By entry.Value Ascending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
            '    End If
            '    ML = SORT_DIC.Keys.ToList
            Case "STATUS"
                'Debug_Print(0,"ONOFF정렬")
                '정렬하기 위해 정렬기준값을 Value로 갖는 Dictionary 생성
                Dim TMP_DIC As New Dictionary(Of String, Integer)
                ML = MOTORS.Keys.ToList()
                For Each k As String In ML
                    Dim t As STRUCT_MOTOR = MOTORS(k)
                    TMP_DIC.Add(k, t.STATUS)
                Next
                '오름/내림 기준에 따라 정렬
                Dim SORT_DIC As Dictionary(Of String, Integer)
                If (SORT_DESC = True) Then
                    SORT_DIC = (From entry In TMP_DIC Order By entry.Value Descending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
                Else
                    SORT_DIC = (From entry In TMP_DIC Order By entry.Value Ascending Select entry).ToDictionary(Function(p) p.Key, Function(p) p.Value)
                End If
                ML = SORT_DIC.Keys.ToList
            Case Else
                'Debug_Print(0,"기본정렬")
                ML = MOTORS.Keys.ToList()
        End Select

        '리스트 갱신(리스트를 새로만들어 덮어씌우면  ListView 갱신안됨. 그래서 기존 _LIST를 수정함 )
        Dim idx As Integer = 0
        For Each k As String In ML
            _LIST(idx) = DIC2BIND(k, MOTORS(k))
            Dim tm As STRUCT_MOTOR = MOTORS(k)
            tm.LIST_IDX = idx
            MOTORS(k) = tm
            idx += 1
        Next


    End Sub

    Private Function DIC2BIND(dk As String, d As STRUCT_MOTOR) As C_LIST_BIND_Data
        Dim dic_key As String = dk
        Dim id As Integer = d.ID
        Dim mtype As MOTOR_TYPE = d.MTYPE
        Dim rpm As Integer = d.RPM
        Dim use As Boolean = d.USE
        Dim status As MOTOR_STATUS = d.STATUS
        Dim rdate As DateTime = d.RDate
        Return New C_LIST_BIND_Data() With {.DIC_KEY = dic_key, .ID = id, .RPM = rpm, .USE = use, .MType = mtype, .STATUS = status, .RDATE = rdate}
    End Function

#Region "송수신"
    '송수신시작
    Public Sub Service_Start()

        DEBUG_PRINT(0, "[" & HOUSE_CODE & "] Start Service!")

        SOURCE_CMD_IDX = 0
        TARGET_CMD_IDX = 0

        '메시지 큐 삭제
        Try
            SOURCE_READ_QUEUE.Clear()
            TARGET_READ_QUEUE.Clear()
            TARGET_WRITE_QUEUE.Clear()
        Catch ex As Exception
        End Try

        '종료변수 false 
        MUST_EXIT = False

        '상태 시작값
        SOURCE_NET_STATUS = False '이상으로 시작
        TARGET_NET_STATUS = False '이상으로 시작

        '원본 수신 쓰레드 ( 송신쓰레드가 시작 후, 시작됨 )
        SOURCE_RECEIVE_THREAD = New Thread(AddressOf src_recv_job)
        SOURCE_RECEIVE_THREAD.IsBackground = True
        '신규 수신 쓰레드 ( 송신쓰레드가 시작 후, 시작됨 )
        TARGET_RECEIVE_THREAD = New Thread(AddressOf target_recv_job)
        TARGET_RECEIVE_THREAD.IsBackground = True

        '원본 송신 쓰레드(맨 마지막에 있어야 함)
        SOURCE_SEND_THREAD = New Thread(AddressOf src_send_job)
        SOURCE_SEND_THREAD.IsBackground = True
        SOURCE_SEND_THREAD.Start()
        '신규 송신 쓰레드 (맨 마지막에 있어야 함)
        TARGET_SEND_THREAD = New Thread(AddressOf target_send_job)
        TARGET_SEND_THREAD.IsBackground = True
        TARGET_SEND_THREAD.Start()
    End Sub

    Public Sub Service_Stop()

        '종료변수 true
        MUST_EXIT = True

        '메시지 큐 삭제
        Try
            SOURCE_READ_QUEUE.Clear()
            TARGET_READ_QUEUE.Clear()
            TARGET_WRITE_QUEUE.Clear()
        Catch ex As Exception
        End Try

        '수신 쓰레드 종료
        Try
            If (SOURCE_RECEIVE_THREAD IsNot Nothing) Then SOURCE_RECEIVE_THREAD.Abort()
        Catch ex As Exception
        End Try
        '송신 스레드 종료
        Try
            If (SOURCE_SEND_THREAD IsNot Nothing) Then SOURCE_SEND_THREAD.Abort()
        Catch ex As Exception
        End Try
        '소켓 종료
        Try
            If (SOURCE_TCPSOCKET IsNot Nothing) Then
                If (SOURCE_TCPSOCKET.Connected = True) Then SOURCE_TCPSOCKET.Close()
                SOURCE_TCPSOCKET = Nothing
            End If
        Catch ex As Exception
        End Try

        '수신 쓰레드 종료
        Try
            If (TARGET_RECEIVE_THREAD IsNot Nothing) Then TARGET_RECEIVE_THREAD.Abort()
        Catch ex As Exception
        End Try
        '송신 스레드 종료
        Try
            If (TARGET_SEND_THREAD IsNot Nothing) Then TARGET_SEND_THREAD.Abort()
        Catch ex As Exception
        End Try
        '소켓 종료
        Try
            If (TARGET_TCPSOCKET IsNot Nothing) Then
                If (TARGET_TCPSOCKET.Connected = True) Then TARGET_TCPSOCKET.Close()
                TARGET_TCPSOCKET = Nothing
            End If
        Catch ex As Exception
        End Try

    End Sub

    Public Sub Service_Restart()
        Service_Stop()
        Thread.Sleep(1000)
        Service_Start()
    End Sub
    Public Sub src_send_job()

        DEBUG_PRINT(0, "[" & HOUSE_CODE & "][S] 전송 시작")

        While (1)

            If (MUST_EXIT = True) Then Exit While

            'STEP1. 소켓이 엾으면 생성
            If (SOURCE_TCPSOCKET Is Nothing) Then
                SOURCE_TCPSOCKET = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            End If

            'STEP2. 미접속시 접속
            If (SOURCE_TCPSOCKET.Connected = False) Then '미접속 상태이면

                DEBUG_PRINT(0, "[" & HOUSE_CODE & "][S] 접속시도")

                If (CHECK_CONNECT_POSSIBLE(MCONFIG(HOUSE_CODE).SOURCE_IP, MCONFIG(HOUSE_CODE).SOURCE_PORT, 1) = True) Then '1초 내에 해당 포트로 접속가능한지 체크
                    Try
                        '접속
                        SOURCE_TCPSOCKET = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                        SOURCE_TCPSOCKET.Connect(MCONFIG(HOUSE_CODE).SOURCE_IP, MCONFIG(HOUSE_CODE).SOURCE_PORT)
                    Catch ex As Exception
                        '접속실패
                        Try
                            If (SOURCE_TCPSOCKET IsNot Nothing) Then SOURCE_TCPSOCKET.Close()
                        Catch ex2 As Exception
                        End Try
                        'On_Update_logmsg("SOURCE", "접속실패!")
                        DEBUG_PRINT(0, "[" & HOUSE_CODE & "][S] 접속실패!")
                        SOURCE_NET_STATUS = False
                        '접속실패로 잠시후 다시 시도
                        Thread.Sleep(10000) '10초
                        Continue While
                    End Try
                Else '접속안되는 정보이면
                    '접속안됨.
                    Try
                        If SOURCE_TCPSOCKET IsNot Nothing Then SOURCE_TCPSOCKET.Close()
                    Catch ex As Exception
                    End Try
                    'On_Update_logmsg("SOURCE", "접속실패!")
                    DEBUG_PRINT(0, "[" & HOUSE_CODE & "][S] 접속실패!!")
                    SOURCE_NET_STATUS = False
                    '접속실패로 잠시후 다시 시도
                    Thread.Sleep(10000) '10초
                    Continue While
                End If

            End If

            '통신상태
            SOURCE_NET_STATUS = True

            'STEP3. 수신 쓰레드가 없으면 시작하기
            Try
                If SOURCE_RECEIVE_THREAD Is Nothing Or SOURCE_RECEIVE_THREAD.IsAlive = False Then
                    SOURCE_RECEIVE_THREAD = New Thread(AddressOf src_recv_job)
                    SOURCE_RECEIVE_THREAD.IsBackground = True
                    SOURCE_RECEIVE_THREAD.Start()
                End If
            Catch ex As Exception
            End Try

            'STEP4. 큐에 있으면 전송, 없으면 기존모터 요청 메시지 큐에 저장
            Try

                Dim logmsg As String = ""
                '큐에 값이 있으면 바로 전송
                If (SOURCE_READ_QUEUE.Count > 0) Then
                    For qidx As Integer = 0 To SOURCE_READ_QUEUE.Count - 1

                        If (MUST_EXIT = True) Then Exit While

                        Dim qmsg As queue_type
                        qmsg = SOURCE_READ_QUEUE.Dequeue
                        If (MOTORS.ContainsKey(qmsg.DKEY)) Then
                            Dim control_byte() As Byte
                            control_byte = qmsg.MSG
                            SOURCE_TCPSOCKET.Send(control_byte)

                            logmsg = ""
                            For i As Integer = 0 To control_byte.Length - 1
                                If logmsg <> "" Then logmsg &= " "
                                logmsg &= control_byte(i).ToString("X2")
                            Next
                            DEBUG_PRINT(qmsg.ID, "[" & HOUSE_CODE & "][S][" & qmsg.ID & "] " & logmsg)

                            Thread.Sleep(500)
                        End If
                    Next
                Else '전송큐에 보낼 데이터가 없으면 모터수만큼 전송 큐에 넣기
                    Dim mlist As List(Of String) = MOTORS.Keys.ToList
                    For Each m As String In mlist

                        If (MUST_EXIT = True) Then Exit While

                        If (MOTORS(m).MTYPE = MOTOR_TYPE.OLD_MOTOR And MOTORS(m).USE = True) Then '기존모터중 사용중인것만
                            Dim control_byte() As Byte
                            control_byte = GET_MSG_SOURCE_RPM_READ(MOTORS(m).ID)

                            Dim qmsg As queue_type
                            qmsg.DKEY = m
                            qmsg.ID = MOTORS(m).ID
                            qmsg.MSG = control_byte
                            SOURCE_READ_QUEUE.Enqueue(qmsg)

                            'logmsg = ""
                            'For i As Integer = 0 To control_byte.Length - 1
                            '    If logmsg <> "" Then logmsg &= " "
                            '    logmsg &= control_byte(i).ToString("X2")
                            'Next
                            'Debug_Print(motors(m).id, "[" & HOUSE_CODE & "][" & MOTORS(m).ID & "] " & logmsg)

                        End If

                    Next

                    Thread.Sleep(1000)

                End If
            Catch ex As Exception
                If SOURCE_TCPSOCKET IsNot Nothing Then SOURCE_TCPSOCKET.Close()
                Thread.Sleep(1000)
            End Try

        End While

    End Sub
    Public Sub src_recv_job()

        Try
            While (1)

                If (MUST_EXIT = True) Then Exit While

                Dim buf(512) As Byte
                Dim rlen As Integer = 0
                rlen = SOURCE_TCPSOCKET.Receive(buf)
                If rlen < 1 Then
                    Exit Sub
                End If

                Dim txtmsg As String = ""
                Dim hexmsg As String = ""

                txtmsg = System.Text.Encoding.UTF8.GetString(buf, 0, rlen)

                For i As Integer = 0 To rlen - 1
                    hexmsg &= buf(i).ToString("X02") & " "
                Next


                If (rlen = 11 And buf(7) = &H4) Then 'FN4번 길이 11자리 ( 읽기에 대한 응답 / buf(6)에 ID값이 있음)
                    Dim r_id As Integer = buf(6)

                    Dim RCOMMAND_IDX As Integer = buf(2) * 256 + buf(3) '혹시 ID값으로 수신값을 구분못하면 명령어 번호로 구분하기 위해
                    Dim NOW_RPM As Integer = buf(9) * 256 + buf(10)

                    DEBUG_PRINT(r_id, "[" & HOUSE_CODE & "][S] " & r_id & "=" & NOW_RPM)
                    Try
                        Dim DKEY As String = MOTOR_TYPE.OLD_MOTOR & "_" & r_id
                        If (MOTORS.ContainsKey(DKEY) = True) Then
                            '모터 정보 변경
                            Dim t As STRUCT_MOTOR = MOTORS(DKEY)
                            t.RPM = NOW_RPM
                            t.RDate = Now
                            MOTORS(DKEY) = t
                            'bind 용 list변경
                            _LIST(t.LIST_IDX).RPM = NOW_RPM
                            _LIST(t.LIST_IDX).RDATE = Now
                        End If
                    Catch ex As Exception
                    End Try
                End If

            End While

        Catch ex As Exception

        End Try
    End Sub
    Public Sub target_send_job()

        DEBUG_PRINT(0, "[" & HOUSE_CODE & "][T] 전송 시작")

        While (1)

            If (MUST_EXIT = True) Then Exit While

            'STEP1. 소켓이 엾으면 생성
            If (TARGET_TCPSOCKET Is Nothing) Then
                TARGET_TCPSOCKET = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            End If

            'STEP2. 미접속시 접속
            If (TARGET_TCPSOCKET.Connected = False) Then '미접속 상태이면

                DEBUG_PRINT(0, "[" & HOUSE_CODE & "][T] 접속시도")

                If (CHECK_CONNECT_POSSIBLE(MCONFIG(HOUSE_CODE).TARGET_IP, MCONFIG(HOUSE_CODE).TARGET_PORT, 1) = True) Then '1초 내에 해당 포트로 접속가능한지 체크
                    Try
                        '접속
                        TARGET_TCPSOCKET = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
                        TARGET_TCPSOCKET.Connect(MCONFIG(HOUSE_CODE).TARGET_IP, MCONFIG(HOUSE_CODE).TARGET_PORT)
                    Catch ex As Exception
                        '접속실패
                        Try
                            If TARGET_TCPSOCKET IsNot Nothing Then TARGET_TCPSOCKET.Close()
                        Catch ex2 As Exception
                        End Try
                        'On_Update_logmsg("SOURCE", "접속실패!")
                        DEBUG_PRINT(0, "[" & HOUSE_CODE & "][T] 접속실패!")
                        '통신상태
                        TARGET_NET_STATUS = False
                        '접속실패로 잠시후 다시 시도
                        Thread.Sleep(10000) '10초
                        Continue While
                    End Try
                Else '접속안되는 정보이면
                    '접속안됨.
                    Try
                        If TARGET_TCPSOCKET IsNot Nothing Then TARGET_TCPSOCKET.Close()
                    Catch ex As Exception
                    End Try
                    'On_Update_logmsg("SOURCE", "접속실패!")
                    DEBUG_PRINT(0, "[" & HOUSE_CODE & "][T] 접속실패!!")
                    '통신상태
                    TARGET_NET_STATUS = False
                    '접속실패로 잠시후 다시 시도
                    Thread.Sleep(10000) '10초
                    Continue While
                End If

            End If

            '통신상태
            TARGET_NET_STATUS = True

            'STEP3. 수신 쓰레드가 없으면 시작하기
            Try
                If TARGET_RECEIVE_THREAD Is Nothing Or TARGET_RECEIVE_THREAD.IsAlive = False Then
                    TARGET_RECEIVE_THREAD = New Thread(AddressOf target_recv_job)
                    TARGET_RECEIVE_THREAD.IsBackground = True
                    TARGET_RECEIVE_THREAD.Start()
                End If
            Catch ex As Exception
            End Try


            'STEP4. 큐에 있으면 전송, 없으면 기존모터 요청 메시지 큐에 저장
            Try

                Dim logmsg As String = ""
                '읽기큐에 값이 있으면 바로 전송
                If (TARGET_READ_QUEUE.Count > 0) Then
                    For qidx As Integer = 0 To TARGET_READ_QUEUE.Count - 1

                        If (MUST_EXIT = True) Then Exit While

                        Dim qmsg As queue_type
                        '쓰기 명령어 큐가 있으면
                        If (TARGET_WRITE_QUEUE.Count > 0) Then

                            For widx As Integer = 0 To TARGET_WRITE_QUEUE.Count - 1
                                '종료
                                If (MUST_EXIT = True) Then Exit While

                                '큐에서 뽑아 전송
                                qmsg = TARGET_WRITE_QUEUE.Dequeue
                                If (MOTORS.ContainsKey(qmsg.DKEY)) Then
                                    Dim control_byte() As Byte
                                    control_byte = qmsg.MSG
                                    TARGET_TCPSOCKET.Send(control_byte)
                                    '로그
                                    logmsg = ""
                                    For i As Integer = 0 To control_byte.Length - 1
                                        If logmsg <> "" Then logmsg &= " "
                                        logmsg &= control_byte(i).ToString("X2")
                                    Next
                                    DEBUG_PRINT(qmsg.ID, "[" & HOUSE_CODE & "][W][" & qmsg.ID & "] " & logmsg)
                                    '쉬기
                                    Thread.Sleep(500)
                                End If

                            Next

                        End If

                        '읽기 명령어 전송
                        qmsg = TARGET_READ_QUEUE.Dequeue
                        If (MOTORS.ContainsKey(qmsg.DKEY)) Then
                            Dim control_byte() As Byte
                            control_byte = qmsg.MSG
                            TARGET_TCPSOCKET.Send(control_byte)

                            logmsg = ""
                            For i As Integer = 0 To control_byte.Length - 1
                                If logmsg <> "" Then logmsg &= " "
                                logmsg &= control_byte(i).ToString("X2")
                            Next
                            DEBUG_PRINT(qmsg.ID, "[" & HOUSE_CODE & "][T][" & qmsg.ID & "] " & logmsg)

                            Thread.Sleep(500)
                        End If
                    Next
                Else '전송큐에 보낼 데이터가 없으면 모터수만큼 전송 큐에 넣기
                    Dim mlist As List(Of String) = MOTORS.Keys.ToList
                    For Each m As String In mlist

                        If (MUST_EXIT = True) Then Exit While

                        If (MOTORS(m).MTYPE = MOTOR_TYPE.NEW_MOTOR And MOTORS(m).USE = True) Then '신규모터중 사용중인것만
                            Dim control_byte() As Byte
                            control_byte = GET_MSG_TAGET_READ(MOTORS(m).ID)

                            Dim qmsg As queue_type
                            qmsg.DKEY = m
                            qmsg.ID = MOTORS(m).ID
                            qmsg.MSG = control_byte
                            TARGET_READ_QUEUE.Enqueue(qmsg)

                            'logmsg = ""
                            'For i As Integer = 0 To control_byte.Length - 1
                            '    If logmsg <> "" Then logmsg &= " "
                            '    logmsg &= control_byte(i).ToString("X2")
                            'Next
                            'Debug_Print(motors(m).id, "[" & HOUSE_CODE & "][" & MOTORS(m).ID & "] " & logmsg)

                        ElseIf (MOTORS(m).MTYPE = MOTOR_TYPE.NEW_MOTOR And MOTORS(m).USE = False And MOTORS(m).MUST_SEND_ZERO = True) Then
                            '신규 모터중 사용중지 되어 있고, 0 보내기 설정되어 있으면 rpm 0 보낵
                            Dim tm As STRUCT_MOTOR = MOTORS(m)
                            '0 보내기 비활성화
                            tm.MUST_SEND_ZERO = False
                            MOTORS(m) = tm

                            Dim qmsg As queue_type
                            qmsg.DKEY = MOTOR_TYPE.NEW_MOTOR & "_" & tm.ID
                            qmsg.ID = tm.ID
                            qmsg.MSG = GET_MSG_TAGET_RPM_WRITE(tm.ID, 0) 'rpm 0
                            TARGET_WRITE_QUEUE.Enqueue(qmsg)

                        End If
                    Next

                    Thread.Sleep(1000)

                End If
            Catch ex As Exception
                If TARGET_TCPSOCKET IsNot Nothing Then TARGET_TCPSOCKET.Close()
                Thread.Sleep(1000)
            End Try

        End While

    End Sub
    Public Sub target_recv_job()

        Try
            While (1)

                If (MUST_EXIT = True) Then Exit While

                Dim buf(512) As Byte
                Dim rlen As Integer = 0
                rlen = TARGET_TCPSOCKET.Receive(buf)
                If rlen < 1 Then
                    Exit Sub
                End If

                'Dim txtmsg As String = ""
                'Dim hexmsg As String = ""
                'txtmsg = System.Text.Encoding.UTF8.GetString(buf, 0, rlen)
                'For i As Integer = 0 To rlen - 1
                '    hexmsg &= buf(i).ToString("X02") & " "
                'Next

                If (rlen = 23 And buf(7) = &H3 And buf(8) = &HE) Then '해당 장비번호의 읽기에 대한 응답이면(길이17, fn &H3, 개수 &HE, 읽기는 수신시 어드레어가 없음 )
                    Dim ID As Integer = buf(6)
                    Dim RCOMMAND_IDX As Integer = buf(2) * 256 + buf(3)
                    Dim NOW_AUTOMANUAL As Integer = buf(9) * 256 + buf(10)
                    Dim NOW_STARTSTOP As Integer = buf(11) * 256 + buf(12)
                    Dim NOW_RPM As Integer = buf(15) * 256 + buf(16)

                    Try
                        Dim DKEY As String = MOTOR_TYPE.NEW_MOTOR & "_" & ID
                        If (MOTORS.ContainsKey(DKEY) = True) Then

                            Dim SETTING_RPM As Integer
                            If (MCONFIG(HOUSE_CODE).SOURCE_MANUAL = True) Then
                                '수동이면 해당RPM 사용
                                SETTING_RPM = MCONFIG(HOUSE_CODE).MANUAL_RPM
                                DEBUG_PRINT(0, "[" & HOUSE_CODE & "][T]-참조 수동:" & SETTING_RPM)
                            Else
                                '해당 모터의 RPM 추출
                                Dim sm As STRUCT_MOTOR = HOUSES(MCONFIG(HOUSE_CODE).SOURCE_LINE).Get_Motor(MOTOR_TYPE.OLD_MOTOR & "_" & MCONFIG(HOUSE_CODE).SOURCE_ID)
                                DEBUG_PRINT(0, "[" & HOUSE_CODE & "][T]-참조 :" & MCONFIG(HOUSE_CODE).SOURCE_LINE & "/" & MCONFIG(HOUSE_CODE).SOURCE_ID & "/" & sm.RPM)
                                If sm.STATUS = MOTOR_STATUS.OK Then '상태가 정상일때
                                    SETTING_RPM = sm.RPM
                                Else
                                    '참조 모터가 정상이 아니면 넘어감. ( 제어하면 안될까? )
                                    DEBUG_PRINT(0, "참조모터 상태이상")
                                    Continue While
                                End If
                            End If

                            If (NOW_AUTOMANUAL <> 1) Then '수동이 아니면 수동세팅 명령어 보내기
                                '제어큐에 넣기
                                Dim qmsg As queue_type
                                qmsg.DKEY = DKEY
                                qmsg.ID = ID
                                qmsg.MSG = GET_MSG_TAGET_MANUAL_WRITE(ID)
                                TARGET_WRITE_QUEUE.Enqueue(qmsg)
                            End If
                            If NOW_STARTSTOP <> 1 Then '동작이 아니면 동작세팅 명령어 보내기
                                '제어큐에 넣기
                                Dim qmsg As queue_type
                                qmsg.DKEY = DKEY
                                qmsg.ID = ID
                                qmsg.MSG = GET_MSG_TAGET_RUN_WRITE(ID)
                                TARGET_WRITE_QUEUE.Enqueue(qmsg)
                            End If
                            DEBUG_PRINT(ID, "[" & HOUSE_CODE & "][" & ID & "]-제어RPM:" & SETTING_RPM)
                            If (NOW_RPM <> SETTING_RPM) Then
                                '제어큐에 넣기
                                Dim qmsg As queue_type
                                qmsg.DKEY = DKEY
                                qmsg.ID = ID
                                qmsg.MSG = GET_MSG_TAGET_RPM_WRITE(ID, SETTING_RPM)
                                TARGET_WRITE_QUEUE.Enqueue(qmsg)
                            End If

                        End If
                    Catch ex As Exception
                    End Try
                End If

            End While

        Catch ex As Exception

        End Try

    End Sub
#End Region

#Region "modbus_function"
    Private Function GET_MSG_SOURCE_RPM_READ(id As Integer) As Byte()
        Dim rbyte(11) As Byte

        '명령 번호, 이 번호로 응답이 옴.
        Dim t As String = SOURCE_CMD_IDX.ToString("X4")
        rbyte(0) = "&H" & t.Substring(0, 2)
        rbyte(1) = "&H" & t.Substring(2, 2)
        rbyte(2) = &H0
        rbyte(3) = &H0
        rbyte(4) = &H0
        rbyte(5) = &H6 '이값 다음부터 끝까지 개수
        rbyte(6) = "&H" & Hex(id) 'ID
        rbyte(7) = &H4 'FN read input register
        rbyte(8) = &H0 'addr
        rbyte(9) = "&H" & Hex(14) 'addr
        rbyte(10) = &H0 '개수
        rbyte(11) = &H1 '개수

        SOURCE_CMD_IDX += 1
        If (SOURCE_CMD_IDX > 9999) Then SOURCE_CMD_IDX = 0

        Return rbyte
    End Function

    Private Function GET_MSG_TAGET_READ(id As Integer) As Byte()

        ' tcp protocol
        Dim rbyte(11) As Byte
        '명령 번호, 이 번호로 응답이 옴.
        Dim t As String = TARGET_CMD_IDX.ToString("X4")
        rbyte(0) = "&H" & t.Substring(0, 2)
        rbyte(1) = "&H" & t.Substring(2, 2)
        rbyte(2) = &H0
        rbyte(3) = &H0
        rbyte(4) = &H0
        rbyte(5) = &H6 '이값 다음부터 끝까지 개수
        rbyte(6) = "&H" & Hex(id) 'UID
        rbyte(7) = &H3 'FN read single register
        rbyte(8) = &H0 'addr
        rbyte(9) = &H13 'addr ( 20번 )
        rbyte(10) = &H0 '개수
        rbyte(11) = &H7 '개수 ( 무조건 7개를 읽어야 함 )

        TARGET_CMD_IDX += 1
        If (TARGET_CMD_IDX > 9999) Then TARGET_CMD_IDX = 0

        Return rbyte
    End Function

    Private Function GET_MSG_TAGET_WRITE(id As Integer, automanual As Integer, startstop As Integer, rpm As Integer) As Byte()
        'tcp protocol

        Dim rbyte(20) As Byte
        '명령 번호, 이 번호로 응답이 옴.
        Dim t As String = TARGET_CMD_IDX.ToString("X4")
        rbyte(0) = "&H" & t.Substring(0, 2)
        rbyte(1) = "&H" & t.Substring(2, 2)
        rbyte(2) = &H0
        rbyte(3) = &H0

        rbyte(4) = &H0
        rbyte(5) = &HF '이값 다음부터 끝까지 개수
        rbyte(6) = "&H" & Hex(id) 'UID
        rbyte(7) = &H10 'FN write muliple register
        rbyte(8) = &H0 'addr
        rbyte(9) = &H13 'addr
        rbyte(10) = &H0 '개수
        rbyte(11) = &H4 '개수
        rbyte(12) = &H2 '데이터당 바이트수
        rbyte(13) = &H0 '자동/수동
        rbyte(14) = "&H" & automanual '자동/수동
        rbyte(15) = &H0 '동작/정지
        rbyte(16) = "&H" & startstop '동작/정지
        rbyte(17) = &H0
        rbyte(18) = &H0
        'rpm
        t = rpm.ToString("X4")
        rbyte(19) = "&H" & t.Substring(0, 2)
        rbyte(20) = "&H" & t.Substring(2, 2)

        TARGET_CMD_IDX += 1
        If (TARGET_CMD_IDX > 9999) Then TARGET_CMD_IDX = 0

        Return rbyte
    End Function

    Private Function GET_MSG_TAGET_RPM_WRITE(id As Integer, rpm As Integer) As Byte()
        Dim rbyte(11) As Byte
        '명령 번호, 이 번호로 응답이 옴.
        Dim t As String = TARGET_CMD_IDX.ToString("X4")
        rbyte(0) = "&H" & t.Substring(0, 2)
        rbyte(1) = "&H" & t.Substring(2, 2)
        rbyte(2) = &H0
        rbyte(3) = &H0

        rbyte(4) = &H0
        rbyte(5) = &H6 '이값 다음부터 끝까지 개수
        rbyte(6) = "&H" & Hex(id) 'UID
        rbyte(7) = &H6 'FN write single register
        rbyte(8) = &H0 'addr
        rbyte(9) = &H16 'addr
        'rpm
        t = rpm.ToString("X4")
        rbyte(10) = "&H" & t.Substring(0, 2)
        rbyte(11) = "&H" & t.Substring(2, 2)

        TARGET_CMD_IDX += 1
        If (TARGET_CMD_IDX > 9999) Then TARGET_CMD_IDX = 0

        Return rbyte
    End Function

    Private Function GET_MSG_TAGET_MANUAL_WRITE(id As Integer) As Byte()
        Dim rbyte(11) As Byte
        '명령 번호, 이 번호로 응답이 옴.
        Dim t As String = TARGET_CMD_IDX.ToString("X4")
        rbyte(0) = "&H" & t.Substring(0, 2)
        rbyte(1) = "&H" & t.Substring(2, 2)
        rbyte(2) = &H0
        rbyte(3) = &H0

        rbyte(4) = &H0
        rbyte(5) = &H6 '이값 다음부터 끝까지 개수
        rbyte(6) = "&H" & Hex(id) 'UID
        rbyte(7) = &H6 'FN write single register
        rbyte(8) = &H0 'addr
        rbyte(9) = &H13 'addr
        rbyte(10) = &H0
        rbyte(11) = &H1 'manual


        TARGET_CMD_IDX += 1
        If (TARGET_CMD_IDX > 9999) Then TARGET_CMD_IDX = 0

        Return rbyte
    End Function

    Private Function GET_MSG_TAGET_RUN_WRITE(id As Integer) As Byte()
        Dim rbyte(11) As Byte
        '명령 번호, 이 번호로 응답이 옴.
        Dim t As String = TARGET_CMD_IDX.ToString("X4")
        rbyte(0) = "&H" & t.Substring(0, 2)
        rbyte(1) = "&H" & t.Substring(2, 2)
        rbyte(2) = &H0
        rbyte(3) = &H0

        rbyte(4) = &H0
        rbyte(5) = &H6 '이값 다음부터 끝까지 개수
        rbyte(6) = "&H" & Hex(id) 'UID
        rbyte(7) = &H6 'FN write single register
        rbyte(8) = &H0 'addr
        rbyte(9) = &H14 'addr
        rbyte(10) = &H0
        rbyte(11) = &H1 'manual


        TARGET_CMD_IDX += 1
        If (TARGET_CMD_IDX > 9999) Then TARGET_CMD_IDX = 0

        Return rbyte
    End Function
    Private Function CHECK_CONNECT_POSSIBLE(ip As String, port As Integer, second As Integer) As Boolean
        '접속가능한지 체크
        Try
            Dim sock As Socket
            sock = New Socket(AddressFamily.InterNetwork, SocketType.Stream, ProtocolType.Tcp)
            Dim result As IAsyncResult = sock.BeginConnect(ip, port, Nothing, Nothing)
            Dim success As Boolean = result.AsyncWaitHandle.WaitOne(second * 1000, True)
            If (sock.Connected) Then
                sock.EndConnect(result)
                sock = Nothing
                Return True
            Else
                sock = Nothing
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region

    Public Sub DEBUG_PRINT(id As Integer, msg As String)

        If (DEBUG_LINE = "ALL" Or DEBUG_LINE = HOUSE_CODE) Then

            If DEBUG_IDS.Contains(id) Or id = 0 Then '0이면 모두 출력

                Debug.Print(msg)

            End If

        End If


    End Sub

End Class
