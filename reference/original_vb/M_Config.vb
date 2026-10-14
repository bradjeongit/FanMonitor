Module M_Config

    Public Sub Get_Local_Config()
        Dim t As String = ""

        '라인 추가
        Dim idx As Integer = 1
        For Each s As String In HOUSE_LINES
            '라인추가
            HOUSES.Add(s, New C_HOUSE(s))

            '설정읽기
            '소스 ip
            Dim sip As String = INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_IP", "192.168.0." & idx)
            '소스 port
            Dim sport As Integer = 0
            Integer.TryParse(INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_PORT", "502"), sport) '기본 포트 502
            '타겟 ip
            Dim tip As String = INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_TARGEIP", "192.168.0." & (idx + 1))
            '타겟 port
            Dim tport As Integer = 0
            Integer.TryParse(INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_TARGEPORT", "502"), tport) '기본 포트 502
            '수동여부
            t = INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_MANUAL", "1") '기본 수동으로
            Dim source_manual As Boolean = True
            If (t = "0") Then source_manual = False
            '수동rpm
            Dim manual_rpm As Integer = 0
            Integer.TryParse(INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_MANUAL_RPM", "0"), manual_rpm) '기본속도 0으로
            '소스 라인이름
            Dim source_line As String = INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_LINE", "")
            '소스 ID
            Dim source_id As Integer = 0
            Integer.TryParse(INIRead(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_ID", "0"), source_id)

            '라인설정추가
            MCONFIG.Add(s, New Struct_Config With {
                        .SOURCE_IP = sip, .SOURCE_PORT = sport,
                        .TARGET_IP = tip, .TARGET_PORT = tport,
                        .SOURCE_MANUAL = source_manual, .MANUAL_RPM = manual_rpm,
                        .SOURCE_LINE = source_line, .SOURCE_ID = source_id})


            Dim motor_data As String = INIRead(Config_File_Path & Config_File_Name, "MOTOR_CONFIG", s & "_MOTORS", "")
            If motor_data.Length > 0 Then
                Dim sp() As String = motor_data.Split(","c)
                For i As Integer = 0 To sp.Length - 1
                    Dim ssp() As String = sp(i).Split("="c)
                    Dim id As Integer = 0
                    Dim mtype As Integer = 0
                    Dim t_use As Integer = 0
                    Dim use As Boolean = False
                    Integer.TryParse(ssp(0), id)
                    Integer.TryParse(ssp(1), mtype)
                    Integer.TryParse(ssp(2), t_use)
                    If (t_use <> "0") Then use = True
                    HOUSES(s).Add_Motor(id, mtype, use)
                Next
            End If

            idx += 2
        Next

        '모터등록(임시모터 )
        'Dim r As New Random
        'For Each k As String In HOUSES.Keys
        '    '모터추가
        '    For i As Integer = 0 To 10
        '        Dim id As Integer = r.Next(0, 100)
        '        Dim mtype As C_HOUSE.MOTOR_TYPE = r.Next(0, 2)
        '        Dim use As Boolean = True
        '        HOUSES(k).Add_Motor(id, mtype, use)
        '    Next
        'Next

        Put_Local_Config()

    End Sub

    Public Sub Put_Local_Config()

        For Each s As String In HOUSE_LINES

            '설정쓰기
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_IP", MCONFIG(s).SOURCE_IP)
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_PORT", MCONFIG(s).SOURCE_PORT.ToString)
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_TARGEIP", MCONFIG(s).TARGET_IP)
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_TARGEPORT", MCONFIG(s).TARGET_PORT.ToString)
            If (MCONFIG(s).SOURCE_MANUAL = True) Then
                INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_MANUAL", "1")
            Else
                INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_MANUAL", "0")
            End If
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_MANUAL_RPM", MCONFIG(s).MANUAL_RPM)
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_LINE", MCONFIG(s).SOURCE_LINE)
            INIWrite(Config_File_Path & Config_File_Name, "LINE_CONFIG", s & "_SOURCE_ID", MCONFIG(s).SOURCE_ID)

            Dim mlistdata As String = ""
            Dim mlist As System.Collections.ObjectModel.ObservableCollection(Of C_LIST_BIND_Data) = HOUSES(s).GetList
            For Each m As C_LIST_BIND_Data In mlist
                Dim mdata As String = ""
                mdata &= m.ID & "="
                mdata &= m.MType & "="
                If (m.USE = True) Then mdata &= 1 Else mdata &= 0
                If (mlistdata <> "") Then mlistdata &= ","
                mlistdata &= mdata
            Next
            INIWrite(Config_File_Path & Config_File_Name, "MOTOR_CONFIG", s & "_MOTORS", mlistdata)

        Next

    End Sub

#Region "INI 라이브러리"

    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String, ByVal lpFileName As String) As Int32
    Private Declare Unicode Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringW" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Int32, ByVal lpFileName As String) As Int32
    Public Function INIRead(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String, ByVal DefaultValue As String) As String
        ' primary version of call gets single value given all parameters
        Dim n As Int32
        Dim sData As String
        sData = Space$(1024) ' allocate some room 
        n = GetPrivateProfileString(SectionName, KeyName, DefaultValue, sData, sData.Length, INIPath)
        If n > 0 Then ' return whatever it gave us
            INIRead = Replace(Replace(sData.Substring(0, n), "&", ""), "|", "") ' &나 | 기호가 있으면 안됨.
        Else
            INIRead = ""
        End If
    End Function
    Public Sub INIWrite(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String, ByVal TheValue As String)
        Call WritePrivateProfileString(SectionName, KeyName, TheValue, INIPath)
    End Sub
    Public Sub INIDelete(ByVal INIPath As String, ByVal SectionName As String, ByVal KeyName As String) ' delete single line from section
        Call WritePrivateProfileString(SectionName, KeyName, Nothing, INIPath)
    End Sub

#End Region




End Module
