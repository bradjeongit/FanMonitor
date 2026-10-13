Module M_Variables

    '설정파일 이름 및 위치
    Public Config_File_Name As String = "Config.ini"
    Public Config_File_Path As String = AppDomain.CurrentDomain.BaseDirectory & "\"

    '각 영역을 담는 사전
    Public HOUSES As New Dictionary(Of String, C_HOUSE)

    '라인리스트
    Public HOUSE_LINES As New List(Of String)({
        "H1R1E", "H1R1O", "H1R2E", "H1R2O", "H1R3E", "H1R3O", "H1R4E", "H1R4O",
        "H2R1E", "H2R1O", "H2R2E", "H2R2O", "H2R3E", "H2R3O", "H2R4E", "H2R4O"
    })

    'RPM최대값
    Public MAX_RPM As Integer = 1100

    Public Structure Struct_Config
        Dim SOURCE_IP As String
        Dim SOURCE_PORT As Integer
        Dim TARGET_IP As String
        Dim TARGET_PORT As Integer
        Dim SOURCE_MANUAL As Boolean
        Dim MANUAL_RPM As Integer
        Dim SOURCE_LINE As String
        Dim SOURCE_ID As Integer
    End Structure

    Public MCONFIG As New Dictionary(Of String, Struct_Config)


End Module
