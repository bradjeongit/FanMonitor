
Module M_Functions
    '코드값을 라인이름으로 변환
    Public Function Get_House_Name(House_Code As String, Optional sep As String = " / ") As String
        Dim hname As String = ""
        If (House_Code.Length <> 5) Then
            hname = ""
        Else
            If (House_Code.Substring(0, 1) = "H" And House_Code.Substring(2, 1) = "R" And (House_Code.Substring(4, 1) = "E" Or House_Code.Substring(4, 1) = "O")) Then
                hname = "온실" & House_Code.Substring(1, 1)
                hname &= sep & House_Code.Substring(3, 1) & "구역"
                If (House_Code.Substring(4, 1) = "E") Then
                    hname &= sep & "짝"
                Else
                    hname &= sep & "홀"
                End If
            End If
        End If
        Return hname
    End Function
    Public Function Get_House_ShortName(House_Code As String, Optional sep As String = " / ") As String
        Dim hname As String = ""
        If (House_Code.Length <> 5) Then
            hname = ""
        Else
            If (House_Code.Substring(0, 1) = "H" And House_Code.Substring(2, 1) = "R" And (House_Code.Substring(4, 1) = "E" Or House_Code.Substring(4, 1) = "O")) Then
                hname = House_Code.Substring(3, 1) & "구역"
                If (House_Code.Substring(4, 1) = "E") Then
                    hname &= sep & "짝"
                Else
                    hname &= sep & "홀"
                End If
            End If
        End If
        Return hname
    End Function

    Public Function Check_Valid_IP(ip As String) As Boolean
        Dim regex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^0-9.-]+") ' [^0-9.-]+ : 점과빼기추가
        If (regex.IsMatch(ip) = True) Then
            'Debug.Print("표현식 안맞음")
            Return False
        Else
            If (ip.Length >= 7 And ip.Length <= 15) Then '1.1.1.1(7개) - 255.255.255.255(15개)
                If (CountCharacter(ip, "."c) <> 3) Then '점 3개 있어야함.
                    'Debug.Print("점 안맞음")
                    Return False
                Else
                    Dim cip As System.Net.IPAddress = Nothing
                    Return System.Net.IPAddress.TryParse(ip, cip)
                End If
            Else
                'Debug.Print("자리수 안맞음")
                Return False
            End If
        End If

    End Function

    Public Function Check_Integer(txt As String) As Boolean
        Dim regex As System.Text.RegularExpressions.Regex = New System.Text.RegularExpressions.Regex("[^0-9]+") ' [^0-9.-]+ : 점과빼기추가
        Return Not (regex.IsMatch(txt))
    End Function

    Public Function CountCharacter(ByVal value As String, ByVal ch As Char) As Integer
        Dim cnt As Integer = 0
        For Each c As Char In value
            If c = ch Then
                cnt += 1
            End If
        Next
        Return cnt
    End Function

    '특정 instance의 특정이름의 값 세팅하기
    Public Function SetPropertyValueByName(obj As [Object], name As String, value As [Object]) As Boolean
        Dim prop = obj.[GetType]().GetProperty(name, System.Reflection.BindingFlags.[Public] Or System.Reflection.BindingFlags.Instance)
        If prop Is Nothing OrElse Not prop.CanWrite Then
            Return False
        End If
        prop.SetValue(obj, value, Nothing)
        Return True
    End Function
End Module
