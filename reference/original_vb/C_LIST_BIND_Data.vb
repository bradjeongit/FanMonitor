Imports System.ComponentModel
Imports System.Runtime.CompilerServices
Public Class C_LIST_BIND_Data
    Implements INotifyPropertyChanged

    ' 값 바뀌었을때 이벤트 발생으로 갱신되게 함.
    Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged
    Protected Sub OnPropertyChanged(<CallerMemberName> Optional name As String = Nothing)
        RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(name))
    End Sub

    Private _DIC_KEY As String '고유값 키 ( MType_ID )
    Private _ID As String 'ID값
    Private _RPM As String '현재RPM
    Private _MType As Integer '모터타입
    Private _USE As Boolean '사용여부
    Private _STATUS As C_HOUSE.MOTOR_STATUS ' 0:정상, 1:이상, 2:중지(사용안함)
    Private _RDATE As DateTime '최근수신일시

    Public Property MType() As Integer
        Get
            Return _MType
        End Get
        Set(ByVal value As Integer)
            If (_MType <> value) Then
                _MType = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public Property DIC_KEY() As String
        Get
            Return _DIC_KEY
        End Get
        Set(ByVal value As String)
            If _DIC_KEY <> value Then
                _DIC_KEY = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public Property ID() As String
        Get
            Return _ID
        End Get
        Set(ByVal value As String)
            If _ID <> value Then
                _ID = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public Property RPM() As String
        Get
            Return _RPM
        End Get
        Set(ByVal value As String)
            If (_RPM <> value) Then
                _RPM = value
                OnPropertyChanged()
            End If
        End Set
    End Property

    Public Property USE() As Boolean
        Get
            Return _USE
        End Get
        Set(ByVal value As Boolean)
            _USE = value
            OnPropertyChanged()
        End Set
    End Property


    Public Property STATUS() As C_HOUSE.MOTOR_STATUS
        Get
            Return _STATUS
        End Get
        Set(ByVal value As C_HOUSE.MOTOR_STATUS)
            _STATUS = value
            OnPropertyChanged()
        End Set
    End Property


    Public Property RDATE() As DateTime
        Get
            Return _RDATE
        End Get
        Set(ByVal value As DateTime)
            _RDATE = value
            'OnPropertyChanged() '화면에 날짜를 따로 표현하지는 않아서 주석처리
        End Set
    End Property

End Class
