VERSION 5.00
Begin VB.UserControl MX 
   ClientHeight    =   585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   615
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   585
   ScaleWidth      =   615
   Begin VB.Label Title 
      Caption         =   "MX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   615
   End
End
Attribute VB_Name = "MX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private Sub UserControl_Initialize()
    SetWinVersion
End Sub

Private Sub UserControl_Resize()
    If UserControl.Width <> 32 Then
        UserControl.Width = 400
    End If
    If UserControl.Height <> 32 Then
        UserControl.Height = 250
    End If
    Title.Left = 0
End Sub

Public Function GetMX() As String
    If IsNetConnectOnline = True Then
        If Not IsNetConnectViaProxy Then
            GetMX = MX_Query
        Else
            Err.Raise 0, "GetMX", "This computer is connected via a proxy server." & vbCrLf & "At this time, the wMX control does not support proxy servers."
            Exit Function
        End If
    Else
        Err.Raise 0, "GetMX", "This computer is not currently connected to the internet."
    End If
End Function

Public Property Get DNSCount() As Integer
    DNSCount = mi_DNSCount
End Property

Public Property Get MXCount() As Integer
    MXCount = mi_MXCount
End Property

Public Property Get PrefCount() As Integer
    PrefCount = mi_MXCount
End Property

Public Property Get Domain() As String
    Domain = ms_Domain
End Property

Public Property Let Domain(ByVal New_Domain As String)
    If Len(New_Domain) > 4 Then 'its a good host
        ms_Domain = New_Domain
    End If
End Property

Public Function DNS(ByVal Index As String) As String
    DNS = sDNS(Index)
End Function

Public Function MX(ByVal Index As String) As String
    MX = sMX(Index)
End Function

Public Function Pref(ByVal Index As String) As String
    Pref = sPref(Index)
End Function

