VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "E-Link Communications"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   "1"
      Top             =   4410
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2160
      TabIndex        =   4
      Text            =   "<none>"
      Top             =   1680
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      Height          =   1935
      Left            =   960
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2160
      TabIndex        =   3
      Top             =   1320
      Width           =   3255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   240
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   600
      Width           =   3255
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2160
      TabIndex        =   2
      Top             =   960
      Width           =   3255
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4920
      Top             =   -120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin Project1.MX MX1 
      Left            =   4440
      Top             =   0
      _ExtentX        =   714
      _ExtentY        =   450
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000E&
      Caption         =   "Number of times to send :"
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   4440
      Width           =   2295
   End
   Begin VB.Label Command1 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "                   SEND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   615
      Left            =   4080
      TabIndex        =   13
      Top             =   4080
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Rcp.'s Email:"
      Height          =   255
      Left            =   960
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      Height          =   255
      Left            =   960
      TabIndex        =   11
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "From:"
      Height          =   255
      Left            =   960
      TabIndex        =   10
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "From (name):"
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Rcp. Name:"
      Height          =   255
      Left            =   960
      TabIndex        =   8
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Idle"
      Height          =   255
      Left            =   960
      TabIndex        =   5
      Top             =   4080
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   4995
      Left            =   -1320
      Picture         =   "Form1.frx":0442
      Stretch         =   -1  'True
      Top             =   0
      Width           =   16500
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-----------------------------------------------------
'Written by     : Daniel Ho
'Email him at   : daniel020@hotmail.com
'Original Code  : Bryan Cairns
'Finished Date  : 28/11/01
'------------------------------------------------------
Dim response As String, Reply As Integer, DateNow As String
Dim first As String, Second As String, Third As String
Dim Fourth As String, Fifth As String, Sixth As String
Dim Seventh As String, Eighth As String
Dim start As Single, Tmr As Single

Private Sub Command1_Click()
Dim sDNS As String
Dim sMailServer As String
Dim sDomain As String

Text7.Enabled = False

sDomain = GetDomainFromAddr(Text3.Text)
If sDomain = "" Then
MsgBox "Please Enter a VALID address!", vbCritical, "Error"
Exit Sub
End If

MX1.Domain = sDomain
sMailServer = MX1.GetMX

If sMailServer = "" Then
MsgBox "Sorry could not locate the mail server for this address.", vbInformation, "Opps..."
Exit Sub
End If

If MX1.DNSCount = 0 Then
MsgBox "Could not Get Local DNS!", vbCritical, "Error"
Exit Sub
End If

sDNS = MX1.DNS(0)

'Error checking for this demo only
If sDNS = "" Then
MsgBox "Could not Retrive Local DNS Server!" & vbCrLf & "Please check your internet settings as EVERYONE has a DNS!", vbCritical, "Opps..."
Exit Sub
End If
'''''''''''''''''''''''''''''''''''

Label3.Caption = "Using Server = " & sMailServer
SendEmail sMailServer, Text5.Text, Text4.Text, Text6.Text, Text3.Text, Text1.Text, Text2.Text
MsgBox "Your mail(s) has been sent.", vbInformation, "Send Mail"
Label3.Caption = "Done"
Text7.Enabled = True

End Sub

Private Function GetMX(sServer As String, sDNS As String) As String
With wsMX
.RemoteHost = sDNS
.RemotePort = 53 'mx lookup port
.connect
End With
End Function

Public Function GetDomainFromAddr(sAddr As String) As String
Dim Ipos As Long
Ipos = InStr(1, sAddr, "@", vbBinaryCompare)
If Ipos > 0 Then
GetDomainFromAddr = Mid(sAddr, Ipos + 1, Len(sAddr))
Exit Function
End If
GetDomainFromAddr = ""
End Function

Sub SendEmail(MailServerName As String, FromName As String, FromEmailAddress As String, ToName As String, ToEmailAddress As String, EmailSubject As String, EmailBodyOfMessage As String)
          
Dim Index
For Index = 1 To Int(Text7.Text) ' loop to send email more than once...

    Winsock1.LocalPort = 0 ' Must set local port to 0 (Zero) or you can only send 1 e-mail pre program start
    
If Winsock1.State = sckClosed Then ' Check to see if socet is closed
    DateNow = Format(Date, "Ddd") & ", " & Format(Date, "dd Mmm YYYY") & " " & Format(Time, "hh:mm:ss") & "" & " -0600"
    first = "mail from:" + Chr(32) + FromEmailAddress + vbCrLf ' Get who's sending E-Mail address
    Second = "rcpt to:" + Chr(32) + ToEmailAddress + vbCrLf ' Get who mail is going to
    Third = "Date:" + Chr(32) + DateNow + vbCrLf ' Date when being sent
    Fourth = "From:" + Chr(32) + FromName + vbCrLf ' Who's Sending
    Fifth = "To:" + Chr(32) + ToNametxt + vbCrLf ' Who it going to
    Sixth = "Subject:" + Chr(32) + EmailSubject + vbCrLf ' Subject of E-Mail
    Seventh = EmailBodyOfMessage + vbCrLf ' E-mail message body
    Ninth = "X-Mailer: EBT Reporter v 2.x" + vbCrLf ' What program sent the e-mail, customize this
    Eighth = Fourth + Third + Ninth + Fifth + Sixth  ' Combine for proper SMTP sending

    Winsock1.protocol = sckTCPProtocol ' Set protocol for sending
    Winsock1.RemoteHost = MailServerName ' Set the server address
    Winsock1.RemotePort = 25 ' Set the SMTP Port
    Winsock1.connect ' Start connection
    
    WaitFor ("220")
    
    Label3.Caption = "Connecting...."
    Label3.Refresh
    
    Winsock1.SendData ("HELO worldcomputers.com" + vbCrLf)

    WaitFor ("250")

    Label3.Caption = "Connected"
    Label3.Refresh

    Winsock1.SendData (first)

    Label3.Caption = "Sending Message"
    Label3.Refresh

    WaitFor ("250")

    Winsock1.SendData (Second)

    WaitFor ("250")

    Winsock1.SendData ("data" + vbCrLf)
    
    WaitFor ("354")


    Winsock1.SendData (Eighth + vbCrLf)
    Winsock1.SendData (Seventh + vbCrLf)
    Winsock1.SendData ("." + vbCrLf)

    WaitFor ("250")

    Winsock1.SendData ("quit" + vbCrLf)
    
    Label3.Caption = "Disconnecting"
    Label3.Refresh

    WaitFor ("221")

    Winsock1.Close
Else
    MsgBox (Str(Winsock1.State))
End If
   
Next Index


   
End Sub
Sub WaitFor(ResponseCode As String)
    start = Timer ' Time event so won't get stuck in loop
    While Len(response) = 0
        Tmr = start - Timer
        DoEvents ' Let System keep checking for incoming response **IMPORTANT**
        If Tmr > 50 Then ' Time in seconds to wait
            MsgBox "SMTP service error, timed out while waiting for response", 64, MsgTitle
            Exit Sub
        End If
    Wend
    While Left(response, 3) <> ResponseCode
        DoEvents
        If Tmr > 50 Then
            MsgBox "SMTP service error, impromper response code. Code should have been: " + ResponseCode + " Code recieved: " + response, 64, MsgTitle
            Exit Sub
        End If
    Wend
response = "" ' Sent response code to blank **IMPORTANT**
End Sub

Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbWhite
Command1.ForeColor = RGB(32, 32, 32)
End Sub

Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H808080
Command1.ForeColor = vbWhite
End Sub

Private Sub Form_Unload(Cancel As Integer)
Winsock1.Close
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Winsock1.GetData response ' Check for incoming response *IMPORTANT*
End Sub

