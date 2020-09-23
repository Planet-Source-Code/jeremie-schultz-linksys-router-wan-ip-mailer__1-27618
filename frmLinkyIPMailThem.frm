VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form frmLinkyIPMailThem 
   Caption         =   "Linky IP Mail Them"
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7335
   ForeColor       =   &H00000000&
   Icon            =   "frmLinkyIPMailThem.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6525
   ScaleWidth      =   7335
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSubject 
      Height          =   285
      Left            =   120
      TabIndex        =   19
      Text            =   "WAN IP CHANGE"
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Timer tmrCheckIp 
      Left            =   4560
      Top             =   2400
   End
   Begin VB.TextBox txtTimeCheck 
      Height          =   285
      Left            =   3000
      TabIndex        =   13
      Text            =   "60"
      Top             =   5520
      Width           =   375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get WAN IP"
      Height          =   735
      Left            =   5880
      TabIndex        =   8
      Top             =   3480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Text            =   "frmLinkyIPMailThem.frx":0442
      Top             =   2280
      Width           =   4095
   End
   Begin VB.TextBox txtLinkyLanURL 
      Height          =   285
      Left            =   120
      TabIndex        =   6
      Text            =   "http://192.168.1.1/Status.htm"
      Top             =   4800
      Width           =   4095
   End
   Begin VB.TextBox txtLinkyUserName 
      Height          =   285
      Left            =   120
      TabIndex        =   5
      Top             =   5400
      Width           =   2655
   End
   Begin VB.TextBox txtLinkyPassword 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   2655
   End
   Begin VB.TextBox txtLinkyWanIP 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Text            =   "0.0.0.0"
      Top             =   4200
      Width           =   4095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   1
      Top             =   5880
      Width           =   1335
   End
   Begin MSMAPI.MAPIMessages mapMess 
      Left            =   5760
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession mapSess 
      Left            =   5160
      Top             =   2280
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.TextBox txtStatus 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2175
      Left            =   100
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   0
      Width           =   7135
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   6600
      Top             =   2400
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   0
      Top             =   3000
      Width           =   1335
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4560
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.Label lblSubject 
      AutoSize        =   -1  'True
      Caption         =   "Email Subject"
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   3360
      Width           =   960
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Time Delay"
      Height          =   195
      Left            =   3000
      TabIndex        =   18
      Top             =   5280
      Width           =   795
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Seconds"
      Height          =   195
      Left            =   3480
      TabIndex        =   17
      Top             =   5520
      Width           =   630
   End
   Begin VB.Label lblNextCheck 
      AutoSize        =   -1  'True
      Caption         =   "60"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3000
      TabIndex        =   16
      Top             =   6120
      Width           =   225
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Seconds"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3480
      TabIndex        =   15
      Top             =   6120
      Width           =   630
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Time til Next Check"
      Enabled         =   0   'False
      Height          =   195
      Left            =   3000
      TabIndex        =   14
      Top             =   5880
      Width           =   1380
   End
   Begin VB.Label lblLinkyUserName 
      AutoSize        =   -1  'True
      Caption         =   "Linky UserName"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   5160
      Width           =   1170
   End
   Begin VB.Label lblLinkyPassword 
      AutoSize        =   -1  'True
      Caption         =   "Linky Password"
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   1110
   End
   Begin VB.Label lblLinkyLanURL 
      AutoSize        =   -1  'True
      Caption         =   "Linky Lan URL"
      Height          =   195
      Left            =   120
      TabIndex        =   10
      Top             =   4560
      Width           =   1065
   End
   Begin VB.Label lblLinkyWanIP 
      AutoSize        =   -1  'True
      Caption         =   "Linky Wan IP"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   960
   End
End
Attribute VB_Name = "frmLinkyIPMailThem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

Dim bNewSession As Boolean

Private Sub cmdExit_Click()

    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Stop" Then
        cmdStart_Click
    End If

    DoEvents

    Unload Me

Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub cmdStart_Click()

    On Error GoTo ErrorHandler

    If cmdStart.Caption = "&Start" Then
        cmdStart.Caption = "&Stop"
        Log Now & vbTab & "---- Logging on to mail server..."
        LogOn
        Log Now & vbTab & "---- AutoMailer services started"
        'Timer1.Enabled = True
        MailThem 'added
    Else
        cmdStart.Caption = "&Start"
        Log Now & vbTab & "---- AutoMailer services stopped"
        'Timer1.Enabled = False
    End If

Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub Command1_Click()
GetWanIP
End Sub
Sub GetWanIP()
Inet1.URL = txtLinkyLanURL.Text
Inet1.UserName = txtLinkyUserName
Inet1.Password = txtLinkyPassword
Text1.Text = Inet1.OpenURL(Inet1.URL)
Dim intmyStart As Integer
Dim intIPLength As Integer
intmyStart = InStr(1, Text1.Text, "<!--WAN head-->")
If intmyStart > 0 Then
    intmyStart = InStr(intmyStart, Text1.Text, "IP Address:")
    intIPLength = InStr(intmyStart + 46, Text1.Text, "</td>") - intmyStart - 46
    If txtLinkyWanIP.Text = Mid$(Text1.Text, intmyStart + 46, intIPLength) Then
    Else
        txtLinkyWanIP.Text = Mid$(Text1.Text, intmyStart + 46, intIPLength)
        txtSubject.Text = "TES WAN IP CHANGE"
        'txtedit.Text = txtLinkyWanIP.Text
        cmdStart_Click
    End If
End If

tmrCheckIp.Interval = txtTimeCheck * 1000
tmrCheckIp.Enabled = True
End Sub

Private Sub Form_Load()
tmrCheckIp.Interval = txtTimeCheck * 1000

    Dim sDBLocation As String
    Dim hSysMenu As Long

    On Error GoTo ErrorHandler

    ' disable the 'X':
    hSysMenu = GetSystemMenu(hwnd, False)
    RemoveMenu hSysMenu, SC_CLOSE, MF_BYCOMMAND

    Log "Linky IP Mailer version " & App.Major & "." & App.Minor & App.Revision
    Log "Press Get WAN IP to begin services"

    sDBLocation = App.Path & "\data.mdb"

    Set cn = New ADODB.Connection
    cn.Open "PROVIDER=MSDASQL;" & _
                "DRIVER={Microsoft Access Driver (*.mdb)};" & _
                "DBQ=" & sDBLocation & ";" & _
                "UID=;PWD=;"

Exit Sub

ErrorHandler:
    Select Case Err.Number
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub

Private Sub Form_Unload(Cancel As Integer)

    End

End Sub


Private Sub m_send_mail(sRecipient As String, sMessage As String)

    On Error GoTo ErrorHandler
    
    Dim strMessage As String
    
    mapMess.Compose

    mapMess.RecipAddress = sRecipient

    mapMess.AddressResolveUI = True
    mapMess.ResolveName

    mapMess.MsgSubject = txtSubject.Text
    mapMess.MsgNoteText = txtLinkyWanIP.Text 'sMessage
    
    mapMess.Send False 'set this to true to view message and then send manually or cancel
    
Exit Sub

ErrorHandler:
    Select Case Err.Number

        Case Else
            Log Err.Number & " - " & Err.Description

    End Select
End Sub

Private Function LogOn() As Boolean

    On Error GoTo ErrorHandler
    
    If mapSess.NewSession Then
        ' Session already established
        LogOn = True
        Exit Function
    End If
    

    With mapSess
        ' Set DownLoadMail to False to prevent immediate download.
        .DownLoadMail = False
        .LogonUI = True ' Use the underlying email system's logon UI.
        '-----------------------------------------------------------------------
        '.LogonUI = False
        '.UserName = "username"  ' Uncomment these lines, add your username and password
        '.Password = "password"   '    to eliminate the logon screen
        '-----------------------------------------------------------------------
        .SignOn
        ' If successful, return True
        LogOn = True
        ' Set NewSession to True and set
        ' variable flag to true
        .NewSession = True
        bNewSession = .NewSession
        mapMess.SessionID = .SessionID ' You must set this before continuing.
    End With
    
Exit Function

ErrorHandler:
    Select Case Err.Number
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Function

Private Sub Log(ByVal sText As String)

    ' this way it doesnt refresh the whole thing every time, no blinking
    With txtStatus
        .SelStart = Len(.Text)
        .SelText = sText & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10)
        .SelLength = 0
    End With

End Sub

Private Sub tmrCheckIp_Timer()
GetWanIP
End Sub

Private Sub txtStatus_GotFocus()

    frmLinkyIPMailThem.SetFocus

End Sub
Private Sub MailThem()
    ' If the Flag field in the database is set to 1 the e-mail address for
    ' this record will be sent an e-mail.

    On Error GoTo ErrorHandler
    
    If LogOn = False Then
        Log Now & vbTab & "---- Error logging on to mail server!  Will try again."
        Exit Sub
    End If

    Set rs = New ADODB.Recordset
    rs.CursorType = adOpenKeyset
    rs.LockType = adLockOptimistic
    rs.Open "Select * from email where Flag = 1", cn, , , adCmdUnknown
    rs.MoveFirst

    If rs.RecordCount <> 0 Then
        While rs.EOF = False
            
            m_send_mail rs.Fields("EMail"), "stuff here"
            
            DoEvents

            Log Now & vbTab & "Mail sent to:  " & rs.Fields("EMail")
            
            rs.MoveNext
        Wend
    Else
        rs.Close
    End If
cmdStart_Click
Exit Sub

ErrorHandler:
    Select Case Err.Number
        Case 3021
            Log Now & vbTab & "---- No requests"
    
        Case Else
            Log Err.Number & " - " & Err.Description

    End Select

End Sub
