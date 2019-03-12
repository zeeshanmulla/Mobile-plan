VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Computer Telephony Integration"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   600
      Top             =   2280
   End
   Begin VB.TextBox txtdisplay 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   1320
      TabIndex        =   21
      Top             =   960
      Width           =   1815
   End
   Begin VB.CommandButton cmdend 
      Caption         =   "End"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2280
      TabIndex        =   20
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdcall 
      Caption         =   "Call"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   19
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdhash 
      Caption         =   "#"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   18
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmd0 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   17
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmdstar 
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   16
      Top             =   3480
      Width           =   615
   End
   Begin VB.CommandButton cmd9 
      Caption         =   "9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmd8 
      Caption         =   "8"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   14
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmd7 
      Caption         =   "7"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   13
      Top             =   2880
      Width           =   615
   End
   Begin VB.CommandButton cmd6 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   12
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmd5 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   11
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton cmd4 
      Caption         =   "4"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   10
      Top             =   2280
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   9
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmd2 
      Caption         =   "2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1920
      TabIndex        =   8
      Top             =   1680
      Width           =   615
   End
   Begin VB.CommandButton cmd1 
      Caption         =   "1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1320
      TabIndex        =   7
      Top             =   1680
      Width           =   615
   End
   Begin MSCommLib.MSComm MSComm2 
      Left            =   6720
      Top             =   5640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
   End
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   6000
      Top             =   3240
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   5280
      Top             =   3960
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   17520
      Top             =   2040
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   495
      Left            =   15120
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   15360
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   12600
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Label lblbalance 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   23
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "Balance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   22
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label6 
      Caption         =   "Label4"
      Height          =   495
      Left            =   13200
      TabIndex        =   6
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Label4"
      Height          =   495
      Left            =   5880
      TabIndex        =   5
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   1335
      Left            =   19320
      TabIndex        =   4
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Label Label2 
      Caption         =   "No."
      Height          =   255
      Left            =   14640
      TabIndex        =   1
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Computer Telephony Integration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strBuffer As String 'receive buffer
Dim cbm As Boolean
Dim cbt As Boolean
Dim count1 As String
Dim user2 As String
Dim user3 As String
Dim user4 As String
Dim user As String
Dim n1 As Integer
Dim userlen As Integer
Dim calllen As Integer
Dim hault As Boolean


Dim oXL As Excel.Application
      Dim oWB As Excel.Workbook
      Dim oSheet As Excel.Worksheet
      Dim oRng As Excel.Range






Private Sub cmd0_Click()
txtdisplay.Text = txtdisplay.Text + "0"

End Sub

Private Sub cmd1_Click()
txtdisplay.Text = txtdisplay.Text + "1"
End Sub

Private Sub cmd2_Click()
txtdisplay.Text = txtdisplay.Text + "2"

End Sub

Private Sub cmd4_Click()
txtdisplay.Text = txtdisplay.Text + "4"

End Sub

Private Sub cmd5_Click()
txtdisplay.Text = txtdisplay.Text + "5"

End Sub

Private Sub cmd6_Click()
txtdisplay.Text = txtdisplay.Text + "6"

End Sub

Private Sub cmd7_Click()
txtdisplay.Text = txtdisplay.Text + "7"

End Sub

Private Sub cmd8_Click()
txtdisplay.Text = txtdisplay.Text + "8"

End Sub

Private Sub cmd9_Click()
txtdisplay.Text = txtdisplay.Text + "9"

End Sub

Private Sub cmdcall_Click()
cmd1.Enabled = False
cmd2.Enabled = False
Command4.Enabled = False
cmd4.Enabled = False
cmd5.Enabled = False
cmd6.Enabled = False
cmd7.Enabled = False
cmd8.Enabled = False
cmd9.Enabled = False
cmd0.Enabled = False
cmdstar.Enabled = False
cmdhash.Enabled = False

lblbalance.Caption = oSheet.Cells(2, 6).Value

lblbalance.Caption = Val(lblbalance.Caption) - 1

Timer1.Enabled = True
calllen = Len(txtdisplay)
If calllen >= "10" Then
oSheet.Cells(count1, 1).Value = count1 - 1
oSheet.Cells(count1, 2).Value = txtdisplay.Text
oSheet.Cells(count1, 3).Value = Date$
oSheet.Cells(count1, 4).Value = Time$
count1 = count1 + 1
oSheet.Cells(2, 5).Value = count1



    MSComm1.Output = "atd "
    Call delay1
    MSComm1.Output = txtdisplay.Text
    Call delay1
    MSComm1.Output = ";"
    Call delay1
    MSComm1.Output = Chr(13)
    Call delay1
End If


End Sub

Private Sub cmdend_Click()
Timer1.Enabled = False
oSheet.Cells(2, 6).Value = lblbalance.Caption
txtdisplay.Text = ""
    MSComm1.Output = "ath"
    Call delay1
    MSComm1.Output = Chr(13)
    Call delay1
    
    cmd1.Enabled = True
cmd2.Enabled = True
Command4.Enabled = True
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = True
cmd7.Enabled = True
cmd8.Enabled = True
cmd9.Enabled = True
cmd0.Enabled = True
cmdstar.Enabled = True
cmdhash.Enabled = True

End Sub

Private Sub Command4_Click()
txtdisplay.Text = txtdisplay.Text + "3"

End Sub

Private Sub Form_Load()
    With MSComm1
       .CommPort = 7
        .Settings = "9600,N,8,1"
        .Handshaking = comNone
        .RTSEnable = True
        .DTREnable = True
        .RThreshold = 1
        .SThreshold = 0
        .InputMode = comInputModeBinary
        .InputLen = 0
        .PortOpen = True 'must be the last
    End With
    
 hault = False
    
    
    Set oXL = CreateObject("excel.application")
    oXL.Visible = True
    
Set oWB = oXL.Workbooks.Open(App.Path & "\book1.xls")
Set oSheet = oWB.Worksheets("Sheet1")
count1 = oSheet.Cells(2, 5).Value
lblbalance.Caption = oSheet.Cells(2, 6).Value
'count1 = count1 + 1
    MSComm1.Output = "at"
    Call delay1
 MSComm1.Output = Chr(13)
   Call delay1
   MSComm1.Output = "at+cmgf=1"
    Call delay1
    MSComm1.Output = Chr(13)
    Call delay1
    MSComm1.Output = "at+cnmi=2,2,2,0,0"
    Call delay1
    MSComm1.Output = Chr(13)
    Call delay1

End Sub

Function delay1()
Dim delaytime
delaytime = Timer()
While (Timer() - delaytime) < 0.1
'do nothing
Wend
End Function
Private Function Receive() As String
    Dim strPart As String
    Dim strInput As String
    strInput = ""
    Do
        strPart = ""
        Call Delay(1)
        strPart = strBuffer
        strBuffer = ""
        If strPart = "" Then Exit Do
        strInput = strInput & strPart
    Loop
    If strInput <> "" Then Call Trace(">> " & strInput)
    Receive = strInput
End Function

Private Sub Delay(ByVal HowLong As Date)
    Dim endDate As Date
    endDate = DateAdd("s", HowLong, Now)
    While endDate > Now
        DoEvents 'Allows windows to handle other stuff
    Wend
End Sub

Private Sub Trace(ByVal message As String)
   ' Dim strLine As String
   ' strLine = DateTime.Now & " " & message
   ' txtLog.Text = txtLog.Text & message & vbCrLf
   ' txtLog.SelStart = Len(txtLog.Text)
End Sub



Private Sub Form_Unload(Cancel As Integer)
oWB.Save
oXL.Quit
 Set oRng = Nothing
      Set oSheet = Nothing
      Set oWB = Nothing
      Set oXL = Nothing
     
      MSComm1.PortOpen = False
End

End Sub

Private Sub MSComm1_OnComm()
    Dim strMessage As String
   Select Case MSComm1.CommEvent
        ' Event messages.
       Case comEvReceive
            strMessage = StrConv(MSComm1.Input, vbUnicode)
    End Select
    strBuffer = strBuffer & strMessage
    Label3.Caption = strBuffer
    Call Receive

End Sub

Private Sub Timer1_Timer()
lblbalance.Caption = Val(lblbalance.Caption) - 1
End Sub

Private Sub Timer2_Timer()
Label3.Caption = LCase$(Label3.Caption)
If InStr(Label3.Caption, "+cmt:") > 0 Then
cbt = True
'Label7.Caption = cbt
Label5.Caption = Mid$(Label3.Caption, 50)
Label6.Caption = Mid$(Label3.Caption, 10, 13)
'lblmess.Caption = Mid$(Label3.Caption, 50)
Label5.Caption = UCase$(Label5.Caption)
Label5.Caption = Trim$(Label5.Caption)
'Call Delay(2)
Timer3.Enabled = True
Timer5.Enabled = True

oSheet.Cells(2, 6).Value = Val(lblbalance.Caption) + Val(Label5.Caption)
lblbalance.Caption = oSheet.Cells(2, 6).Value
hault = False

cmd1.Enabled = True
'Call Delay(2)
cmd2.Enabled = True
Command4.Enabled = True
cmd4.Enabled = True
cmd5.Enabled = True
cmd6.Enabled = True
cmd7.Enabled = True
cmd8.Enabled = True
cmd9.Enabled = True
cmd0.Enabled = True
cmdstar.Enabled = True
cmdhash.Enabled = True
cmdcall.Enabled = True
cmdend.Enabled = True

Label3.Caption = ""
Label5.Caption = ""


Else
cbt = False
'Label7.Caption = cbt

End If

End Sub


Private Sub Timer3_Timer()
If Val(lblbalance.Caption) = 0 And hault = False Then
    MSComm1.Output = "ath"
    Call delay1
    MSComm1.Output = Chr(13)
    Call delay1
hault = True
txtdisplay.Text = ""
Timer1.Enabled = False
    
cmd1.Enabled = False
cmd2.Enabled = False
Command4.Enabled = False
cmd4.Enabled = False
cmd5.Enabled = False
cmd6.Enabled = False
cmd7.Enabled = False
cmd8.Enabled = False
cmd9.Enabled = False
cmd0.Enabled = False
cmdstar.Enabled = False
cmdhash.Enabled = False
cmdcall.Enabled = False
cmdend.Enabled = False


End If
End Sub

Private Sub Timer5_Timer()

Timer5.Enabled = False
End Sub
