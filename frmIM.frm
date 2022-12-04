VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmIM 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chatting: "
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   4740
   Icon            =   "frmIM.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   4740
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar prgXfr 
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4725
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   1e14
      Scrolling       =   1
   End
   Begin VB.TextBox txtTime 
      Height          =   285
      Left            =   4560
      TabIndex        =   5
      Top             =   4200
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtInfo 
      Height          =   285
      Left            =   4560
      TabIndex        =   4
      Top             =   4560
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox txtOut 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Timer tmrIM 
      Interval        =   20000
      Left            =   0
      Top             =   3840
   End
   Begin MSComctlLib.StatusBar staIM 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5895
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   8308
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfIn 
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   240
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7858
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmIM.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   7
      Top             =   4755
      Width           =   60
   End
   Begin VB.Label lblSend 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   345
      Left            =   3765
      TabIndex        =   1
      Top             =   5280
      Width           =   810
   End
   Begin VB.Menu mnuActions 
      Caption         =   "&Actions"
      Begin VB.Menu mnuIWant 
         Caption         =   "&I Want To..."
         Begin VB.Menu mnuLeave 
            Caption         =   "&Leave This Conversation"
         End
         Begin VB.Menu mnuInvite 
            Caption         =   "Invite &Someone To This Conversation"
         End
         Begin VB.Menu mnuSend 
            Caption         =   "&Send A File"
         End
         Begin VB.Menu mnuSpace 
            Caption         =   "-"
         End
         Begin VB.Menu mnuCancel 
            Caption         =   "&Cancel This Transfer"
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "&Options"
      Begin VB.Menu mnuFont 
         Caption         =   "&Font"
         Begin VB.Menu mnuFonts 
            Caption         =   "-"
            Index           =   0
         End
      End
   End
End
Attribute VB_Name = "frmIM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private strFont As String
Private strFileName As String


Private Sub Form_Load()

    Dim intfont As Integer

    mnuFonts(0).Caption = Screen.Fonts(0)
    
    For intfont = 1 To Screen.FontCount - 1
        Load mnuFonts(intfont)
        mnuFonts(0).Caption = Screen.Fonts(intfont)
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    On Error Resume Next
    frmOnline.XMSNC1.MSNLeaveSession CInt(Me.Tag)
    Me.Tag = ""
    Unload Me

End Sub






Private Sub Label1_Click()

End Sub

Private Sub lblSend_Click()

    If strFont = "" Then
        strFont = "Verdana"
    End If
    
    strFont = Replace(strFont, " ", "%20")

    frmOnline.XMSNC1.MSNMessage CInt(Me.Tag), txtOut.Text, strFont
    rtfIn.SelColor = vbBlack
    rtfIn.SelStart = Len(rtfIn.Text)
    rtfIn.SelFontSize = 10
    rtfIn.SelFontName = "Tahoma"
    rtfIn.SelText = frmOnline.XMSNC1.MSNFriendlyName & " says:" & vbCrLf & "   "
    strFont = Replace(strFont, "%20", " ")
    rtfIn.SelFontName = strFont
    rtfIn.SelStart = Len(rtfIn.Text)
    rtfIn.SelText = txtOut.Text & vbCrLf
    txtOut.Text = ""

End Sub

Private Sub lblSend_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    lblSend.MousePointer = 1

End Sub

Private Sub mnuCancel_Click()

    frmOnline.XMSNC1.MSNCancelFileTransfer (CInt(Me.Tag))
    mnuCancel.Enabled = False

End Sub

Private Sub mnuFonts_Click(Index As Integer)

    strFont = mnuFonts(Index).Caption
    txtOut.Font = strFont

End Sub

Private Sub mnuInvite_Click()

    Dim strContacts As String, arrContacts() As String
    Dim intLoop As Integer, arrSplitContacts() As String
    Dim mNode As Node
    
    strContacts = frmOnline.XMSNC1.MSNRetrieveContacts
    arrContacts() = Split(strContacts, ";")
        
    For intLoop = 0 To UBound(arrContacts)
        If arrContacts(intLoop) <> "" Then
            arrSplitContacts() = Split(arrContacts(intLoop), ",")
            Set mNode = frmOnline.tvContacts.Nodes(arrSplitContacts(0))
            If InStr(txtInfo.Text, arrSplitContacts(0)) = 0 And mNode.Parent <> "Offline" Then
                frmInvite.lstInvite.AddItem arrSplitContacts(1)
                frmInvite.lstInvite1.AddItem arrSplitContacts(0)
            End If
        End If
    Next intLoop
    
    If frmInvite.lstInvite.ListCount = 0 Then
        MsgBox "Either you are chatting with all of your contacts currently, or" & vbCrLf & "all of your remaining contacts are offline."
        Exit Sub
    End If
    
    frmInvite.Tag = Me.Tag
    frmInvite.Show vbModal

End Sub

Private Sub mnuLeave_Click()
    
    Unload Me
    Me.Hide

End Sub

Private Sub mnuSend_Click()
    
    Dim strResult As String, strPath As String, strName As String, lngSize As Long
    Dim intChar As Integer, lngKbs As Long, lngTime As Long, strTime As String
    
    On Error GoTo ErrHandler
    
    strResult = ShowOpen(Me, , , "Select A File To Send")
    
    If strResult = "" Then
        GoTo ErrHandler
    End If
    
    lngSize = FileLen(strResult)
    intChar = InStrRev(strResult, "\")
    strPath = Mid(strResult, 1, intChar)
    strName = Mid(strResult, (intChar + 1))
    strName = Replace(strName, Chr(0), "")
    
    frmOnline.XMSNC1.MSNSendFile CInt(Me.Tag), strPath, strName, lngSize
    
    lngKbs = (lngSize / 1024)
    
    lngKbs = (lngKbs * 8)
    lngTime = (lngKbs / 32) * 2
    
    If lngTime < 60 Then
    
        strTime = "seconds"
    
    ElseIf lngTime >= 60 Then
        lngTime = (lngTime \ 60)
        strTime = "minutes"
    
    End If
    
    prgXfr.Tag = strName
    prgXfr.Min = 0
    prgXfr.Max = lngSize
    prgXfr.Value = 0
    prgXfr.Visible = True
    lblStatus.Caption = "Sending..."
    rtfIn.SelStart = Len(rtfIn.Text)
    rtfIn.SelColor = vbBlack
    rtfIn.SelBold = True
    rtfIn.SelText = vbCrLf & "Sending: " & strName & vbCrLf & "Approx Transfer Time: " & lngTime & " " & strTime & " @ 28.8 Kbps" & vbCrLf & vbCrLf
    rtfIn.SelColor = vbBlack
    rtfIn.SelBold = False
    
    strFileName = strName
    mnuCancel.Enabled = True
    
    Exit Sub
    
ErrHandler:
    Exit Sub

End Sub

Private Sub tmrIM_Timer()

    If txtInfo.Text = "" Then
        staIM.Panels(1).Text = "Waiting"
    Else
        staIM.Panels(1).Text = "Last Message Received: " & txtTime.Text
    End If

End Sub

Private Sub txtOut_KeyPress(KeyAscii As Integer)

    Static intTyped As Integer

    If KeyAscii = 13 Then
        lblSend_Click
        staIM.Panels(1).Text = "Last Message Received: " & txtTime.Text
        Exit Sub
    End If
    
    staIM.Panels(1).Text = "Typing a message"
    
    If intTyped = 0 Then
        intTyped = (intTyped + 1)
        frmOnline.XMSNC1.MSNTyping (Me.Tag)
        Exit Sub
    ElseIf intTyped <> 10 Then
        intTyped = (intTyped + 1)
        Exit Sub
    ElseIf intTyped = 10 Then
        intTyped = 0
        Exit Sub
    End If
    
  

End Sub





