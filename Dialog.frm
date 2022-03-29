VERSION 5.00
Begin VB.Form HighScoreDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dinosaur - 高分榜"
   ClientHeight    =   3195
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.TextBox Namereq 
      CausesValidation=   0   'False
      Height          =   270
      Left            =   960
      TabIndex        =   16
      Text            =   "无名小子"
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   18
      Top             =   2760
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "重置所有记录"
      Height          =   375
      Left            =   2880
      TabIndex        =   19
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label HighScoreName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   4
      Left            =   4080
      TabIndex        =   17
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label HighScoreName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   3
      Left            =   4200
      TabIndex        =   15
      Top             =   1920
      Width           =   90
   End
   Begin VB.Label HighScoreName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   2
      Left            =   4200
      TabIndex        =   14
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label HighScoreName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   1
      Left            =   4200
      TabIndex        =   13
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label HighScoreName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   0
      Left            =   4080
      TabIndex        =   12
      Top             =   600
      Width           =   90
   End
   Begin VB.Label HighScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   4
      Left            =   1440
      TabIndex        =   11
      Top             =   2160
      Width           =   90
   End
   Begin VB.Label HighScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   3
      Left            =   1440
      TabIndex        =   10
      Top             =   1800
      Width           =   90
   End
   Begin VB.Label HighScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   2
      Left            =   1440
      TabIndex        =   9
      Top             =   1440
      Width           =   90
   End
   Begin VB.Label HighScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   1
      Left            =   1440
      TabIndex        =   8
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label HighScore 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   180
      Index           =   0
      Left            =   1320
      TabIndex        =   7
      Top             =   600
      Width           =   90
   End
   Begin VB.Label rankl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第5名"
      Height          =   180
      Index           =   4
      Left            =   120
      TabIndex        =   6
      Top             =   2160
      Width           =   450
   End
   Begin VB.Label rankl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第4名"
      Height          =   180
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   450
   End
   Begin VB.Label rankl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第3名"
      Height          =   180
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   450
   End
   Begin VB.Label rankl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第2名"
      Height          =   180
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   450
   End
   Begin VB.Label rankl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "第1名"
      Height          =   180
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   450
   End
   Begin VB.Label namecap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "用户"
      Height          =   180
      Left            =   4080
      TabIndex        =   1
      Top             =   360
      Width           =   360
   End
   Begin VB.Label scorecap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "分数"
      Height          =   180
      Left            =   1440
      TabIndex        =   0
      Top             =   360
      Width           =   360
   End
End
Attribute VB_Name = "HighScoreDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim high(0 To 6) As Long
Dim names(0 To 6) As String
Dim t As Long
Dim n As Integer
Dim s As Integer
Dim havedone As Boolean
Public change As Long
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub CancelButton_Click()
    If Not havedone And s >= 0 Then
        HighScoreName(s).Caption = Namereq.Text
        Namereq.Visible = False
        names(s) = Namereq.Text
        Namereq.Text = "无名小子"
        HighScoreName(s).Left = namecap.Left + namecap.Width / 2 - HighScoreName(s).Width / 2
        havedone = True
    End If
    Call savehighscore
    MainForm.Enabled = True
    Unload Me
End Sub
Private Sub savehighscore()
Dim filesys, file, textfile
    Set filesys = CreateObject("scripting.filesystemobject")
    Set file = filesys.createtextfile(App.Path + "\Data\highscore.hgs")
    file.writeline (Now())
    For i = 0 To 4
        file.writeline (HighScore(i).Caption)
        file.writeline (HighScoreName(i).Caption)
    Next i
End Sub
Private Sub Form_Initialize()
InitCommonControls
End Sub

Private Sub Form_Load()
    Dim filesys, file, textfile
    Set filesys = CreateObject("scripting.filesystemobject")
    Me.Caption = "Dinosaur - 高分榜"
    n = 0
    s = -1
    havedone = False
    If filesys.fileexists(App.Path + "\Data\highscore.hgs") Then
        Set file = filesys.getfile(App.Path + "\Data\highscore.hgs")
        Set textfile = file.openastextstream(1, -2)
        textfile.readline
        For i = 0 To 4
            If Not textfile.atendofline Then
                n = n + 1
                high(i) = CLng(textfile.readline)
                If high(i) = 0 Then
                    n = n - 1
                    Exit For
                End If
                names(i) = textfile.readline
                For j = 1 To i
                    If high(j) > high(j - 1) Then t = high(j): high(j) = high(j - 1): high(j - 1) = t: ts = names(j): names(j) = names(j - 1): names(j - 1) = ts
                Next j
            End If
        Next i
    End If
    If change <> 0 Then
        s = 1
        For i = n To 0 Step -1
            If high(i) < change Then
                high(i + 1) = high(i): names(i + 1) = names(i)
                s = i
            Else
                Exit For
            End If
        Next i
        high(s) = change
        Namereq.Top = rankl(s).Top
        Me.Caption = "Dinosaur - 尊姓大名？"
        Namereq.Visible = True
        Namereq.Left = namecap.Left + namecap.Width / 2 - Namereq.Width / 2
        n = n + 1
        If n > 5 Then n = 5
        change = 0
    End If
output:
    For i = 0 To n - 1
        HighScore(i).Caption = CStr(high(i))
        HighScoreName(i).Caption = names(i)
        HighScoreName(i).Top = rankl(i).Top
        HighScore(i).Left = scorecap.Left + scorecap.Width / 2 - HighScore(i).Width / 2
        HighScoreName(i).Left = namecap.Left + namecap.Width / 2 - HighScoreName(i).Width / 2
        HighScore(i).Top = rankl(i).Top
    Next i
    MainForm.highscorel.Caption = "HI    " + pluszero(CStr(high(0)))

End Sub
Function pluszero(x As String) As String
    Do While Len(x) < 5
        x = "0" + x
    Loop
    pluszero = x
End Function

Private Sub Form_Unload(Cancel As Integer)
    MainForm.showhighdialog = high(4)
    Call CancelButton_Click
    MainForm.Enabled = True
End Sub

Private Sub Namereq_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
End If
End Sub

Private Sub Namereq_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Not havedone And s >= 0 Then
            HighScoreName(s).Caption = Namereq.Text
            Namereq.Visible = False
            names(s) = Namereq.Text
            Namereq.Text = "无名小子"
            HighScoreName(s).Left = namecap.Left + namecap.Width / 2 - HighScoreName(s).Width / 2
            havedone = True
        End If
    End If
End Sub

Private Sub OKButton_Click()
    Dim right
    right = MsgBox("是否重置记录？", vbYesNo, "Dinosaur")
    If right = vbYes Then
        For i = 0 To 4
            Namereq.Visible = False
            Namereq.Text = ""
            HighScore(i).Caption = ""
            HighScoreName(i).Caption = ""
            high(i) = 0
            names(i) = ""
            high(5) = 0: high(6) = 0
        Next i
        MainForm.highscorel.Caption = "HI    00000"
    End If
End Sub
