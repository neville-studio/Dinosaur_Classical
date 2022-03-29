VERSION 5.00
Begin VB.Form OptionDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "设置"
   ClientHeight    =   2565
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6030
   Icon            =   "OptionDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   171
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   402
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.CheckBox jyrdCheck1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "随机交易"
      Height          =   180
      Left            =   2040
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.VScrollBar jydjsetscr 
      Height          =   255
      Left            =   1680
      Max             =   9
      Min             =   1
      TabIndex        =   4
      Top             =   600
      Value           =   8
      Width           =   255
   End
   Begin VB.TextBox jydjsettings 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Text            =   "2"
      Top             =   600
      Width           =   615
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "取消"
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "确定"
      Height          =   375
      Left            =   4680
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "更多内容，敬请期待。"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1560
      TabIndex        =   5
      Top             =   1440
      Width           =   2100
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "交易等级"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   2
      Top             =   600
      Width           =   840
   End
End
Attribute VB_Name = "OptionDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Sub CancelButton_Click()
    Call Form_Unload(0)
End Sub

Private Sub jyrdCheck1_Click()
    jydjsettings.Enabled = Not (jyrdCheck1.Value * True)
    jydjsetscr.Enabled = Not (jyrdCheck1.Value * True)
    
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub
Private Sub Form_Load()
    Dim filesys, file, textfile
    Dim settingtexts As String, jiaoyidengji As Integer
    Set filesys = CreateObject("scripting.filesystemobject")
    If filesys.fileexists(App.Path + "\Data\dinoprof.set") Then
        Set file = filesys.getfile(App.Path + "\Data\dinoprof.set")
        Set textfile = file.openastextstream(1, -2)
        settingtexts = textfile.readline
        jydjsettings.Text = Val(Mid(settingtexts, 14))
        If Val(jydjsettings.Text) > 10 Then
            jydjsettings.Text = 10
        ElseIf Val(jydjsettings.Text) < 1 Then
            jydjsettings.Text = 1
        End If
        jyrdCheck1.Value = Val(Mid(textfile.readline, 14))
        Call jyrdCheck1_Click
        textfile.Close
    Else
        jydjsettings.Text = 3
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload OptionDialog
    MainForm.Enabled = True
End Sub

Private Sub jydjsettings_Change()
    If Val(jydjsettings.Text) > 9 Then
        jydjsettings.Text = "9"
    ElseIf Val(jydjsettings.Text) <= 1 Then
        jydjsettings.Text = "1"
    End If
    jydjsetscr.Value = 10 - Val(jydjsettings.Text)
End Sub

Private Sub jydjsettings_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 38 Then
        jydjsettings.Text = CStr(Val(jydjsettings.Text) + 1)
    ElseIf KeyCode = 40 Then
        jydjsettings.Text = CStr(Val(jydjsettings.Text) - 1)
    ElseIf Not (KeyCode >= vbKey0 And KeyCode <= vbKey9 Or KeyCode >= vbKeyNumpad0 And KeyCode <= vbKeyNumpad9) Or KeyCode = vbKeyDelete Or KeyCode = vbKeyBack Then
        KeyCode = 3
    End If
    Call jydjsettings_Change
    jydjsetscr.Value = 10 - Val(jydjsettings.Text)
End Sub

Private Sub jydjsetscr_Change()
    jydjsettings.Text = CStr(10 - jydjsetscr.Value)
End Sub

Private Sub OKButton_Click()
    Dim file, filesys
    Set filesys = CreateObject("scripting.filesystemobject")
    Set file = filesys.createtextfile(App.Path + "\data\dinoprof.set")
    file.writeline ("jiaoyidengji=" + CStr(jydjsettings.Text))
    file.writeline ("jiaoyirandom=" + CStr(jyrdCheck1.Value))
    file.Close
    MainForm.Enabled = True
    If jyrdCheck1.Value = 1 Then
        MainForm.jiaoyidengji = -1
        
    Else
        MainForm.jiaoyidengji = jydjsettings.Text
    End If
    Unload OptionDialog
End Sub
