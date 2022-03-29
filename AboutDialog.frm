VERSION 5.00
Begin VB.Form AboutDialog 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dinosaur - 关于 Dinosaur 2.0"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   Icon            =   "AboutDialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton okCommand 
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   2400
      Width           =   2655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "灵感来自chrome://dino"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Copyrightlabel 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright"
      Height          =   420
      Left            =   1200
      TabIndex        =   2
      Top             =   960
      Width           =   2940
   End
   Begin VB.Label appversionlabel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   180
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   630
   End
   Begin VB.Image Image1 
      Height          =   1200
      Left            =   240
      Picture         =   "AboutDialog.frx":0442
      Top             =   600
      Width           =   750
   End
   Begin VB.Label APPNamel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Dinosaur 2.0"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1200
      TabIndex        =   0
      Top             =   480
      Width           =   1080
   End
End
Attribute VB_Name = "AboutDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Sub Command1_Click()
    
End Sub

Private Sub Form_Load()
    APPNamel.Caption = App.EXEName
    appversionlabel.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    Copyrightlabel.Caption = App.LegalCopyright
End Sub
Private Sub Form_Initialize()
    InitCommonControls
End Sub
Private Sub Form_Unload(Cancel As Integer)

    MainForm.Enabled = True
        Unload Me
End Sub

Private Sub okCommand_Click()
    Call Form_Unload(0)
End Sub
