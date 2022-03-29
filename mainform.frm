VERSION 5.00
Begin VB.Form MainForm 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dinosaur 2.0"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   735
   ClientWidth     =   13455
   Icon            =   "mainform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   0  'User
   ScaleTop        =   -256
   ScaleWidth      =   897
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer wait 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   480
      Top             =   480
   End
   Begin VB.Timer bossgif 
      Interval        =   500
      Left            =   2880
      Top             =   480
   End
   Begin VB.PictureBox bosspic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   10800
      Picture         =   "mainform.frx":0442
      ScaleHeight     =   870
      ScaleWidth      =   1305
      TabIndex        =   12
      Top             =   2400
      Visible         =   0   'False
      Width           =   1305
   End
   Begin VB.Timer bossnightchufa 
      Interval        =   1000
      Left            =   2880
      Top             =   0
   End
   Begin VB.Timer thinggif 
      Interval        =   1900
      Left            =   2400
      Top             =   480
   End
   Begin VB.Timer businessmanchufa 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2400
      Top             =   0
   End
   Begin VB.Timer attack 
      Enabled         =   0   'False
      Interval        =   15
      Left            =   1920
      Top             =   480
   End
   Begin VB.PictureBox firmic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   840
      Picture         =   "mainform.frx":0746
      ScaleHeight     =   375
      ScaleWidth      =   600
      TabIndex        =   10
      Top             =   2760
      Visible         =   0   'False
      Width           =   600
   End
   Begin VB.PictureBox hit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   2
      Left            =   10560
      Picture         =   "mainform.frx":0858
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer dnchange 
      Interval        =   5
      Left            =   1920
      Top             =   0
   End
   Begin VB.PictureBox moon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   6960
      Picture         =   "mainform.frx":0908
      ScaleHeight     =   750
      ScaleWidth      =   750
      TabIndex        =   7
      Top             =   720
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.PictureBox hit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1425
      Index           =   3
      Left            =   11040
      Picture         =   "mainform.frx":26FA
      ScaleHeight     =   95
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   75
      TabIndex        =   4
      Top             =   2385
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.PictureBox hit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   9600
      Picture         =   "mainform.frx":2BB8
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   50
      TabIndex        =   3
      Top             =   2550
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.Timer gaming 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   480
   End
   Begin VB.Timer productrun 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1440
      Top             =   0
   End
   Begin VB.Timer product 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   960
      Top             =   0
   End
   Begin VB.PictureBox hit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   825
      Index           =   1
      Left            =   9720
      Picture         =   "mainform.frx":2C7A
      ScaleHeight     =   55
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   2985
      Visible         =   0   'False
      Width           =   675
   End
   Begin VB.Timer begintextc 
      Interval        =   250
      Left            =   480
      Top             =   0
   End
   Begin VB.Timer beginscene 
      Interval        =   1500
      Left            =   960
      Top             =   480
   End
   Begin VB.Timer runnerjump 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   0
      Top             =   480
   End
   Begin VB.Timer runnerrun 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox runnerpic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   60
      Picture         =   "mainform.frx":49F4
      ScaleHeight     =   1200
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   2610
      Width           =   750
   End
   Begin VB.Line bosshp 
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   288
      X2              =   616
      Y1              =   -224
      Y2              =   -224
   End
   Begin VB.Label highscorel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "HI "
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   4680
      TabIndex        =   13
      Top             =   0
      Width           =   450
   End
   Begin VB.Label dialog 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   3000
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.Label attacklabel 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Attack  0"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   240
      TabIndex        =   8
      Top             =   0
      Width           =   1485
   End
   Begin VB.Label scorel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SCORE  00000"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   10200
      TabIndex        =   6
      Top             =   0
      Width           =   1980
   End
   Begin VB.Label gameover 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   13.5
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   6360
      TabIndex        =   5
      Top             =   1560
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.Label begintext 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press Space/W key to start."
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   12
         Charset         =   255
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   5040
      TabIndex        =   1
      Top             =   2040
      Width           =   4860
   End
   Begin VB.Line ground 
      BorderWidth     =   5
      X1              =   0
      X2              =   896
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Line Bosswait 
      BorderColor     =   &H00777777&
      BorderWidth     =   10
      Visible         =   0   'False
      X1              =   288
      X2              =   616
      Y1              =   -224
      Y2              =   -224
   End
   Begin VB.Menu MainFormf 
      Caption         =   "游戏(&G)"
      NegotiatePosition=   2  'Middle
      Begin VB.Menu mainformbus 
         Caption         =   "设置(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu record 
         Caption         =   "英雄榜(&R)"
      End
      Begin VB.Menu nosound 
         Caption         =   "静音(&P)"
         Shortcut        =   ^M
      End
      Begin VB.Menu onExitClick 
         Caption         =   "退出(&X)"
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu MainFormh 
      Caption         =   "帮助(&H)"
      NegotiatePosition=   2  'Middle
      WindowList      =   -1  'True
      Begin VB.Menu help 
         Caption         =   "帮助(&P)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu onAbout 
         Caption         =   "关于  Dinosaur 2.0(&A)"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim gif As Integer, buscount As Integer, bossgifw As Integer
Const gravity = 10: Const pi = 3.14159265358979
Dim runner(7) As Picture, businessman(1) As Picture, flyer(1) As Picture, boss(1) As Picture, tree(1) As Picture
Const adaylong = 400
Dim moonn(1) As Picture, attackpic(1) As Picture, bossattack As Picture
Dim day As Integer, action As Integer, liedown As Integer, showbusday As Integer, showbossday As Integer
Dim runnervy As Single, score As Long, amount As Integer
Dim difficulty As Integer
Public jiaoyidengji As Integer, bosshpnum As Integer
Dim time As Long, dday As Integer
Dim hitshow As Integer, chufa As Integer, attackfire As Integer, bosschufa As Integer
Public showhighdialog As Integer
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Private Declare Function sndPlaySoundStop Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As Long, ByVal uFlags As Long) As Long
Dim plays As Long
Private Sub attack_Timer()
    firmic.Picture = attackpic(day)
    firmic.Left = firmic.Left + 10
    If firmic.Left > MainForm.Width / 15 Then attack.Enabled = False: firmic.Visible = False
    If Abs(2 * hit(0).Left + hit(0).Width - 2 * firmic.Left - firmic.Width + 10) <= firmic.Width - 10 + hit(0).Width And Abs(2 * hit(0).Top + hit(0).Height - firmic.Top * 2 - firmic.Height) <= firmic.Height + hit(0).Height Then hit(0).Left = -hit(0).Width: firmic.Left = MainForm.Width / 15: attack.Enabled = False: firmic.Visible = False
    If Abs(2 * hit(1).Left + hit(1).Width - 2 * firmic.Left - firmic.Width) <= firmic.Width + hit(1).Width And Abs(2 * hit(1).Top + hit(1).Height - firmic.Top * 2 - firmic.Height) <= firmic.Height + hit(1).Height Then hit(1).Left = -hit(1).Width: firmic.Left = MainForm.Width / 15: attack.Enabled = False: firmic.Visible = False
    If Abs(2 * bosspic.Left + bosspic.Width - 2 * firmic.Left - firmic.Width) <= firmic.Width + bosspic.Width And Abs(2 * bosspic.Top + bosspic.Height - firmic.Top * 2 - firmic.Height) <= firmic.Height + bosspic.Height And bosspic.Visible Then
        firmic.Left = MainForm.Width / 15
        attack.Enabled = False
        firmic.Visible = False
        bosshp.X2 = bosshp.X2 - (Bosswait.X2 - Bosswait.X1) / bosshpnum
        If bosshp.X2 <= bosshp.X1 Then
            bosshp.Visible = False
            Bosswait.Visible = False
            bosspic.Visible = False
            score = score + 100 * bosshpnum
            bosshpnum = bosshpnum + 1
            If bosshpnum = 8 Then bosshpnum = 8
        End If
    End If
End Sub

Private Sub beginscene_Timer()
    beginscene.Interval = 1500 + Int((Rnd * 4) * 200) - 200
    If Not runnerjump.Enabled Then
        runnerjump.Enabled = True
        runnervy = 5
    End If
End Sub


Private Sub begintextc_Timer()
    begintext.Visible = Not begintext.Visible
End Sub

Private Sub bossgif_Timer()
    bossgifw = 1 - bossgifw
    bosspic.Picture = boss(bossgifw)
End Sub

Private Sub bossnightchufa_Timer()
    If showbossday < time \ adaylong Then
        showbossday = Int(Rnd * 3) + 2 + time \ adaylong
    ElseIf showbossday = time \ adaylong Then
        bosschufa = 1
    Else
        bosschufa = 0
    End If
End Sub
Private Sub businessmanchufa_Timer()
    If showbusday < time \ adaylong Then
        showbusday = Int(Rnd * 3) + 1 + time \ adaylong
    ElseIf showbusday = time \ adaylong And buscount = 0 Then
        chufa = 1
    End If
    If showbusday <> time \ adaylong Then buscount = 0
End Sub

Private Sub dnchange_Timer()
Dim jcolor As Long
jcolor = (CLng(256) * 256 + CLng(256) + 1) * 1
    If day = 1 And MainForm.BackColor <> vbBlack Then
        MainForm.BackColor = MainForm.BackColor - jcolor
    ElseIf time \ (adaylong \ 2) Mod 2 = 0 And MainForm.BackColor <> vbWhite Then
        MainForm.BackColor = MainForm.BackColor + jcolor
    Else
        scorel.ForeColor = vbWhite * day
        ground.BorderColor = vbWhite * day
        gameover.ForeColor = vbWhite * day
        begintext.ForeColor = vbWhite * day
        moon.Visible = True * day
        highscorel.ForeColor = vbWhite * day
        bosshp.BorderColor = vbWhite * day
        If Not bosspic.Visible Then
        hit(0).Picture = flyer(day)
        End If
        hit(1).Picture = tree(day)
        attacklabel.ForeColor = vbWhite * day
        dialog.BackColor = vbWhite * day
    End If
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer) '用户按键操作
    Select Case (KeyAscii)
        Case 32, 87, 119
            If Not begintextc.Enabled And Not dialog.Visible And Not begintextc.Enabled And Not gameover.Visible Then
                If Not (runnerjump.Enabled) And liedown = 0 Then
                    runnerjump.Enabled = True
                    runnervy = 5
                    runnerjump.Interval = 50 - difficulty
                End If
            ElseIf begintextc.Enabled Then
                bosshp.Visible = False
                Bosswait.Visible = False
                businessmanchufa.Enabled = True
                begintextc.Enabled = False
                begintext.Visible = False
                beginscene.Enabled = False
                product.Enabled = True
                gaming.Enabled = True
                record.Enabled = False
                For i = 0 To 3
                    hit(i).Left = MainForm.Width / 15
                Next i
                time = 0: score = 0
                MainForm.BackColor = vbWhite
                gameover.Visible = False
                liedown = 0: chufa = 0: showbusday = -1: showbossday = -1
                attackfire = 0: bosschufa = 0
                Call runnerjump_Timer
                vy = 0: runnerpic.Top = -82
                Call runnerrun_Timer
                firmic.Visible = False
                firmic.Left = MainForm.Width / 15
                attacklabel.Caption = "Attack  " + CStr(attackfire)
                wait.Enabled = False
                mainformbus.Enabled = False
            End If
        Case 83, 115
            If Not (runnerjump.Enabled Or gameover.Visible Or dialog.Visible) And liedown = 0 And Not begintextc.Enabled Then
                liedown = 1
                Call runnerrun_Timer
                runnerpic.Top = runnerpic.Top + 21
                attacklabel.Caption = "Attack  " + CStr(attackfire)
            End If
        Case 74, 106
            If (((productrun.Enabled And attackfire > 0 And Not firmic.Visible) Or (bosspic.Visible And attackfire > 0 And Not firmic.Visible))) And Not begintextc.Enabled Then
                firmic.Left = runnerpic.Left + runnerpic.Width
                firmic.Top = runnerpic.Top
                firmic.Visible = True
                attack.Enabled = True
                attackfire = attackfire - 1 'test!!!!!!!!!!!!
                attacklabel.Caption = "Attack  " + CStr(attackfire)
            End If
        Case 89, 121
            If dialog.Visible And score >= amount * 100 Then
                score = score - amount * 100
                attackfire = attackfire + amount * 3
                attacklabel.Caption = "Attack  " + CStr(attackfire)
                product.Enabled = True: productrun.Enabled = True: gaming.Enabled = True
                If runnerpic.Top < -82 Then runnerjump.Enabled = True
                attack.Enabled = True: chufa = 0
                dialog.Visible = False
                businessmanchufa.Enabled = True
                hit(hitshow).Left = -hit(hitshow).Width
            End If
        Case 78, 110
            If dialog.Visible Then
                product.Enabled = True: productrun.Enabled = True: gaming.Enabled = True
                attack.Enabled = True
                businessmanchufa.Enabled = True
                If runnerpic.Top < -82 Then runnerjump.Enabled = True
                dialog.Visible = False: chufa = 0
                hit(hitshow).Left = -hit(hitshow).Width
            End If
    End Select
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case (KeyCode)
        Case 83, 115
            If Not (runnerjump.Enabled Or gameover.Visible Or dialog.Visible) Then
                liedown = 0
                runnerpic.Top = -82
                Call runnerrun_Timer
            End If
     End Select
End Sub

Private Sub Form_Load() '加载游戏需要的图片并初始化相关设置项
    Dim settingtexts As String
    Dim filesys, file, textfile
    For i = 0 To 7
        Set runner(i) = LoadPicture(App.Path + "\images\runner" + CStr(i + 1) + ".bmp")
    Next i
    For i = 0 To 1
        Set flyer(i) = LoadPicture(App.Path + "\images\flyer" + CStr(i + 1) + ".bmp")
        Set boss(i) = LoadPicture(App.Path + "\images\boss" + CStr(i + 1) + ".bmp")
        Set businessman(i) = LoadPicture(App.Path + "\images\businessman" + CStr(i + 1) + ".bmp")
        Set tree(i) = LoadPicture(App.Path + "\images\tree" + CStr(i + 1) + ".bmp")
        Set attackpic(i) = LoadPicture(App.Path + "\images\userattack" + CStr(i + 1) + ".bmp")
    Next i
    Set moonn(0) = LoadPicture(App.Path + "\images\redmoon.bmp")
    Set moonn(1) = LoadPicture(App.Path + "\images\moon.bmp")
    Set bossattack = LoadPicture(App.Path + "\images\bossattack.bmp")
    Set filesys = CreateObject("scripting.filesystemobject")
    If filesys.fileexists(App.Path + "\Data\highscore.hgs") Then
        Load HighScoreDialog
        Unload HighScoreDialog
    Else
        highscorel.Caption = "HI    00000"
    End If
        If filesys.fileexists(App.Path + "\Data\dinoprof.set") Then
            Set file = filesys.getfile(App.Path + "\Data\dinoprof.set")
            Set textfile = file.openastextstream(1, -2)
            settingtexts = textfile.readline
            jiaoyidengji = Val(Mid(settingtexts, 14))
            If jiaoyidengji > 9 Then
                jiaoyidengji = 9
            ElseIf jiaoyidengji < 1 Then
                jiaoyidengji = 1
            End If
            settingtexts = textfile.readline
            If Val(Mid(settingtexts, 14)) = 1 Then
                jiaoyidengji = -1
            End If
            textfile.Close
        Else
            jiaoyidengji = 3
        End If
    plays = sndPlaySound(App.Path + "\sound\bgm.wav", 11)
    bosshpnum = 3
    Call beginscene_Timer
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim exit1 As Integer
    Dim filesys, file, textfile
    If gaming.Enabled Then
        exit1 = MsgBox("游戏正在进行，是否结束游戏？", vbYesNo, "Dinosaur")
        If exit1 = vbNo Then Cancel = -1 Else Cancel = 0
    End If
    If Cancel = 0 Then
        plays = sndPlaySoundStop(0, 0)
        End
    End If
End Sub

Private Sub gaming_Timer()
    time = time + 1
    score = score + 1
    difficulty = time / 200 + 5
    day = time \ (adaylong \ 2) Mod 2
    If showbossday - time / adaylong <= 1 And showbossday - time / adaylong >= -0.5 Then
        bosshp.Visible = True: Bosswait.Visible = True: bosshp.X2 = (Bosswait.X2 - Bosswait.X1) * (adaylong * 1.5 - (showbossday * adaylong + adaylong * 0.5 - time)) / (adaylong * 1.5) + Bosswait.X1
    ElseIf showbossday - time / adaylong <= -1 Then
        bosshp.Visible = False
        Bosswait.Visible = False
    End If
    If day = 1 And bosschufa = 1 And bosshp.X1 < bosshp.X2 Then
        moon.Picture = moonn(0): bosspic.Visible = True
    ElseIf bosshp.X1 >= bosshp.X2 And day = 1 And bosschufa = 1 Then
        bosspic.Visible = False: moon.Picture = moonn(0)
    Else
        moon.Picture = moonn(1): bosspic.Visible = False
    End If
    If day = 0 Then moon.Visible = False: bosspic.Visible = False
    If difficulty > 22 Then difficulty = 22
    scorel.Caption = "SCORE  " + pluszero(CStr(score))
    runnerpic.Left = -(MainForm.Width / 60) * Cos(time / 1000) + (MainForm.Width / 60)
    moon.Left = -(MainForm.Width / 30) * Cos(2 * pi / adaylong * (time Mod adaylong)) + (MainForm.Width / 30)
End Sub
Private Sub runnerishit()
    Dim showgameover As Boolean
    Dim filesys, file, textfile
    Dim HighScore As Long
    If hitshow = 0 Or hitshow = 1 Then
        If Abs(2 * hit(hitshow).Left + hit(hitshow).Width - 2 * runnerpic.Left - 34) <= 34 + hit(hitshow).Width And Abs(2 * hit(hitshow).Top + hit(hitshow).Height - (runnerpic.Top + 55 - 20 * liedown) * 2 - 14) <= 14 + hit(hitshow).Height Then showgameover = True
        If Abs(2 * hit(hitshow).Left + hit(hitshow).Width - 2 * (runnerpic.Left + 5 + 10 * action) - 14) <= 14 + hit(hitshow).Width And Abs(2 * hit(hitshow).Top + hit(hitshow).Height - (runnerpic.Top + 65 - 20 * liedown) * 2 - 14) <= 14 + hit(hitshow).Height Then showgameover = True
        If Abs(2 * hit(hitshow).Left + hit(hitshow).Width - 2 * (runnerpic.Left + 15) - (19 + 15 * liedown)) <= (19 + 15 * liedown) + hit(hitshow).Width And Abs(2 * hit(hitshow).Top + hit(hitshow).Height - runnerpic.Top * 2 - (55 - 19 * liedown)) <= 55 - 19 * liedown + hit(hitshow).Height Then showgameover = True 'need to done
        If Abs(2 * hit(hitshow).Left + hit(hitshow).Width - 2 * (runnerpic.Left + 30 + 19 * liedown) - (19 - 3 * liedown)) <= (19 - 3 * liedown) + hit(hitshow).Width And Abs(2 * hit(hitshow).Top + hit(hitshow).Height - (runnerpic.Top + 5 * liedown) * 2 - 19) <= 19 + hit(hitshow).Height Then showgameover = True
    End If
    If Abs(2 * hit(2).Left + hit(2).Width - 2 * runnerpic.Left - runnerpic.Width) <= runnerpic.Width + hit(2).Width And Abs(2 * hit(2).Top + hit(2).Height - runnerpic.Top * 2 - runnerpic.Height) <= runnerpic.Height + hit(2).Height Then attackfire = attackfire + 1: attacklabel.Caption = "Attack  " + CStr(attackfire): hit(hitshow).Left = 0 - hit(hitshow).Width
    If Abs(2 * hit(3).Left + hit(3).Width - 2 * runnerpic.Left - runnerpic.Width) <= runnerpic.Width + hit(3).Width And Abs(2 * hit(3).Top + hit(3).Height - runnerpic.Top * 2 - runnerpic.Height) <= runnerpic.Height + hit(3).Height Then Call showexchangedialog: chufa = False
    If showgameover Then
        mainformbus.Enabled = True
        gameover.Visible = True: wait.Enabled = True: product.Enabled = False: productrun.Enabled = False: gaming.Enabled = False
        runnerjump.Enabled = False: attack.Enabled = False
        record.Enabled = True
        If score > showhighdialog Then
            HighScoreDialog.change = score
            Load HighScoreDialog
            HighScoreDialog.Show
            Me.Enabled = False
        End If
    End If
End Sub
Private Sub showexchangedialog()
    If jiaoyidengji = -1 Then amount = Rnd * 10 + 1 Else amount = jiaoyidengji
    dialog.Caption = "Dialog:" + Chr(10) + "Hey,I'm businessman, would you like" + Chr(10) + "to change " + CStr(amount * 100) + " Score into " + CStr(amount * 3) + " fires?" + Chr(10) + "Yes(Y)         No(N)"
    product.Enabled = False: productrun.Enabled = False: gaming.Enabled = False
    runnerjump.Enabled = False: attack.Enabled = False
    dialog.Visible = True
End Sub
Function pluszero(j As String) As String
    Do While Len(j) < 5
        j = "0" + j
    Loop
    pluszero = j
End Function
Private Sub help_Click()
    Shell "explorer.exe " + App.Path + "\help\Dinosaur Help Document.htm", vbNormalFocus
End Sub

Private Sub mainformbus_Click()
    OptionDialog.Show
    MainForm.Enabled = False
End Sub

Private Sub nosound_Click()
    nosound.Checked = Not nosound.Checked
    If nosound.Checked Then
        plays = sndPlaySoundStop(0, 0)
        
    Else
        plays = sndPlaySound(App.Path + "\sound\bgm.wav", 11)
        
    End If
End Sub

Private Sub onAbout_Click()
    Load AboutDialog
    AboutDialog.Show
    Me.Enabled = False
End Sub
Private Sub onExitClick_Click()
    End
End Sub

Private Sub product_Timer()
    If chufa = 1 And day = 0 And buscount = 0 Then
        hitshow = 3 * Int(Rnd * 2)
    End If
    If hitshow <> 3 Then hitshow = Int(Rnd * 2) + Int(Rnd * 2) * day Else buscount = 1
    If bosspic.Visible And hitshow = 0 Then
        hit(0).Top = -82: hit(0).Left = bosspic.Left - hit(0).Width: hit(0).Picture = bossattack
    ElseIf hitshow = 0 And Not bosspic.Visible Then
        hit(0).Picture = flyer(day)
    ElseIf hitshow = 1 Then
        hit(1).Picture = tree(day)
    ElseIf hitshow = 3 Then
        hit(3).Picture = businessman(0)
    End If
    productrun.Enabled = True
    hit(hitshow).Visible = True
    product.Enabled = False
End Sub
Private Sub productrun_Timer()
    hit(hitshow).Left = hit(hitshow).Left - difficulty
    If hit(hitshow).Left <= 0 - hit(hitshow).Width Then
        hit(hitshow).Left = hit(hitshow).Width + MainForm.Width / 15
        hit(hitshow).Visible = False
        product.Enabled = True
        productrun.Enabled = False
        product.Interval = 1000 - difficulty * Int(Rnd * 10) + 100
        hitshow = -1
    End If
    Call runnerishit
End Sub
Private Sub record_Click()
    MainForm.Enabled = False
    HighScoreDialog.Show
End Sub
Private Sub runnerjump_Timer()
    runnerpic.Top = runnerpic.Top - (runnervy * 0.05 - 1 / 2 * gravity * 0.05 ^ 2) * 64
    runnervy = runnervy - (0.05 * gravity)
    If runnerpic.Top >= -82 Then runnerpic.Top = -82: vy = 0: runnerjump.Enabled = False
End Sub
Private Sub runnerrun_Timer()
    Dim isrun As Integer
    action = (action + 1) Mod 2
    isrun = action + day * 4 + liedown * 2
    Set runnerpic.Picture = runner(isrun)
End Sub
Private Sub thinggif_Timer()
    gif = 1 - gif
    hit(3).Picture = businessman(gif)
    thinggif.Interval = 2000 - thinggif.Interval
End Sub

Private Sub wait_Timer()
    begintextc.Enabled = True
    bosshpnum = 3
End Sub
