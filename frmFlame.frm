VERSION 5.00
Begin VB.Form frmFlame 
   Appearance      =   0  'Flat
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   10800
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13665
   LinkTopic       =   "Form1"
   ScaleHeight     =   720
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   911
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hscrAdj 
      Height          =   255
      Index           =   3
      Left            =   6240
      Max             =   100
      TabIndex        =   10
      Top             =   4710
      Value           =   74
      Width           =   1935
   End
   Begin VB.HScrollBar hscrAdj 
      Height          =   255
      Index           =   2
      Left            =   4200
      Max             =   255
      Min             =   140
      TabIndex        =   8
      Top             =   4710
      Value           =   255
      Width           =   1935
   End
   Begin VB.HScrollBar hscrAdj 
      Height          =   255
      Index           =   1
      Left            =   2160
      Max             =   255
      TabIndex        =   6
      Top             =   4710
      Value           =   20
      Width           =   1935
   End
   Begin VB.HScrollBar hscrAdj 
      Height          =   255
      Index           =   0
      Left            =   120
      Max             =   100
      Min             =   15
      TabIndex        =   4
      Top             =   4710
      Value           =   15
      Width           =   1935
   End
   Begin VB.Timer tmrAni 
      Interval        =   1
      Left            =   480
      Top             =   6120
   End
   Begin VB.PictureBox picDecMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   3240
      Picture         =   "frmFlame.frx":0000
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   3
      Top             =   5280
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.PictureBox picDecMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   1080
      Picture         =   "frmFlame.frx":83D8C
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   139
      TabIndex        =   2
      Top             =   5280
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.PictureBox picFlame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   4335
      Left            =   0
      ScaleHeight     =   289
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   601
      TabIndex        =   1
      Top             =   0
      Width           =   9015
   End
   Begin VB.Timer tmrFlame 
      Interval        =   1
      Left            =   480
      Top             =   5520
   End
   Begin VB.PictureBox picPAL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   135
      Left            =   120
      Picture         =   "frmFlame.frx":107B18
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   0
      Top             =   5040
      Visible         =   0   'False
      Width           =   5145
   End
   Begin VB.Label lblDivider 
      BackStyle       =   0  'Transparent
      Caption         =   "Flame Divider"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   4485
      Width           =   1935
   End
   Begin VB.Label lblMax 
      BackStyle       =   0  'Transparent
      Caption         =   "Maximum Brightness"
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   4485
      Width           =   1935
   End
   Begin VB.Label lblMin 
      BackStyle       =   0  'Transparent
      Caption         =   "Minimum Brightness"
      Height          =   255
      Left            =   2160
      TabIndex        =   7
      Top             =   4485
      Width           =   1935
   End
   Begin VB.Label lblScatter 
      BackStyle       =   0  'Transparent
      Caption         =   "Scatter Factor"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4485
      Width           =   2895
   End
End
Attribute VB_Name = "frmFlame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Plot(320, 200) As Integer, BlkSize As Integer, XOff As Long, YOff As Long, Angle(3) As Single
Dim StopRandFlame As Boolean, ScatterFactor As Single, MinBrightnessFactor As Integer, MaxBrightnessFactor As Integer, FlameDivider As Single

Private Sub Form_Load()
    Dim XPos As Integer, YPos As Integer
    picPAL.Width = 255 * 5
    picPAL.Height = 10
    
    ScatterFactor = 1.5         'From 1 to 10
    MinBrightnessFactor = 20    'From 0 to 255
    MaxBrightnessFactor = 255   'From 140 to 255
    FlameDivider = 0.0074
            
    BlkSize = 3
    picFlame.Width = 200 * BlkSize + 3 - 40
    picFlame.Height = 100 * BlkSize - 2
    Me.Width = picFlame.Width * 15
    Me.Height = picFlame.Height * 15 + 550
    Me.Top = Screen.Height / 2 - Me.Height / 2
    Me.Left = Screen.Width / 2 - Me.Width / 2
    picFlame.ForeColor = vbWhite
    
    picDecMain.Width = 603
    picDecMask.Width = 603
    picDecMain.Height = 298
    picDecMask.Height = 298
    
    Call InitFlame
    StopRandFlame = True
End Sub

Private Sub InitFlame()
    Dim XPos As Integer, YPos As Integer
    For YPos = 0 To 200
        For XPos = 0 To 320
            Plot(XPos, YPos) = 0
        Next XPos
    Next YPos
End Sub

Private Sub DrawFlame(StopGen As Boolean)
    Dim XPos As Integer, YPos As Integer, Sum As Integer, Ind As Integer, Modder As Integer
    Randomize Timer
    For YPos = 100 To 98 Step -1
        For XPos = 0 To 200
            If (StopGen = True) Then Sum = Int(Rnd * 300) + 1 Else Sum = 0
            If (Sum > MaxBrightnessFactor) Then Sum = MaxBrightnessFactor
            If (Sum < MinBrightnessFactor) Then Sum = MinBrightnessFactor
            Plot(XPos, YPos) = Sum
        Next XPos
    Next YPos
    For YPos = 98 To 1 Step -1
        For XPos = 1 To 199
            Sum = Abs(Plot(XPos - 1, YPos + 1) + Plot(XPos, YPos + 1) + Plot(XPos + 1, YPos + 1) + Plot(XPos, YPos))
            Sum = Sum / (4 + FlameDivider)
            If ((Int(Rnd * 2) + 1) > 1) Then Modder = -1 Else Modder = 1
            Plot(XPos, YPos) = Sum + Int(Rnd * ScatterFactor) * Modder
            If (YPos <= 97) Then BitBlt picFlame.hdc, XPos * BlkSize - 20, YPos * BlkSize, BlkSize, BlkSize, picPAL.hdc, Plot(XPos, YPos) * 4, 1, SRCCOPY
        Next XPos
    Next YPos
    If (StopGen = False) Then
        picFlame.FontSize = 20: picFlame.CurrentX = 130: picFlame.CurrentY = 20
        picFlame.Print "Thanks for viewing..."
    Else
        picFlame.FontSize = 16: picFlame.CurrentX = 130: picFlame.CurrentY = 120
        picFlame.Print "Visual Basic 6.0  -  Use of BitBlt API"
        picFlame.FontSize = 20: picFlame.CurrentX = 130: picFlame.CurrentY = 140
        picFlame.Print "Blazing Flames Demonstration"
        picFlame.FontSize = 16: picFlame.CurrentX = 130: picFlame.CurrentY = 168
        picFlame.Print "Programmed by Tan Hong Hwee"
    End If
    picFlame.Refresh
End Sub

Private Sub hscrAdj_Change(Index As Integer)
    Select Case Index
    Case 0: ScatterFactor = hscrAdj(0).Value / 10
    Case 1: MinBrightnessFactor = hscrAdj(1).Value
    Case 2: MaxBrightnessFactor = hscrAdj(2).Value
    Case 3: FlameDivider = hscrAdj(3).Value / 10000
    End Select
End Sub

Private Sub hscrAdj_Scroll(Index As Integer)
    Call hscrAdj_Change(Index)
End Sub

Private Sub picFlame_Click()
    StopRandFlame = False
End Sub

Private Sub tmrFlame_Timer()
    'This is the animation timer
    Static Count As Integer
    If (StopRandFlame = False) Then Count = Count + 1
    Call DrawFlame(StopRandFlame)
    If (Count = 25) Then End
End Sub

