VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Os - ill  2000.DmkWare   (use keys z-/ & sd ghj l;)"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7095
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command6 
      Caption         =   "Save Patch"
      Height          =   375
      Left            =   5940
      TabIndex        =   57
      Top             =   4080
      Width           =   1050
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Load Patch"
      Height          =   375
      Left            =   4920
      TabIndex        =   56
      Top             =   4080
      Width           =   1035
   End
   Begin OsIll.Knob Knob20 
      Height          =   615
      Left            =   6240
      TabIndex        =   55
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob19 
      Height          =   615
      Left            =   4200
      TabIndex        =   54
      Top             =   3600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob18 
      Height          =   615
      Left            =   3600
      TabIndex        =   53
      Top             =   3600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob17 
      Height          =   615
      Left            =   3000
      TabIndex        =   52
      Top             =   3600
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob16 
      Height          =   615
      Left            =   5520
      TabIndex        =   51
      Top             =   2400
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob15 
      Height          =   375
      Left            =   2400
      TabIndex        =   50
      Top             =   2160
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OsIll.Knob Knob14 
      Height          =   375
      Left            =   2400
      TabIndex        =   49
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OsIll.Knob Knob13 
      Height          =   855
      Left            =   2760
      TabIndex        =   48
      Top             =   2160
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob12 
      Height          =   855
      Left            =   1680
      TabIndex        =   47
      Top             =   2160
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob11 
      Height          =   855
      Left            =   1680
      TabIndex        =   46
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob10 
      Height          =   855
      Left            =   180
      TabIndex        =   45
      Top             =   3540
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob4 
      Height          =   375
      Left            =   3480
      TabIndex        =   39
      Top             =   600
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OsIll.Knob Knob9 
      Height          =   855
      Left            =   2760
      TabIndex        =   44
      Top             =   600
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob8 
      Height          =   855
      Left            =   1020
      TabIndex        =   43
      Top             =   3540
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob7 
      Height          =   615
      Left            =   6240
      TabIndex        =   42
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob6 
      Height          =   615
      Left            =   5520
      TabIndex        =   41
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1085
   End
   Begin OsIll.Knob Knob5 
      Height          =   375
      Left            =   3480
      TabIndex        =   40
      Top             =   2160
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
   End
   Begin OsIll.Knob Knob3 
      Height          =   855
      Left            =   3840
      TabIndex        =   38
      Top             =   2160
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob2 
      Height          =   855
      Left            =   3840
      TabIndex        =   37
      Top             =   600
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   1508
   End
   Begin OsIll.Knob Knob1 
      Height          =   855
      Left            =   1920
      TabIndex        =   36
      Top             =   3540
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   1508
   End
   Begin VB.PictureBox WF1PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   975
      Left            =   240
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   140.367
      TabIndex        =   11
      Top             =   480
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog CDB 
      Left            =   6720
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Load UDE"
      Height          =   255
      Left            =   5040
      TabIndex        =   35
      Top             =   3600
      Width           =   1815
   End
   Begin VB.CheckBox WCE 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   4960
      TabIndex        =   33
      Top             =   3260
      Width           =   255
   End
   Begin VB.ComboBox FX2 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   5040
      List            =   "Form1.frx":0037
      TabIndex        =   29
      Text            =   "None"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "RND"
      Height          =   615
      Left            =   5040
      TabIndex        =   28
      Top             =   2400
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "RND"
      Height          =   615
      Left            =   5040
      TabIndex        =   27
      Top             =   840
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "R"
      Height          =   195
      Left            =   1560
      TabIndex        =   22
      Top             =   3360
      Width           =   200
   End
   Begin VB.PictureBox WF2PB 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      Height          =   975
      Left            =   240
      ScaleHeight     =   255
      ScaleMode       =   0  'User
      ScaleWidth      =   140.367
      TabIndex        =   16
      Top             =   2040
      Width           =   1335
   End
   Begin VB.ComboBox FX1 
      Height          =   315
      ItemData        =   "Form1.frx":012C
      Left            =   5040
      List            =   "Form1.frx":0163
      TabIndex        =   7
      Text            =   "None"
      Top             =   480
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   6720
      Top             =   0
   End
   Begin VB.CheckBox EC 
      BackColor       =   &H00000000&
      Height          =   255
      Left            =   2925
      TabIndex        =   2
      Top             =   3260
      Width           =   255
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   960
      X2              =   960
      Y1              =   3240
      Y2              =   4440
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   1800
      X2              =   1800
      Y1              =   3240
      Y2              =   4440
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "UD Envelope"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   21
      Left            =   4920
      TabIndex        =   34
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   735
      Index           =   8
      Left            =   4920
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Attack"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   22
      Left            =   3000
      TabIndex        =   3
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   20
      Left            =   6015
      TabIndex        =   32
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   19
      Left            =   5265
      TabIndex        =   31
      Top             =   2400
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Osc 2 FX"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   18
      Left            =   5040
      TabIndex        =   30
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Index           =   7
      Left            =   4920
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Index           =   5
      Left            =   120
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Index           =   4
      Left            =   120
      Top             =   1680
      Width           =   4710
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Index           =   1
      Left            =   120
      Top             =   120
      Width           =   4710
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   15
      Left            =   3480
      TabIndex        =   26
      Top             =   960
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "FT"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   13
      Left            =   3480
      TabIndex        =   25
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   10
      Left            =   2400
      TabIndex        =   24
      Top             =   960
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "H"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   9
      Left            =   2400
      TabIndex        =   23
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tune"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   7
      Left            =   1080
      TabIndex        =   21
      Top             =   3360
      Width           =   495
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Octave"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   17
      Left            =   2640
      TabIndex        =   20
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Waveform"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   8
      Left            =   1680
      TabIndex        =   19
      Top             =   1920
      Width           =   735
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Osc. 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   4
      Left            =   240
      TabIndex        =   18
      Top             =   1680
      Width           =   1335
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Amplitude"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   3840
      TabIndex        =   17
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Waveform"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   16
      Left            =   1680
      TabIndex        =   15
      Top             =   360
      Width           =   735
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Octave"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   14
      Left            =   2640
      TabIndex        =   14
      Top             =   360
      Width           =   975
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Osc. 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   13
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Amplitude"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   3840
      TabIndex        =   12
      Top             =   360
      Width           =   975
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   1215
      Index           =   2
      Left            =   2880
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Shape StaticS 
      BorderColor     =   &H00FFFFFF&
      Height          =   1455
      Index           =   3
      Left            =   4920
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Osc 1 FX"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   6
      Left            =   5040
      TabIndex        =   10
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "A"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   11
      Left            =   5270
      TabIndex        =   9
      Top             =   840
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "B"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   12
      Left            =   6020
      TabIndex        =   8
      Top             =   840
      Width           =   375
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "AHD Envelope"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   25
      Left            =   3000
      TabIndex        =   6
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Decay"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   24
      Left            =   4200
      TabIndex        =   5
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Hold"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Index           =   23
      Left            =   3675
      TabIndex        =   4
      Top             =   4200
      Width           =   345
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Porta"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   855
   End
   Begin VB.Label StaticLBL 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Volume"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   3
      Left            =   1920
      TabIndex        =   0
      Top             =   3360
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EA As Byte, EA1 As Single, EA2 As Single, F1 As Single, F2 As Single, WF1 As Byte, WF2 As Byte, MT As Long, Hfreq1 As Single, Hfreq2 As Single
Dim ATT As Byte, DK As Byte, SUS As Byte, SV As Byte
Dim TargetFreq As Long, CurFreq As Long, Porta As Byte

Dim nT(BSize) As Single, TPi(BSize) As Single, TTT(BSize) As Long

Sub BuildNTable()
For i = 0 To BSize
nT(i) = i * 0.01745329251994 'Time * (Pi / 180)
TPi(i) = i * pi 'Time * Pi
TTT(i) = i * 2 'Time times two
Next
End Sub

Private Sub Command1_Click()
Knob8.SetVal 50
MT = 0
End Sub

Private Sub Command2_Click()
Knob6.SetVal RndRange(0, 100)
Knob7.SetVal RndRange(0, 100)
End Sub

Private Sub Command3_Click()
Knob16.SetVal RndRange(0, 100)
Knob20.SetVal RndRange(0, 100)
End Sub

Private Sub Command4_Click()
On Error GoTo Canceled
CDB.Filter = "Os-ill Envelope Files|*.env"
CDB.ShowOpen
Open CDB.Filename For Binary As #1
Get #1, , eByte()
Close #1
Canceled:
End Sub

Private Sub Command5_Click()
On Error GoTo Canceled
CDB.Filter = "Os-ill Path Files|*.oip"
CDB.ShowOpen
LoadPatch CDB.Filename
Canceled:
End Sub

Private Sub Command6_Click()
On Error GoTo Canceled
CDB.Filter = "Os-ill Path Files|*.oip"
CDB.ShowSave
SavePatch CDB.Filename
Canceled:
End Sub

Private Sub EC_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
 WCE.Value = 0
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
''''Begin Keyboard
If GetKeyState(90) < 0 Then TargetFreq = 6540: ATT = 1: DK = 0: SUS = 0
If GetKeyState(83) < 0 Then TargetFreq = 6930: ATT = 1: DK = 0: SUS = 0
If GetKeyState(88) < 0 Then TargetFreq = 7340: ATT = 1: DK = 0: SUS = 0
If GetKeyState(68) < 0 Then TargetFreq = 7780: ATT = 1: DK = 0: SUS = 0
If GetKeyState(67) < 0 Then TargetFreq = 8240: ATT = 1: DK = 0: SUS = 0
If GetKeyState(86) < 0 Then TargetFreq = 8730: ATT = 1: DK = 0: SUS = 0
If GetKeyState(71) < 0 Then TargetFreq = 9250: ATT = 1: DK = 0: SUS = 0
If GetKeyState(66) < 0 Then TargetFreq = 9800: ATT = 1: DK = 0: SUS = 0
If GetKeyState(72) < 0 Then TargetFreq = 10385: ATT = 1: DK = 0: SUS = 0
If GetKeyState(78) < 0 Then TargetFreq = 11000: ATT = 1: DK = 0: SUS = 0
If GetKeyState(74) < 0 Then TargetFreq = 11650: ATT = 1: DK = 0: SUS = 0
If GetKeyState(77) < 0 Then TargetFreq = 12345: ATT = 1: DK = 0: SUS = 0
If GetKeyState(188) < 0 Then TargetFreq = 13080: ATT = 1: DK = 0: SUS = 0
If GetKeyState(76) < 0 Then TargetFreq = 13860: ATT = 1: DK = 0: SUS = 0
If GetKeyState(190) < 0 Then TargetFreq = 14685: ATT = 1: DK = 0: SUS = 0
If GetKeyState(186) < 0 Then TargetFreq = 15555: ATT = 1: DK = 0: SUS = 0
If GetKeyState(191) < 0 Then TargetFreq = 16480: ATT = 1: DK = 0: SUS = 0
''''End Keyboard
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
''''Begin Keyboard
If GetKeyState(90) < 0 Then TargetFreq = 6540: ATT = 1: DK = 0: SUS = 0
If GetKeyState(83) < 0 Then TargetFreq = 6930: ATT = 1: DK = 0: SUS = 0
If GetKeyState(88) < 0 Then TargetFreq = 7340: ATT = 1: DK = 0: SUS = 0
If GetKeyState(68) < 0 Then TargetFreq = 7780: ATT = 1: DK = 0: SUS = 0
If GetKeyState(67) < 0 Then TargetFreq = 8240: ATT = 1: DK = 0: SUS = 0
If GetKeyState(86) < 0 Then TargetFreq = 8730: ATT = 1: DK = 0: SUS = 0
If GetKeyState(71) < 0 Then TargetFreq = 9250: ATT = 1: DK = 0: SUS = 0
If GetKeyState(66) < 0 Then TargetFreq = 9800: ATT = 1: DK = 0: SUS = 0
If GetKeyState(72) < 0 Then TargetFreq = 10385: ATT = 1: DK = 0: SUS = 0
If GetKeyState(78) < 0 Then TargetFreq = 11000: ATT = 1: DK = 0: SUS = 0
If GetKeyState(74) < 0 Then TargetFreq = 11650: ATT = 1: DK = 0: SUS = 0
If GetKeyState(77) < 0 Then TargetFreq = 12345: ATT = 1: DK = 0: SUS = 0
If GetKeyState(188) < 0 Then TargetFreq = 13080: ATT = 1: DK = 0: SUS = 0
If GetKeyState(76) < 0 Then TargetFreq = 13860: ATT = 1: DK = 0: SUS = 0
If GetKeyState(190) < 0 Then TargetFreq = 14685: ATT = 1: DK = 0: SUS = 0
If GetKeyState(186) < 0 Then TargetFreq = 15555: ATT = 1: DK = 0: SUS = 0
If GetKeyState(191) < 0 Then TargetFreq = 16480: ATT = 1: DK = 0: SUS = 0
''''End Keyboard
End Sub

Private Sub Form_Load()
If Not Init_DX7(Me.Hwnd) Then End
BuildNTable

WF1PB.ScaleWidth = BSize
WF2PB.ScaleWidth = BSize

Knob1.SetMode 0
Knob2.SetMode 0
Knob3.SetMode 0
Knob9.SetMode 0
Knob10.SetMode 0
Knob11.SetMode 0
Knob12.SetMode 0
Knob13.SetMode 0
Knob17.SetMode 0
Knob19.SetMode 0

Knob1.SetVal 10
Knob2.SetVal 5
Knob3.SetVal 5
Knob17.SetVal 5
Knob19.SetVal 5

Knob8.SetVal 50

Knob14.SetVal 30
Knob15.SetVal 30

Knob1.SetBackColor 0
Knob2.SetBackColor 0
Knob3.SetBackColor 0
Knob4.SetBackColor 0
Knob5.SetBackColor 0
Knob6.SetBackColor 0
Knob7.SetBackColor 0
Knob8.SetBackColor 0
Knob9.SetBackColor 0
Knob10.SetBackColor 0
Knob11.SetBackColor 0
Knob12.SetBackColor 0
Knob13.SetBackColor 0
Knob14.SetBackColor 0
Knob15.SetBackColor 0
Knob16.SetBackColor 0
Knob17.SetBackColor 0
Knob18.SetBackColor 0
Knob19.SetBackColor 0
Knob20.SetBackColor 0

Knob9.SetStep 2
Knob13.SetStep 2

Hfreq1 = Knob14.KnobValue / 1000
Hfreq2 = Knob15.KnobValue / 1000

'default octave
F1 = 0.0223
F2 = 0.0223

'Default Porta/Freq settings
CurFreq = 6540
TargetFreq = 6540
Porta = 100
ATT = 1

SetVolume Knob1.KnobValue
DSB(0).Play DSBPLAY_LOOPING
DSB(1).Play DSBPLAY_LOOPING
End Sub

Private Sub Form_Unload(Cancel As Integer)
Term_DX7
End
End Sub

Private Sub FX1_Click()
EC.SetFocus
End Sub

Private Sub FX2_Click()
EC.SetFocus
End Sub

Private Sub Knob1_Changed()
SetVolume Knob1.KnobValue
End Sub

Private Sub Knob10_Changed()
Select Case Knob10.KnobValue
Case 0: Porta = 100
Case 1: Porta = 10
Case 2: Porta = 9
Case 3: Porta = 8
Case 4: Porta = 7
Case 5: Porta = 6
Case 6: Porta = 5
Case 7: Porta = 4
Case 8: Porta = 3
Case 9: Porta = 2
Case 10: Porta = 1
End Select
End Sub

Private Sub Knob13_Changed()
Select Case Knob13.KnobValue
Case 0: F2 = 0.0223 + (Knob5.KnobValue / 10000)
Case 2:  F2 = 0.0446 + (Knob5.KnobValue / 10000)
Case 4: F2 = 0.0669 + (Knob5.KnobValue / 10000)
Case 6:  F2 = 0.0892 + (Knob5.KnobValue / 10000)
Case 8:  F2 = 0.1114 + (Knob5.KnobValue / 10000)
Case 10: F2 = 0.1338 + (Knob5.KnobValue / 10000)
End Select
End Sub

Private Sub Knob14_Changed()
Hfreq1 = Knob14.KnobValue / 500
End Sub

Private Sub Knob15_Changed()
Hfreq2 = Knob15.KnobValue / 500
End Sub

Private Sub Knob21_Changed()
WCEDIV = Knob21.KnobValue
End Sub

Private Sub Knob4_Changed()
Knob9_Changed
End Sub

Private Sub Knob5_Changed()
Knob13_Changed
End Sub

Private Sub Knob8_Changed()
MT = (Knob8.KnobValue - 50) * 100
End Sub

Private Sub Knob9_Changed()
Select Case Knob9.KnobValue
Case 0: F1 = 0.0223 + (Knob4.KnobValue / 10000)
Case 2:  F1 = 0.0446 + (Knob4.KnobValue / 10000)
Case 4: F1 = 0.0669 + (Knob4.KnobValue / 10000)
Case 6:  F1 = 0.0892 + (Knob4.KnobValue / 10000)
Case 8:  F1 = 0.1114 + (Knob4.KnobValue / 10000)
Case 10: F1 = 0.1338 + (Knob4.KnobValue / 10000)
End Select
End Sub




Private Sub Timer1_Timer()
BenchStart = timeGetTime
On Error Resume Next 'Skip Overflows (if any clipping samples (< 0  or  > 255))

''''Clear waveform monitor
WF1PB.Cls: WF2PB.Cls

''''Begin Envelope 1
If EC.Value = 1 Then 'on/off
If ATT = 1 Then
If EA < 100 Then
If EA > 100 Then EA = 100
EA = EA + Knob17.KnobValue
Else
ATT = 0: SUS = 1: SV = 0
End If
End If
If SUS = 1 Then
If SV < Knob18.KnobValue Then
SV = SV + 1
Else
SUS = 0: DK = 1
End If
End If
If DK = 1 Then
If EA > 0 Then
EA = EA - Knob19.KnobValue
If EA < 0 Then EA = 0
Else
DK = 0
End If
End If
ElseIf WCE.Value = 1 Then
EPos = EPos + 1
If EPos > 399 Then EPos = 0
EA = eByte(EPos)
Else
EA = 100
End If
''''End Envelope

WF1 = Knob11.KnobValue
WF2 = Knob12.KnobValue
EA1 = Knob2.KnobValue / 10
EA2 = Knob3.KnobValue / 10

''''Begin Synthesis
For i = 0 To BSize
 
''''Begin Porta
For ii = 0 To 1
Select Case Porta
Case 100:  CurFreq = TargetFreq
Case Else
If DSB(ii).GetFrequency < TargetFreq Then
If CurFreq > TargetFreq Then CurFreq = TargetFreq
CurFreq = CurFreq + Porta
ElseIf DSB(ii).GetFrequency > TargetFreq Then
If CurFreq < TargetFreq Then CurFreq = TargetFreq
CurFreq = CurFreq - Porta
End If
End Select
Next
''''End Porta

Select Case WF1
Case 0: Osc1Samp = (EA1 * EA) * Sin(F1 * TPi(i))
Case 1: Osc1Samp = ((EA1 * EA) * Sin(F1 * TPi(i))) + RndRange(-10, 10)
Case 2: Osc1Samp = (EA1 * EA) * Abs(Sin(F1 * TPi(i)))
Case 3: Osc1Samp = ((EA1 * EA) * Sin(F1 * TPi(i))) * Cos(Hfreq1 * TTT(i))
Case 4: Osc1Samp = ((EA1 * EA) * Sin(F1 * TPi(i))) * Cos(Hfreq1 * TTT(i)) * Cos(Hfreq1 * TTT(i))
Case 5: Osc1Samp = ((EA1 * EA) * Sin(F1 * TPi(i))) + (5 * Sin(Hfreq1 * TTT(i))) + (10 * Sin(Hfreq1 * TTT(i))) + (5 * Sin(Hfreq1 * TTT(i)))
Case 6: Osc1Samp = (EA1 * EA) * Sin(F1 * i * nT(i))
Case 7: Osc1Samp = (EA1 * EA) * Cos(F1 * TPi(i))
Case 8: Osc1Samp = ((EA1 * EA) * Sin(F1 * TPi(i))) * Cos(Sin(F1 * TPi(i)))
Case 9: Osc1Samp = (EA1 * EA) * Int(Sin(2 * F1 * i))
Case 10: Osc1Samp = ((EA1 * EA) * Int(Sin(F1 * i))) + RndRange(-10, 10)
End Select

If Not EA2 = 0 Then
Select Case WF2
Case 0: Osc2Samp = (EA2 * EA) * Sin(F2 * TPi(i))
Case 1: Osc2Samp = ((EA2 * EA) * Sin(F2 * TPi(i))) + RndRange(-10, 10)
Case 2: Osc2Samp = (EA2 * EA) * Abs(Sin(F2 * TPi(i)))
Case 3: Osc2Samp = ((EA2 * EA) * Sin(F2 * TPi(i))) * Cos(Hfreq2 * TTT(i))
Case 4: Osc2Samp = ((EA2 * EA) * Sin(F2 * TPi(i))) * Cos(Hfreq2 * TTT(i)) * Cos(Hfreq2 * TTT(i))
Case 5: Osc2Samp = ((EA2 * EA) * Sin(F2 * TPi(i))) + (5 * Sin(Hfreq2 * TTT(i))) + (10 * Sin(Hfreq2 * TTT(i))) + (5 * Sin(Hfreq2 * TTT(i)))
Case 6: Osc2Samp = (EA2 * EA) * Sin(F2 * i * nT(i))
Case 7: Osc2Samp = (EA2 * EA) * Cos(F1 * TPi(i))
Case 8: Osc2Samp = ((EA2 * EA) * Sin(F2 * TPi(i))) * Cos(Sin(F2 * TPi(i)))
Case 9: Osc2Samp = (EA2 * EA) * Int(Sin(2 * F2 * i))
Case 10: Osc2Samp = ((EA2 * EA) * Int(Sin(F2 * i))) + RndRange(-10, 10)
End Select
End If

Select Case FX1.Text
Case "None"
Case "Split (A)": If Osc1Samp < 0 Then Osc1Samp = Osc1Samp - Knob6.KnobValue Else Osc1Samp = Osc1Samp + Knob6.KnobValue
Case "Crazed (A&B)": Osc1Samp = Osc1Samp + ((Cos(Knob7.KnobValue * i) + Sin(i)) * Knob6.KnobValue)
Case "FreakD1 (A&B)": Osc1Samp = Osc1Samp / (Knob6.KnobValue * Tan(nT(i) * Knob7.KnobValue))
Case "FreakD2 (A)": Osc1Samp = Osc1Samp / (Knob6.KnobValue * Tan(nT(i) * EA))
Case "Disto (A)": Osc1Samp = Osc1Samp / Int(Cos(2 * 3.14 * i * Knob6.KnobValue) * Atn(i * 4))
Case "Noise (A&B)": Osc1Samp = Osc1Samp + (0.1 * RndRange(-Knob6.KnobValue, Knob7.KnobValue))
Case "Harmonix1 (A)": Osc1Samp = Osc1Samp * Cos((Knob6.KnobValue / 1000) * TTT(i))
Case "Harmonix2 (A&B)": Osc1Samp = Osc1Samp * Cos((Knob6.KnobValue / 1000) * TTT(i)) * Cos((Knob7.KnobValue / 1000) * TTT(i))
Case "Xor Disto (A&B)": Osc1Samp = Osc1Samp Xor Knob6.KnobValue / Knob7.KnobValue
Case "AddVoice1 (A&B)": Osc1Samp = Osc1Samp + ((Knob7.KnobValue / 5) * Cos((Knob6.KnobValue / 100) * i))
Case "AddVoice2 (A&B)": Osc1Samp = Osc1Samp + (10 * Cos((Knob6.KnobValue / 100) * i)) + (10 * Cos((Knob7.KnobValue / 100) * i))
Case "AV-Half&Half (A&B)": If i < 89 Then Osc1Samp = Osc1Samp + (10 * Cos((Knob6.KnobValue / 100) * i)) Else Osc1Samp = Osc1Samp + (10 * Cos((Knob7.KnobValue / 100) * i))
Case "AHD Sweep1 (A)": Osc1Samp = Osc1Samp + ((Knob6.KnobValue / 5) * Sin(Osc1Samp))
Case "AHD Sweep2 (A)": Osc1Samp = Osc1Samp - ((Knob6.KnobValue / 5) * Cos(Osc1Samp))
Case "AHD Sweep3 (A&B)": Osc1Samp = Osc1Samp * Cos((Knob6.KnobValue / 1000) * (Knob7.KnobValue * Cos(Osc1Samp)))
Case "ABS": Osc1Samp = Abs(Osc1Samp)
End Select

Select Case FX2.Text
Case "None"
Case "Split (A)": If Osc2Samp < 0 Then Osc2Samp = Osc2Samp - Knob16.KnobValue Else Osc2Samp = Osc2Samp + Knob16.KnobValue
Case "Crazed (A&B)": Osc2Samp = Osc2Samp + ((Cos(Knob20.KnobValue * i) + Sin(i)) * Knob16.KnobValue)
Case "FreakD1 (A&B)": Osc2Samp = Osc2Samp / (Knob16.KnobValue * Tan(nT(i) * Knob20.KnobValue))
Case "FreakD2 (A)": Osc2Samp = Osc2Samp / (Knob16.KnobValue * Tan(nT(i) * EA))
Case "Disto (A)": Osc2Samp = Osc2Samp / Int(Cos(2 * 3.14 * i * Knob16.KnobValue) * Atn(i * 4))
Case "Noise (A&B)": Osc2Samp = Osc2Samp + (0.1 * RndRange(-Knob16.KnobValue, Knob20.KnobValue))
Case "Harmonix1 (A)": Osc2Samp = Osc2Samp * Cos((Knob16.KnobValue / 1000) * TTT(i))
Case "Harmonix2 (A&B)": Osc2Samp = Osc2Samp * Cos((Knob16.KnobValue / 1000) * TTT(i)) * Cos((Knob20.KnobValue / 1000) * TTT(i))
Case "Xor Disto (A&B)": Osc2Samp = Osc2Samp Xor Knob16.KnobValue / Knob20.KnobValue
Case "AddVoice1 (A&B)": Osc2Samp = Osc2Samp + ((Knob20.KnobValue / 5) * Cos((Knob16.KnobValue / 100) * i))
Case "AddVoice2 (A&B)": Osc2Samp = Osc2Samp + (10 * Cos((Knob16.KnobValue / 100) * i)) + (10 * Cos((Knob20.KnobValue / 100) * i))
Case "AV-Half&Half (A&B)": If i < 89 Then Osc2Samp = Osc2Samp + (10 * Cos((Knob16.KnobValue / 100) * i)) Else Osc2Samp = Osc2Samp + (10 * Cos((Knob20.KnobValue / 100) * i))
Case "AHD Sweep1 (A)": Osc2Samp = Osc2Samp + ((Knob16.KnobValue / 5) * Sin(Osc2Samp))
Case "AHD Sweep2 (A)": Osc2Samp = Osc2Samp - ((Knob16.KnobValue / 5) * Cos(Osc2Samp))
Case "AHD Sweep3 (A&B)": Osc2Samp = Osc2Samp * Cos((Knob16.KnobValue / 1000) * (Knob20.KnobValue * Cos(Osc2Samp)))
Case "ABS": Osc2Samp = Abs(Osc2Samp)
End Select

''''Plot points
DrawPOINT i, Osc1Samp, WF1PB
DrawPOINT i, Osc2Samp, WF2PB
O1SBuffer(i) = Osc1Samp + &H7F
O2SBuffer(i) = Osc2Samp + &H7F

DSB(0).SetFrequency CurFreq + MT
DSB(1).SetFrequency CurFreq + MT
Next
''''End Synthesis

''''Write Waveforms
DSBWRITE 0, O1SBuffer()
DSBWRITE 1, O2SBuffer()
Me.Caption = "Os - ill  2000.DmkWare   (use keys z-/ & sd ghj l;)   " & timeGetTime - BenchStart & "ms intervals"
End Sub

Private Sub WCE_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
EC.Value = 0
End Sub
