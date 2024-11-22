VERSION 5.00
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form FrmCommunication 
   BorderStyle     =   1  'Fixed Single
   Caption         =   " Communication settings"
   ClientHeight    =   7350
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   8445
   Icon            =   "FrmCommunication.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   8445
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Robot Alarm"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3015
      Left            =   240
      TabIndex        =   22
      Top             =   4080
      Width           =   4575
      Begin VB.TextBox TextErrorCode 
         Height          =   495
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   4000
      End
      Begin VB.CommandButton CmdResetRobotAlarm 
         Caption         =   "Reset Robot Alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         TabIndex        =   24
         Top             =   2160
         Width           =   4000
      End
      Begin VB.CommandButton CmdGetRobotError 
         Caption         =   "Display Robot alarm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   23
         Top             =   480
         Width           =   4000
      End
   End
   Begin Threed.SSFrame SSFrame3 
      Height          =   3660
      Left            =   4995
      TabIndex        =   11
      Top             =   225
      Width           =   3210
      _Version        =   65536
      _ExtentX        =   5662
      _ExtentY        =   6456
      _StockProps     =   14
      Caption         =   "HandShake"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin VB.CommandButton BtnStartTimer 
         Caption         =   "Start Timer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1455
         TabIndex        =   13
         Top             =   450
         Width           =   1500
      End
      Begin VB.CommandButton BtnStopTimer 
         Caption         =   "Stop Timer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   1455
         TabIndex        =   12
         Top             =   990
         Width           =   1500
      End
      Begin Threed.SSOption OptionComm 
         Height          =   465
         Index           =   0
         Left            =   240
         TabIndex        =   16
         Top             =   1770
         Width           =   2490
         _Version        =   65536
         _ExtentX        =   4392
         _ExtentY        =   820
         _StockProps     =   78
         Caption         =   " Work Off Line"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptionComm 
         Height          =   555
         Index           =   1
         Left            =   240
         TabIndex        =   17
         Top             =   2220
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   " Read Only"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption OptionComm 
         Height          =   555
         Index           =   2
         Left            =   240
         TabIndex        =   18
         Top             =   2760
         Width           =   2355
         _Version        =   65536
         _ExtentX        =   4154
         _ExtentY        =   979
         _StockProps     =   78
         Caption         =   " Work On Line"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Value           =   -1  'True
      End
      Begin VB.Label StartTimer 
         Caption         =   "Start Timer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   315
         TabIndex        =   15
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label StopTimer 
         Caption         =   "Stop Timer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   270
         TabIndex        =   14
         Top             =   1080
         Width           =   1140
      End
   End
   Begin Threed.SSFrame SSFrame2 
      Height          =   1800
      Left            =   195
      TabIndex        =   6
      Top             =   2115
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8290
      _ExtentY        =   3175
      _StockProps     =   14
      Caption         =   "Test Comm"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin VB.CommandButton CommandTest 
         Caption         =   "run test 500"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3100
         TabIndex        =   8
         Top             =   405
         Width           =   1500
      End
      Begin VB.CommandButton BtnCommTest 
         Caption         =   "Comm Test"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3100
         TabIndex        =   7
         Top             =   945
         Width           =   1500
      End
      Begin VB.Label Label3 
         Caption         =   "One Cycle time-Read Byte"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   315
         TabIndex        =   10
         Top             =   495
         Width           =   2670
      End
      Begin VB.Label Label4 
         Caption         =   "one cycle time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   990
         TabIndex        =   9
         Top             =   990
         Width           =   1995
      End
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1800
      Left            =   195
      TabIndex        =   1
      Top             =   225
      Width           =   4695
      _Version        =   65536
      _ExtentX        =   8290
      _ExtentY        =   3175
      _StockProps     =   14
      Caption         =   "Communication"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin VB.CommandButton StartComm 
         Caption         =   "Start Comm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3100
         TabIndex        =   3
         Top             =   540
         Width           =   1500
      End
      Begin VB.CommandButton BtnCloseComm 
         Caption         =   "Close Comm"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   3100
         TabIndex        =   2
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label StartCommunication 
         Caption         =   "Start Communication with robot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   180
         TabIndex        =   5
         Top             =   585
         Width           =   2940
      End
      Begin VB.Label CloseCommunication 
         Caption         =   "Close Communication With Robot"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   45
         TabIndex        =   4
         Top             =   1125
         Width           =   3030
      End
   End
   Begin Threed.SSCommand BtnExit 
      Height          =   870
      Left            =   5040
      TabIndex        =   0
      Top             =   6240
      Width           =   3165
      _Version        =   65536
      _ExtentX        =   5583
      _ExtentY        =   1535
      _StockProps     =   78
      Caption         =   "CLOSE"
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BevelWidth      =   3
      Font3D          =   4
   End
   Begin Threed.SSFrame SSFrame4 
      Height          =   1440
      Left            =   5040
      TabIndex        =   19
      Top             =   4080
      Width           =   3135
      _Version        =   65536
      _ExtentX        =   5530
      _ExtentY        =   2540
      _StockProps     =   14
      Caption         =   " Profibus"
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   3
      Begin VB.CommandButton ResetProfibus 
         Caption         =   "Reset Communication"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   240
         TabIndex        =   20
         Top             =   720
         Width           =   2340
      End
      Begin VB.Label Label2 
         Caption         =   "Reset Profibus"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   360
         TabIndex        =   21
         Top             =   360
         Width           =   1755
      End
   End
End
Attribute VB_Name = "FrmCommunication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim ii As Double
Dim a1 As Double
Dim a2 As Double
Dim a3 As Double
Dim ret As Integer
Dim NCState As Integer


Private Sub BtnCloseComm_Click()

    Call MotoComToolBox.CloseCommunication

End Sub

Private Sub BtnCommTest_Click()
Dim bb As Boolean

    bb = TestCommunication()
    
    If bb = False Then
        Exit Sub
    End If
        
    ArrayByte(0) = 0

    a1 = Timer
    ret = BscGetVarData(m_nCid, 0, 1, ArrayByte(0))
    a2 = Timer
    
    a3 = (a2 - a1) * 1000
    Label4 = (a3)
    
    

End Sub

Private Sub BtnExit_Click()

    Me.Hide

End Sub



Public Sub BtnStartTimer_Click()
    
    ''turn the timers ON
    fMainForm.tmrUpdateRobotStatus.Enabled = True
    fMainForm.tmrRobotQuery.Enabled = True
    
End Sub

Private Sub BtnStopTimer_Click()
    
    ''turn the timers off
    fMainForm.tmrUpdateRobotStatus.Enabled = False
    fMainForm.tmrRobotQuery.Enabled = False

End Sub




Private Sub CmdCancelRobotError_Click()
''
Dim ret As Integer

    If HermleCommState = OffLine Then
        Exit Sub
    End If
    
    ret = BscCancel(m_nCid)

End Sub

Public Sub CmdGetRobotError_Click()
''
Dim ret As Integer
Dim Data As Integer
Dim Msg As String
Dim AlarmString As String * 36

    If HermleCommState = OffLine Then
        TextErrorCode.Text = " Offline ! "
        Exit Sub
    End If
    
    Call ModHandShake.ReadRobotAlarm
    TextErrorCode.Text = RobotAlarmString
    
End Sub

Private Sub CmdResetRobotAlarm_Click()
''
Dim ret As Integer


    If HermleCommState = OffLine Then
        ret = MsgBox("the system is off line.", vbInformation, _
            "CmdResetRobotAlarm_Click()")
        Exit Sub
    End If
    
    
    If AppKeyState <> remote Then
        ret = MsgBox("The Key is not in remote mode.", vbInformation _
        , "CmdResetRobotAlarm_Click()")
        Exit Sub
    End If
    
    
    ret = BscReset(m_nCid)
    Call FrmCommunication.CmdGetRobotError_Click
    
End Sub



Private Sub CommandTest_Click()

    a1 = Timer
    For ii = 1 To 500
        Call ReadByte(1, ArrayByte())
    Next
    a2 = Timer
    a3 = (a2 - a1) * 1000 / 500
    Label3 = (a3)

End Sub





Private Sub Form_Activate()

    If OptionComm(0).Value = True Then
        HermleCommState = OffLine
        
    ElseIf OptionComm(1).Value = True Then
        HermleCommState = ReadOnly
    
    ElseIf OptionComm(2).Value = True Then
        HermleCommState = OnLine
        
    End If

End Sub



Private Sub Form_Load()


    
    If OptionComm(0).Value = True Then
        HermleCommState = OffLine
        
    ElseIf OptionComm(1).Value = True Then
        HermleCommState = ReadOnly
    
    ElseIf OptionComm(2).Value = True Then
        HermleCommState = OnLine
        
    End If
    
''index  0 finish
''index  1 not finish

''state 1 finish
''state 2 notfinish
    
End Sub



Private Sub OptionComm_Click(Index As Integer, Value As Integer)

    If Index = 0 And Value = 1 Then
        HermleCommState = OffLine
        
    ElseIf Index = 1 And Value = 1 Then
        HermleCommState = ReadOnly
        
    ElseIf Index = 2 And Value = 1 Then
        HermleCommState = OnLine
        
    End If
    

End Sub

Public Sub ResetProfibus_Click()

Dim bb As Boolean

    If HermleCommState = OffLine Then
        Exit Sub
    End If

    ret = SetServo(True) ''give power to motors.
    If ret = -1 Then
        Exit Sub
    End If
    
    bb = StartJob("RESET_PROFIBUS.JBI")

    If bb = False Then
        Call FrmDialog.ShowDialogForm(33, 32, 32, "mdiMain", " BtnTestAllPockets_Click()", InputIncomplete)
        Exit Sub
    End If

End Sub




Public Sub StartComm_Click()

    Call SetCommunicationParameters
    Call MotoComToolBox.StartCommunication
    
End Sub

