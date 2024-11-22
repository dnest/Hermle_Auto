VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About Application"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   8880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CloseAndExit 
      Caption         =   "Exit"
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
      Left            =   7425
      TabIndex        =   2
      Top             =   3825
      Width           =   1320
   End
   Begin VB.PictureBox Picture1 
      Height          =   3930
      Left            =   120
      Picture         =   "FrmAbout.frx":0000
      ScaleHeight     =   3870
      ScaleWidth      =   3870
      TabIndex        =   0
      Top             =   120
      Width           =   3930
   End
   Begin VB.Label LabelRobotVersion 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4400
      TabIndex        =   7
      Top             =   2880
      Width           =   4100
   End
   Begin VB.Label Label4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   5880
      Width           =   4335
   End
   Begin VB.Label Label1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4400
      TabIndex        =   5
      Top             =   2280
      Width           =   4100
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "IS2904"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   22.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4400
      TabIndex        =   4
      Top             =   1476
      Width           =   4100
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Automation for Hermle"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   348
      Left            =   4400
      TabIndex        =   3
      Top             =   1044
      Width           =   4100
   End
   Begin VB.Label LabelTitle 
      Alignment       =   2  'Center
      Caption         =   "Shafir production Systems"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4400
      TabIndex        =   1
      Top             =   360
      Width           =   4100
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub CloseAndExit_Click()

    Me.Hide
    Unload Me
    
End Sub

Private Sub Form_Load()

Dim RobotVersion As String

    ReadString (30)
    FrmAbout.LabelRobotVersion.Caption = "Robot version :" & ArrayString
    
    Label1.Caption = "HMI Version :" & App.Major & "." & App.Minor & "." & App.Revision
    
End Sub

