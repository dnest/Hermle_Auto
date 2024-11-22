VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   8400
   ClientLeft      =   540
   ClientTop       =   825
   ClientWidth     =   7980
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   7980
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmPassword 
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   177
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1185
      Left            =   240
      TabIndex        =   38
      Top             =   120
      Width           =   3465
      Begin VB.TextBox txtPassword 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFC0&
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   39
         ToolTipText     =   "(1-100)"
         Top             =   405
         Width           =   1755
      End
      Begin Threed.SSCommand cmdPassword 
         Height          =   510
         Left            =   120
         TabIndex        =   40
         Top             =   405
         Width           =   1305
         _Version        =   65536
         _ExtentX        =   2302
         _ExtentY        =   900
         _StockProps     =   78
         Caption         =   "Set"
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
         Font3D          =   1
      End
   End
   Begin TabDlg.SSTab OptionsTab 
      Height          =   5775
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   10186
      _Version        =   393216
      Tabs            =   6
      Tab             =   2
      TabsPerRow      =   6
      TabHeight       =   1058
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Information"
      TabPicture(0)   =   "frmOptions.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Label5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label8"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label9"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label10"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label11"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "TextCountry"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "TextAutoName"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "CmdSaveInfo"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TextAutoNumber"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TextHermleNumber"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TextHermleType"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextFactory"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "CmdReadInfo"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      TabCaption(1)   =   "Backup and restore"
      TabPicture(1)   =   "frmOptions.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FrameReadLocationsFromRobot"
      Tab(1).Control(1)=   "Frame11"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "configuration"
      TabPicture(2)   =   "frmOptions.frx":0342
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "PanelShelf(3)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "PanelShelf(2)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "PanelShelf(1)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "CmdDisplayConfig"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "SSPanel1"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "More    Options"
      TabPicture(3)   =   "frmOptions.frx":035E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "CmdClearAllStatus"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "CmdResetAllWP"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "CmdResetToolCounter"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "TextMoreOptions"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).ControlCount=   4
      TabCaption(4)   =   "Gripper"
      TabPicture(4)   =   "frmOptions.frx":037A
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "LabelSelectGripper"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "CmdSelectGripper"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "TextSelectGripper"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "SelectGripper(1)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).Control(4)=   "SelectGripper(2)"
      Tab(4).Control(4).Enabled=   0   'False
      Tab(4).Control(5)=   "SelectGripper(3)"
      Tab(4).Control(5).Enabled=   0   'False
      Tab(4).Control(6)=   "SelectGripper(4)"
      Tab(4).Control(6).Enabled=   0   'False
      Tab(4).Control(7)=   "SelectGripper(5)"
      Tab(4).Control(7).Enabled=   0   'False
      Tab(4).Control(8)=   "SelectGripper(6)"
      Tab(4).Control(8).Enabled=   0   'False
      Tab(4).Control(9)=   "SelectGripper(7)"
      Tab(4).Control(9).Enabled=   0   'False
      Tab(4).ControlCount=   10
      TabCaption(5)   =   "All     Locations"
      TabPicture(5)   =   "frmOptions.frx":0396
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "AllLocationsTable"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).Control(1)=   "CmdAllLocationsTableDisplay"
      Tab(5).Control(1).Enabled=   0   'False
      Tab(5).ControlCount=   2
      Begin VB.TextBox TextMoreOptions 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   -74520
         ScrollBars      =   3  'Both
         TabIndex        =   60
         Top             =   4800
         Width           =   6255
      End
      Begin VB.CommandButton CmdResetToolCounter 
         Caption         =   "Reset Work Piece"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -73800
         TabIndex        =   59
         Top             =   3000
         Width           =   4900
      End
      Begin VB.CommandButton CmdAllLocationsTableDisplay 
         Caption         =   "Display all Locations"
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
         Left            =   -74760
         TabIndex        =   58
         Top             =   4800
         Width           =   3255
      End
      Begin VB.CommandButton CmdResetAllWP 
         Caption         =   "Clear All Work Piece table"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -73800
         TabIndex        =   56
         Top             =   2280
         Width           =   4900
      End
      Begin VB.OptionButton SelectGripper 
         Caption         =   "Round Big"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   7
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   4080
         Width           =   1860
      End
      Begin VB.OptionButton SelectGripper 
         Caption         =   "Round Small"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   6
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   3600
         Width           =   1860
      End
      Begin VB.OptionButton SelectGripper 
         Caption         =   "HSK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   5
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   3120
         Width           =   1860
      End
      Begin VB.OptionButton SelectGripper 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   4
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   2640
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.OptionButton SelectGripper 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   3
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   2160
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.OptionButton SelectGripper 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Index           =   2
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   1680
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.OptionButton SelectGripper 
         Height          =   400
         Index           =   1
         Left            =   -73680
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   1200
         Visible         =   0   'False
         Width           =   1860
      End
      Begin VB.TextBox TextSelectGripper 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -71280
         TabIndex        =   46
         Text            =   "7"
         Top             =   3360
         Width           =   735
      End
      Begin Threed.SSPanel SSPanel1 
         Height          =   1005
         Left            =   3720
         TabIndex        =   43
         Top             =   1080
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   1773
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label labelGripperType 
            Alignment       =   2  'Center
            Caption         =   "Tool Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   45
            Top             =   360
            Width           =   1320
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Gripper"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            TabIndex        =   44
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdReadInfo 
         Caption         =   "Display Information"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74160
         TabIndex        =   42
         Top             =   4680
         Width           =   3000
      End
      Begin VB.TextBox TextFactory 
         Height          =   375
         Left            =   -74040
         TabIndex        =   41
         Top             =   2760
         Width           =   2000
      End
      Begin VB.CommandButton CmdClearAllStatus 
         Caption         =   "Clear All Pockets Status"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   -73800
         TabIndex        =   37
         Top             =   1560
         Width           =   4900
      End
      Begin VB.CommandButton CmdDisplayConfig 
         Caption         =   "Display Configuration"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   4560
         Width           =   3135
      End
      Begin VB.TextBox TextHermleType 
         Height          =   400
         Left            =   -70515
         TabIndex        =   31
         Top             =   3960
         Width           =   2000
      End
      Begin VB.TextBox TextHermleNumber 
         Height          =   400
         Left            =   -70515
         TabIndex        =   29
         Top             =   2760
         Width           =   2000
      End
      Begin VB.TextBox TextAutoNumber 
         Height          =   400
         Left            =   -70545
         TabIndex        =   27
         Top             =   1560
         Width           =   2000
      End
      Begin VB.CommandButton CmdSaveInfo 
         Caption         =   "Save Changes"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -71040
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   4680
         Width           =   3000
      End
      Begin Threed.SSPanel PanelShelf 
         Height          =   1005
         Index           =   1
         Left            =   600
         TabIndex        =   20
         Top             =   1080
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   1764
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.21
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Shelf 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   280
            TabIndex        =   33
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label LabelType 
            Alignment       =   2  'Center
            Caption         =   "Tool Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   1485
            TabIndex        =   23
            Top             =   375
            Width           =   1320
         End
      End
      Begin VB.TextBox TextAutoName 
         Height          =   400
         Left            =   -74025
         TabIndex        =   18
         Top             =   3960
         Width           =   2000
      End
      Begin VB.TextBox TextCountry 
         Height          =   400
         Left            =   -74025
         TabIndex        =   17
         Top             =   1560
         Width           =   2000
      End
      Begin VB.Frame Frame4 
         Caption         =   "Gripper Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1275
         Left            =   360
         TabIndex        =   10
         Top             =   4320
         Width           =   3360
         Begin VB.OptionButton OptionsHSK 
            Caption         =   "HSK"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   240
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   360
            Width           =   900
         End
         Begin VB.OptionButton OptionsDrills 
            Caption         =   "Drill"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   360
            Width           =   900
         End
         Begin VB.OptionButton OptionsRound 
            Caption         =   "Round"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   650
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   360
            Width           =   900
         End
      End
      Begin VB.Frame Frame11 
         Caption         =   "Read From File "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2385
         Left            =   -74640
         TabIndex        =   2
         Top             =   840
         Width           =   6720
         Begin Threed.SSCommand ButtonRestoreGeneralLocations 
            Height          =   800
            Left            =   3350
            TabIndex        =   3
            Top             =   360
            Width           =   3200
            _Version        =   65536
            _ExtentX        =   5644
            _ExtentY        =   1411
            _StockProps     =   78
            Caption         =   "Restore General Locations"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin Threed.SSCommand BtnReadPocketLocations 
            Height          =   795
            Left            =   3345
            TabIndex        =   4
            Top             =   1320
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5644
            _ExtentY        =   1411
            _StockProps     =   78
            Caption         =   "Read pocket Locations "
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label Label1 
            Caption         =   "Restore 11 positions to robot"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   180
            TabIndex        =   6
            Top             =   600
            Width           =   3030
         End
         Begin VB.Label Label2 
            Caption         =   "read all pockets location from file to Computer's memory."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   1320
            Width           =   3135
         End
      End
      Begin Threed.SSFrame FrameReadLocationsFromRobot 
         Height          =   1695
         Left            =   -74640
         TabIndex        =   7
         Top             =   3720
         Width           =   6720
         _Version        =   65536
         _ExtentX        =   11853
         _ExtentY        =   2990
         _StockProps     =   14
         Caption         =   "From Robot To File"
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin Threed.SSCommand CmdBackUpAllGeneralLocations 
            Height          =   795
            Left            =   3345
            TabIndex        =   8
            Top             =   600
            Width           =   3195
            _Version        =   65536
            _ExtentX        =   5644
            _ExtentY        =   1411
            _StockProps     =   78
            Caption         =   "BackUp General locations"
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   3
         End
         Begin VB.Label Label29 
            Caption         =   "Read all general locations  from robot into file."
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   3075
         End
      End
      Begin Threed.SSPanel PanelShelf 
         Height          =   1005
         Index           =   2
         Left            =   600
         TabIndex        =   21
         Top             =   2160
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   1764
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.21
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Shelf 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   280
            TabIndex        =   34
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label LabelType 
            Alignment       =   2  'Center
            Caption         =   "Tool Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   1485
            TabIndex        =   24
            Top             =   375
            Width           =   1320
         End
      End
      Begin Threed.SSPanel PanelShelf 
         Height          =   1005
         Index           =   3
         Left            =   600
         TabIndex        =   22
         Top             =   3240
         Width           =   3000
         _Version        =   65536
         _ExtentX        =   5292
         _ExtentY        =   1764
         _StockProps     =   15
         BackColor       =   13160660
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.21
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Shelf 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   280
            TabIndex        =   35
            Top             =   375
            Width           =   1005
         End
         Begin VB.Label LabelType 
            Alignment       =   2  'Center
            Caption         =   "Tool Type"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   1485
            TabIndex        =   25
            Top             =   375
            Width           =   1320
         End
      End
      Begin Threed.SSCommand CmdSelectGripper 
         Height          =   855
         Left            =   -70440
         TabIndex        =   47
         Top             =   3360
         Width           =   2415
         _Version        =   65536
         _ExtentX        =   4260
         _ExtentY        =   1508
         _StockProps     =   78
         Caption         =   "Send to robot"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Enabled         =   0   'False
         BevelWidth      =   3
         Font3D          =   2
      End
      Begin MSFlexGridLib.MSFlexGrid AllLocationsTable 
         Height          =   3795
         Left            =   -74760
         TabIndex        =   57
         Top             =   840
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   6694
         _Version        =   393216
         Rows            =   129
         Cols            =   8
         RowHeightMin    =   250
         AllowBigSelection=   0   'False
         FocusRect       =   0
         HighLight       =   0
         GridLines       =   3
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label LabelSelectGripper 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "...Select gripper..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74760
         TabIndex        =   55
         Top             =   4920
         Width           =   6855
      End
      Begin VB.Label Label11 
         Caption         =   "Hermle Type"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70515
         TabIndex        =   32
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label Label10 
         Caption         =   "Hermle Number"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   -70515
         TabIndex        =   30
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label9 
         Caption         =   "Automation Number"
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
         Left            =   -70545
         TabIndex        =   28
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label8 
         Caption         =   "Country"
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
         Left            =   -74025
         TabIndex        =   19
         Top             =   1200
         Width           =   1995
      End
      Begin VB.Label Label6 
         Caption         =   "Automation Name"
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
         Left            =   -74025
         TabIndex        =   16
         Top             =   3600
         Width           =   1995
      End
      Begin VB.Label Label5 
         Caption         =   "Factory"
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
         Left            =   -74025
         TabIndex        =   15
         Top             =   2400
         Width           =   1995
      End
      Begin VB.Label Label3 
         Caption         =   "Country"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   177
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   -74500
         TabIndex        =   14
         Top             =   840
         Width           =   1500
      End
   End
   Begin Threed.SSCommand CloseAndExit 
      Height          =   735
      Left            =   4200
      TabIndex        =   0
      Top             =   7440
      Width           =   3315
      _Version        =   65536
      _ExtentX        =   5847
      _ExtentY        =   1296
      _StockProps     =   78
      Caption         =   "EXIT"
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   2
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub BtnReadGeneralLocation_Click()
''1.the function read data from the controller into the memory
''2.the function write the data into the TextFile.

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim GeneralPositionas As String
Dim jj As Integer

 GeneralPositionas = ComboGeneralLocations.Text

    Select Case GeneralPositionas
   
        '''*********
        '''DRILL
        '''*********
    Case "Pocket 1_1_1"
        Call MotoComToolBox.ReadPosition(21, TempPosition())
        shelf = 1: column = 1: pocket = 1
        DrillLocations(shelf, column).diameter(1).X = TempPosition(2)
        DrillLocations(shelf, column).diameter(1).Y = TempPosition(3)
        DrillLocations(shelf, column).diameter(1).z = TempPosition(4)
        DrillLocations(shelf, column).diameter(1).Rx = TempPosition(5)
        DrillLocations(shelf, column).diameter(1).Ry = TempPosition(6)
        DrillLocations(shelf, column).diameter(1).Rz = TempPosition(7)
        Call SaveArray("DrillLocations")
        
    Case "Pocket 1_12_1"
        Call MotoComToolBox.ReadPosition(22, TempPosition())
        shelf = 1: column = 12: pocket = 1
        DrillLocations(shelf, column).diameter(12).X = TempPosition(2)
        DrillLocations(shelf, column).diameter(12).Y = TempPosition(3)
        DrillLocations(shelf, column).diameter(12).z = TempPosition(4)
        DrillLocations(shelf, column).diameter(12).Rx = TempPosition(5)
        DrillLocations(shelf, column).diameter(12).Ry = TempPosition(6)
        DrillLocations(shelf, column).diameter(12).Rz = TempPosition(7)
        Call SaveArray("DrillLocations")
    
    Case "Pocket 2_1_1"
        Call MotoComToolBox.ReadPosition(23, TempPosition())
        shelf = 2: column = 1: pocket = 1
        DrillLocations(shelf, column).diameter(1).X = TempPosition(2)
        DrillLocations(shelf, column).diameter(1).Y = TempPosition(3)
        DrillLocations(shelf, column).diameter(1).z = TempPosition(4)
        DrillLocations(shelf, column).diameter(1).Rx = TempPosition(5)
        DrillLocations(shelf, column).diameter(1).Ry = TempPosition(6)
        DrillLocations(shelf, column).diameter(1).Rz = TempPosition(7)
        Call SaveArray("DrillLocations")
        
    Case "Pocket 2_12_1"
        Call MotoComToolBox.ReadPosition(24, TempPosition())
        shelf = 1: column = 12: pocket = 1
        DrillLocations(shelf, column).diameter(12).X = TempPosition(2)
        DrillLocations(shelf, column).diameter(12).Y = TempPosition(3)
        DrillLocations(shelf, column).diameter(12).z = TempPosition(4)
        DrillLocations(shelf, column).diameter(12).Rx = TempPosition(5)
        DrillLocations(shelf, column).diameter(12).Ry = TempPosition(6)
        DrillLocations(shelf, column).diameter(12).Rz = TempPosition(7)
        Call SaveArray("DrillLocations")
    
    Case "Pocket 3_1_1"
        Call MotoComToolBox.ReadPosition(25, TempPosition())
        shelf = 3: column = 1: pocket = 1
        DrillLocations(shelf, column).diameter(1).X = TempPosition(2)
        DrillLocations(shelf, column).diameter(1).Y = TempPosition(3)
        DrillLocations(shelf, column).diameter(1).z = TempPosition(4)
        DrillLocations(shelf, column).diameter(1).Rx = TempPosition(5)
        DrillLocations(shelf, column).diameter(1).Ry = TempPosition(6)
        DrillLocations(shelf, column).diameter(1).Rz = TempPosition(7)
        Call SaveArray("DrillLocations")
    
    Case "Pocket 3_12_1"
        Call MotoComToolBox.ReadPosition(26, TempPosition())
        shelf = 3: column = 12: pocket = 1
        DrillLocations(shelf, column).diameter(12).X = TempPosition(2)
        DrillLocations(shelf, column).diameter(12).Y = TempPosition(3)
        DrillLocations(shelf, column).diameter(12).z = TempPosition(4)
        DrillLocations(shelf, column).diameter(12).Rx = TempPosition(5)
        DrillLocations(shelf, column).diameter(12).Ry = TempPosition(6)
        DrillLocations(shelf, column).diameter(12).Rz = TempPosition(7)
        Call SaveArray("DrillLocations")
        
        '''*********
        '''HSK
        '''*********
        Case "Pocket 1_1"
        Call MotoComToolBox.ReadPosition(21, TempPosition())
        shelf = 1: column = 1: pocket = 1
        HSKLocations(shelf, column).X = TempPosition(2)
        HSKLocations(shelf, column).Y = TempPosition(3)
        HSKLocations(shelf, column).z = TempPosition(4)
        HSKLocations(shelf, column).Rx = TempPosition(5)
        HSKLocations(shelf, column).Ry = TempPosition(6)
        HSKLocations(shelf, column).Rz = TempPosition(7)
        Call SaveArray("HSKLocations")
        
        GeneralLocation(21).X = TempPosition(2)
        GeneralLocation(21).Y = TempPosition(3)
        GeneralLocation(21).z = TempPosition(4)
        GeneralLocation(21).Rx = TempPosition(5)
        GeneralLocation(21).Ry = TempPosition(6)
        GeneralLocation(21).Rz = TempPosition(7)
        Call WriteGeneralLocations
        
        GoTo FeedBack ''''Exit Sub
        
    Case "Pocket 1_10"
        Call MotoComToolBox.ReadPosition(22, TempPosition())
        shelf = 1: column = 10: pocket = 1
        HSKLocations(shelf, column).X = TempPosition(2)
        HSKLocations(shelf, column).Y = TempPosition(3)
        HSKLocations(shelf, column).z = TempPosition(4)
        HSKLocations(shelf, column).Rx = TempPosition(5)
        HSKLocations(shelf, column).Ry = TempPosition(6)
        HSKLocations(shelf, column).Rz = TempPosition(7)
        Call SaveArray("HSKLocations")
        
        GeneralLocation(22).X = TempPosition(2)
        GeneralLocation(22).Y = TempPosition(3)
        GeneralLocation(22).z = TempPosition(4)
        GeneralLocation(22).Rx = TempPosition(5)
        GeneralLocation(22).Ry = TempPosition(6)
        GeneralLocation(22).Rz = TempPosition(7)
        Call WriteGeneralLocations
        
        GoTo FeedBack ''''Exit Sub
    
    Case "Pocket 2_1"
        Call MotoComToolBox.ReadPosition(23, TempPosition())
        shelf = 2: column = 1: pocket = 1
        HSKLocations(shelf, column).X = TempPosition(2)
        HSKLocations(shelf, column).Y = TempPosition(3)
        HSKLocations(shelf, column).z = TempPosition(4)
        HSKLocations(shelf, column).Rx = TempPosition(5)
        HSKLocations(shelf, column).Ry = TempPosition(6)
        HSKLocations(shelf, column).Rz = TempPosition(7)
        Call SaveArray("HSKLocations")
        
        GeneralLocation(23).X = TempPosition(2)
        GeneralLocation(23).Y = TempPosition(3)
        GeneralLocation(23).z = TempPosition(4)
        GeneralLocation(23).Rx = TempPosition(5)
        GeneralLocation(23).Ry = TempPosition(6)
        GeneralLocation(23).Rz = TempPosition(7)
        Call WriteGeneralLocations
        
        GoTo FeedBack ''''Exit Sub
    
    Case "Pocket 2_10"
        Call MotoComToolBox.ReadPosition(24, TempPosition())
        shelf = 1: column = 10: pocket = 1
        HSKLocations(shelf, column).X = TempPosition(2)
        HSKLocations(shelf, column).Y = TempPosition(3)
        HSKLocations(shelf, column).z = TempPosition(4)
        HSKLocations(shelf, column).Rx = TempPosition(5)
        HSKLocations(shelf, column).Ry = TempPosition(6)
        HSKLocations(shelf, column).Rz = TempPosition(7)
        Call SaveArray("HSKLocations")
        
        GeneralLocation(24).X = TempPosition(2)
        GeneralLocation(24).Y = TempPosition(3)
        GeneralLocation(24).z = TempPosition(4)
        GeneralLocation(24).Rx = TempPosition(5)
        GeneralLocation(24).Ry = TempPosition(6)
        GeneralLocation(24).Rz = TempPosition(7)
        Call WriteGeneralLocations
        
       GoTo FeedBack ''''Exit Sub
    
    Case "Pocket 3_1"
        Call MotoComToolBox.ReadPosition(25, TempPosition())
        shelf = 3: column = 1: pocket = 1
        HSKLocations(shelf, column).X = TempPosition(2)
        HSKLocations(shelf, column).Y = TempPosition(3)
        HSKLocations(shelf, column).z = TempPosition(4)
        HSKLocations(shelf, column).Rx = TempPosition(5)
        HSKLocations(shelf, column).Ry = TempPosition(6)
        HSKLocations(shelf, column).Rz = TempPosition(7)
        Call SaveArray("HSKLocations")
        
        GeneralLocation(25).X = TempPosition(2)
        GeneralLocation(25).Y = TempPosition(3)
        GeneralLocation(25).z = TempPosition(4)
        GeneralLocation(25).Rx = TempPosition(5)
        GeneralLocation(25).Ry = TempPosition(6)
        GeneralLocation(25).Rz = TempPosition(7)
        Call WriteGeneralLocations
        
        GoTo FeedBack ''''Exit Sub
    
    Case "Pocket 3_10"
        Call MotoComToolBox.ReadPosition(26, TempPosition())
        shelf = 3: column = 10: pocket = 1
        HSKLocations(shelf, column).X = TempPosition(2)
        HSKLocations(shelf, column).Y = TempPosition(3)
        HSKLocations(shelf, column).z = TempPosition(4)
        HSKLocations(shelf, column).Rx = TempPosition(5)
        HSKLocations(shelf, column).Ry = TempPosition(6)
        HSKLocations(shelf, column).Rz = TempPosition(7)
        Call SaveArray("HSKLocations")
        
        GeneralLocation(26).X = TempPosition(2)
        GeneralLocation(26).Y = TempPosition(3)
        GeneralLocation(26).z = TempPosition(4)
        GeneralLocation(26).Rx = TempPosition(5)
        GeneralLocation(26).Ry = TempPosition(6)
        GeneralLocation(26).Rz = TempPosition(7)
        Call WriteGeneralLocations
        GoTo FeedBack ''''Exit Sub
        
    Case "ALL"
        For jj = 1 To AmountOfGeneralLocations
            Call ReadPosition(jj, ArrayPosition())
            GeneralLocation(jj).X = ArrayPosition(2)
            GeneralLocation(jj).Y = ArrayPosition(3)
            GeneralLocation(jj).z = ArrayPosition(4)
            GeneralLocation(jj).Rx = ArrayPosition(5)
            GeneralLocation(jj).Ry = ArrayPosition(6)
            GeneralLocation(jj).Rz = ArrayPosition(7)
        Next
        Call WriteGeneralLocations
        GoTo FeedBack ''''Exit Sub
        
    End Select
    
FeedBack:
    
    Call FrmDialog.ShowDialogForm(50, 50, 50, "frmOptions", "BtnReadGeneralLocation_Click()", GreenV)
Exit Sub

End Sub



Private Sub ButtonRestoreGeneralLocations_Click()


''1.this function read all general locations from file into the memory array.
''2.the function send data in array to the controller
''3.the f take no parameters and return no parameters.
''4.the f is called from the GUI ONLY.
Dim ret As Integer

On Error GoTo error
    
    If AppKeyState <> remote Then
        ret = MsgBox("The key is not in remote mode." & vbCrLf _
        & "Please set the key to remote and press again." _
        , vbExclamation _
        , "Restore locations from file to robot.")
        Exit Sub
    End If
    
    ''read general location from the right file (according to the ToolType) to memory
    Call ModCsvFile.ReadGeneralLocation
    
    ''send data to the controller.
    Call MotoComToolBox.WriteAllGeneralLocationToRobot
    
    ret = MsgBox("general locatiion had been restored" _
        , vbInformation, "restore general locations")
        
    ModLogFile.LogAddLine (" restore general locations:")
    Call ModGUI.ListHMIUpdate(" restore general locations:")
    Exit Sub
        
error:
    ret = MsgBox("Error while restore general locations" & vbCrLf _
        & "the error is : " & Err.Description _
        , vbExclamation _
        , "Restore general Locations")
    
End Sub

Private Sub CloseAndExit_Click()

    PasswordOk = False
    txtPassword.Text = ""
    Me.Hide
    
End Sub

Public Sub CmdAllLocationsTableDisplay_Click()


Dim line As Long
Dim column As Integer

    
    ModCsvFile.ReadGeneralLocation
    AllLocationsTable.ColAlignment(0) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(1) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(2) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(3) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(4) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(5) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(6) = flexAlignCenterCenter
    AllLocationsTable.ColAlignment(7) = flexAlignCenterCenter

    AllLocationsTable.TextMatrix(0, 0) = ""
    AllLocationsTable.TextMatrix(0, 1) = " Name "
    AllLocationsTable.TextMatrix(0, 2) = "X"
    AllLocationsTable.TextMatrix(0, 3) = "Y"
    AllLocationsTable.TextMatrix(0, 4) = "Z"
    AllLocationsTable.TextMatrix(0, 5) = "Rx"
    AllLocationsTable.TextMatrix(0, 6) = "Ry"
    AllLocationsTable.TextMatrix(0, 7) = "Rz"
        
    For line = 1 To 127
    
        AllLocationsTable.ColWidth(0) = 1000
        AllLocationsTable.ColWidth(1) = 2500
        AllLocationsTable.ColWidth(2) = 1000
        AllLocationsTable.ColWidth(3) = 1000
        AllLocationsTable.ColWidth(4) = 1000
        AllLocationsTable.ColWidth(5) = 1000
        AllLocationsTable.ColWidth(6) = 1000
        AllLocationsTable.ColWidth(7) = 1000
        
        AllLocationsTable.TextMatrix(line, 0) = CStr(line)
        Call ModGUI.DisplayGenLocNames
        AllLocationsTable.TextMatrix(line, 2) = GeneralLocation(line).X
        AllLocationsTable.TextMatrix(line, 3) = GeneralLocation(line).Y
        AllLocationsTable.TextMatrix(line, 4) = GeneralLocation(line).z
        AllLocationsTable.TextMatrix(line, 5) = GeneralLocation(line).Rx
        AllLocationsTable.TextMatrix(line, 6) = GeneralLocation(line).Ry
        AllLocationsTable.TextMatrix(line, 7) = GeneralLocation(line).Rz
        
    Next line
    
    
End Sub

Private Sub CmdBackUpAllGeneralLocations_Click()

Dim jj As Integer


    jj = 10
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
    
    jj = 11
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
        
            
    jj = 12
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 21
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 22
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 23
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 24
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 25
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 26
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 120
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
           
    jj = 121
    Call ReadPosition(jj, ArrayPosition())    ''read one position
    GeneralLocation(jj).X = ArrayPosition(2)  ''Copy position in memory
    GeneralLocation(jj).Y = ArrayPosition(3)
    GeneralLocation(jj).z = ArrayPosition(4)
    GeneralLocation(jj).Rx = ArrayPosition(5)
    GeneralLocation(jj).Ry = ArrayPosition(6)
    GeneralLocation(jj).Rz = ArrayPosition(7)
       
    ModCsvFile.WriteGeneralLocations ''write to hardDisk.
    ModLogFile.LogAddLine (" back general locations ")
    
    Call ModGUI.ListHMIUpdate(" back up general locations ")

    
End Sub

Private Sub CmdClearAllStatus_Click()

On Error GoTo error

    Call ModAutomationStatus.AutomationStatusReset
    Call fMainForm.BtnDisplayPocketStatus_Click
    
    ModLogFile.LogAddLine (" Clear all pockets status ")
    Call ModGUI.ListHMIUpdate("Clear all pockets status ")
    TextMoreOptions.Text = "Reset All pockets's status done."
    
    Exit Sub
error:
    TextMoreOptions.Text = "Error while reset pokets status."
    
End Sub

Public Sub CmdDisplayConfig_Click()
''other = 0
''hsk = 1
''drill = 2
''Round = 3

Dim ii As Integer

    Debug.Print "CmdDisplayShelvs_Click"
    
    Call ModIni.ReadShelvsConfig ''read from ini file
    Call ModIni.ReadAppToolType  ''read from ini file
    
    If AppShelvs(1).ShelfToolType = Drill Then
        frmOptions.LabelType(1).Caption = "Drill"
       
    ElseIf AppShelvs(1).ShelfToolType = HSK Then
    frmOptions.LabelType(1).Caption = "HSK"
       
    ElseIf AppShelvs(1).ShelfToolType = Round Then
        frmOptions.LabelType(1).Caption = "Round"
        
    ElseIf AppShelvs(1).ShelfToolType = other Then
        frmOptions.LabelType(1).Caption = "Other"
        
    End If
        
    If AppShelvs(2).ShelfToolType = Drill Then
        frmOptions.LabelType(2).Caption = "Drill"
        
    ElseIf AppShelvs(2).ShelfToolType = HSK Then
        frmOptions.LabelType(2).Caption = "HSK"
       
    ElseIf AppShelvs(2).ShelfToolType = Round Then
        frmOptions.LabelType(2).Caption = "Round"
       
    ElseIf AppShelvs(2).ShelfToolType = other Then
        frmOptions.LabelType(2).Caption = "other"
       
    End If
        
    If AppShelvs(3).ShelfToolType = Drill Then
        frmOptions.LabelType(3).Caption = "Drill"
       
    ElseIf AppShelvs(3).ShelfToolType = HSK Then
        frmOptions.LabelType(3).Caption = "HSK"
        
    ElseIf AppShelvs(3).ShelfToolType = Round Then
        frmOptions.LabelType(3).Caption = "Round"
       
    ElseIf AppShelvs(3).ShelfToolType = other Then
        frmOptions.LabelType(3).Caption = "other"
       
    End If
        
    If AppToolType = Drill Then ''set the application GUI acoording to the ToolType.
        frmOptions.OptionsDrills_Click
        frmOptions.OptionsDrills.Value = True
        labelGripperType.Caption = "Drill"
        
    ElseIf AppToolType = HSK Then
        frmOptions.OptionsHSK_Click
        frmOptions.OptionsHSK.Value = True
        labelGripperType.Caption = "HSK"
        
    ElseIf AppToolType = Round Then
        frmOptions.OptionsRound_Click
        frmOptions.OptionsRound.Value = True
        labelGripperType.Caption = "Round"
        
    End If

End Sub

Private Sub cmdPassword_Click()

    If txtPassword = "2468" Then
        PasswordOk = True
        OptionsTab.Enabled = True
        frmOptions.Width = 8040
        frmOptions.Height = 8880
        OptionsTab.Visible = True
         
    Else
        PasswordOk = False
        OptionsTab.Enabled = False
        frmOptions.Width = 8040
        frmOptions.Height = 8880

    End If
    
    ModLogFile.LogAddLine (" password entered ")
    Call ModGUI.ListHMIUpdate(" password entered ")
    
End Sub


Public Sub CmdReadInfo_Click()
    
''1.the function read information from file and display info on screen.

Dim TempString As String * 10
Dim LongRet As Long

    LongRet = GetPrivateProfileString _
    ("information", "Country", "   ", TempString, 10, App.path & "\WorkDirectory\data\IS2904.ini")
    HermleAutomation.LocationCountry = Left(TempString, LongRet)
    
    
    LongRet = GetPrivateProfileString _
    ("information", "factory", "   ", TempString, 10, App.path & "\WorkDirectory\data\IS2904.ini")
    HermleAutomation.LocationFactory = Left(TempString, LongRet)
    
    
    LongRet = GetPrivateProfileString _
    ("information", "AutoName", "   ", TempString, 10, App.path & "\WorkDirectory\data\IS2904.ini")
    HermleAutomation.AutomationName = Left(TempString, LongRet)
    
    
    LongRet = GetPrivateProfileString _
    ("information", "AutoNumber", "   ", TempString, 10, App.path & "\WorkDirectory\data\IS2904.ini")
    HermleAutomation.AutomationNumber = Left(TempString, LongRet)
    
    
    LongRet = GetPrivateProfileString _
    ("information", "HermleNumber", "   ", TempString, 10, App.path & "\WorkDirectory\data\IS2904.ini")
    HermleAutomation.HermleNumber = Left(TempString, LongRet)
    
    
    LongRet = GetPrivateProfileString _
    ("information", "HermleType", "   ", TempString, 10, App.path & "\WorkDirectory\data\IS2904.ini")
    HermleAutomation.HermleType = Left(TempString, LongRet)
    
    
    TextAutoName.Text = HermleAutomation.AutomationName
    TextAutoNumber.Text = HermleAutomation.AutomationNumber
    TextHermleNumber.Text = HermleAutomation.HermleNumber
    TextHermleType.Text = HermleAutomation.HermleType
    TextCountry.Text = HermleAutomation.LocationCountry
    TextFactory.Text = HermleAutomation.LocationFactory
    
    
End Sub



Private Sub CmdResetAllWP_Click()

On Error GoTo error

    modAllWorkPiece.AllWPReset
    Call fMainForm.BtnDisplayWPTable_Click

    TextMoreOptions.Text = "Reset All WorkPiece Done."
    Exit Sub
    
    
error:
     TextMoreOptions.Text = "Error : Reset All WorkPiece Fail."

End Sub

Private Sub CmdResetToolCounter_Click()
    
Dim TempArray(8) As Double

    TempArray(0) = 0
    Call WriteInteger(11, TempArray())  '' tool counter
    Call WriteByte(104, TempArray()) ''    WPiece Counter
    Call WriteByte(23, TempArray()) ''     no more workpiece
    Call WriteByte(105, TempArray()) '''   the first part of the first WP.
    
    ChuckUnloadFirstTime = True
    TextMoreOptions.Text = "Reset Tool counter[I 11] , Reset Wp Counter[b 104] , Reset b23 , Reset b105."
    Exit Sub
    
error:
     TextMoreOptions.Text = "Error : Reset Tool counter Fail."

End Sub



Public Sub CmdSaveInfo_Click()

    HermleAutomation.AutomationName = TextAutoName.Text
    HermleAutomation.AutomationNumber = TextAutoNumber.Text
    HermleAutomation.HermleNumber = TextHermleNumber.Text
    HermleAutomation.HermleType = TextHermleType.Text
    HermleAutomation.LocationCountry = TextCountry.Text
    HermleAutomation.LocationFactory = TextFactory.Text
    
    Call SaveMachineParameters

End Sub

Private Sub CmdSelectGripper_Click()

Dim bb As Boolean
Dim ret As Integer

On Error GoTo label

    If SelectGripper(5).Value = True Then
        ArrayDouble(0) = 5
        
    ElseIf SelectGripper(6).Value = True Then
        ArrayDouble(0) = 6
        
    ElseIf SelectGripper(7).Value = True Then
        ArrayDouble(0) = 7
        
    Else
        ArrayDouble(0) = 7
        
    End If

    If AppKeyState <> remote Then
        LabelSelectGripper.ForeColor = vbRed
         LabelSelectGripper.Caption = "The Key is not in REMOTE mode"
        Exit Sub
    End If
    
    ret = SetServo(True) ''give power to motors.
    If ret = -1 Then
        LabelSelectGripper.ForeColor = vbRed
        LabelSelectGripper.Caption = "Can  not set Servo ON."
        Exit Sub
    End If
        
    bb = WriteInteger(56, ArrayDouble())
    If bb = False Then
        LabelSelectGripper.ForeColor = vbRed
        LabelSelectGripper.Caption = "Error while sending Gripper number :" & ArrayDouble(0) & "to controller"
        Exit Sub
    End If
        
    bb = StartJob("TOOL.JBI")
    If bb = False Then
        LabelSelectGripper.ForeColor = vbRed
        LabelSelectGripper.Caption = "Can not start JOB :" & "TOOL.JBI"
        Exit Sub
    End If
    
    LabelSelectGripper.ForeColor = &H8000&
    LabelSelectGripper.Caption = "Gripper number " & ArrayDouble(0) & " Sent to Integer 56 in the Controller"
    
    ModLogFile.LogAddLine (" send gripper number to robot.Gripper number : " & CStr(ArrayDouble(0)))
    Call ModGUI.ListHMIUpdate(" send gripper number to robot.Gripper number : " & CStr(ArrayDouble(0)))
    
    
    Exit Sub
    
Exit Sub
label:
    ret = MsgBox("error while sending gripper to robot" _
    , vbExclamation, "Select gripper")
End Sub


Private Sub Form_Activate()

    PasswordOk = False
    
'    If PasswordOk = True Then
'        OptionsTab.Enabled = True
'        frmOptions.Width = 8040
'        frmOptions.Height = 8880
'    Else
'        OptionsTab.Enabled = False
'        frmOptions.Width = 8040
'        frmOptions.Height = 8880
'    End If

    OptionsTab.Visible = False ''hide the form till the password set

End Sub

Private Sub Form_Load()

Dim kk As Integer
Dim jj As Integer
Dim BoolShelvs(3) As Boolean

    PasswordOk = False
    
    If PasswordOk = True Then
        OptionsTab.Enabled = True
        frmOptions.Width = 8040
        frmOptions.Height = 8880
    Else
        OptionsTab.Enabled = False
        frmOptions.Width = 8040
        frmOptions.Height = 8880
    End If
    
    OptionsTab.Visible = False ''hide the form till the password set
    frmOptions.CmdDisplayConfig_Click

End Sub

Private Sub frmPassword_Click()

    txtPassword = "2468"
    
End Sub

Public Sub OptionsDrills_Click()

    ModIni.ReadGripperStyle
    AppToolType = Drill
    
        
    If AppToolType = HSK Then
        ReDim AutomationStatus(3, 10)
    ElseIf AppToolType = Drill Then
        ReDim AutomationStatus(3, 12)
    ElseIf AppToolType = Round Then
        ReDim AutomationStatus(3, 12)
    End If
    
    
    fMainForm.LabelApp(0).Caption = "Drill"
    fMainForm.txtWorkPiece(7) = "Drill"
    ModIni.WriteAppToolType ("2")
    fMainForm.FrameToolOffset.Caption = "Drill Machine Offset:"
    ''fMainForm.LabelApp(1).Caption = "Diameter : " & CStr(AppDiameter)
    ''ModAutomationStatus.SetAutomationCurrentTool (Drill)
    
    fMainForm.BtnDisplayPocketStatus_Click
    fMainForm.BtnViewPocketsLocations_Click
    
    Call WriteOneCommByte(52, tooltype.Drill)
    fMainForm.SliderAutoSpeed.Value = 50

    fMainForm.TxtSingleToolDiameter.Visible = True
    fMainForm.UpDown1.Visible = True
    
    fMainForm.TablePocketsLocations.Height = 2800
    fMainForm.TablePocketsLocations.Rows = 8
    fMainForm.TablePocketsLocations.Width = 8070
    
    fMainForm.SSFrame10.Visible = True
    fMainForm.SSFrame14.Visible = True
    fMainForm.Label11.Visible = True
    fMainForm.Label30.Visible = True
   
    fMainForm.UpDnDiameter.Visible = True
    fMainForm.TextToolsDiameter.Visible = True
    fMainForm.TextAllPocketsDiameter.Visible = True
    
    fMainForm.TxtViewPocket.Visible = True
    fMainForm.Label18.Visible = True
    fMainForm.UpDnViewPocket.Visible = True
    
    fMainForm.UpDownTestPocket(0).Max = TotalDRILL
    fMainForm.UpDownTestPocket(1).Max = TotalDRILL
    fMainForm.UpDownLoadUnloadPocket.Max = TotalDRILL
    fMainForm.UpDnPocketIndex.Max = TotalDRILL
    fMainForm.UpDownPocket.Max = TotalDRILL
    fMainForm.UpDnViewPocket.Max = TotalDRILL
    fMainForm.UpDown2.Max = TotalDRILL
    fMainForm.UpDown5.Max = 7
    
    ''change status for single pocket
    fMainForm.Label13.Visible = False
    fMainForm.UpDown1.Visible = False
    fMainForm.TxtSingleToolDiameter.Visible = False
    
    fMainForm.textToolOffset(0).Enabled = True
    fMainForm.textToolOffset(1).Enabled = True
    fMainForm.textToolOffset(2).Enabled = False
    fMainForm.textToolOffset(3).Enabled = False
    fMainForm.textToolOffset(4).Enabled = False
    fMainForm.textToolOffset(5).Enabled = False
    
    fMainForm.textToolOffset(0).Enabled = True
    fMainForm.textToolOffset(1).Enabled = True
    fMainForm.textToolOffset(2).Text = " - - "
    fMainForm.textToolOffset(3).Text = " - - "
    fMainForm.textToolOffset(4).Text = " - - "
    fMainForm.textToolOffset(5).Text = " - - "
    
    fMainForm.Label29(0).Enabled = True
    fMainForm.Label29(1).Enabled = True
    fMainForm.Label29(2).Enabled = False
    fMainForm.Label29(3).Enabled = False
    fMainForm.Label29(4).Enabled = False
    fMainForm.Label29(5).Enabled = False
    
    fMainForm.Label32.Enabled = True
    fMainForm.Label33.Enabled = True
    fMainForm.Label34.Enabled = False
    fMainForm.Label35.Enabled = False
    fMainForm.Label36.Enabled = False
    fMainForm.Label37.Enabled = False
    
    fMainForm.CmdSetOffset.Enabled = True
    fMainForm.CmdRestoreDefult.Enabled = True
    
    If AppGripperStyle = 1 Then
        fMainForm.lblOut1(6).Enabled = False
        fMainForm.cmdOutOn(10).Enabled = False
        fMainForm.cmdOutOff(10).Enabled = False
        fMainForm.lblIn1(34).Visible = False
        fMainForm.LabelInput(20).Visible = False
        
    ElseIf AppGripperStyle = 2 Then
        fMainForm.lblOut1(6).Enabled = True
        fMainForm.cmdOutOn(10).Enabled = True
        fMainForm.cmdOutOn(10).Enabled = True
        fMainForm.lblIn1(34).Visible = True
        fMainForm.LabelInput(20).Visible = True
         
    Else
        fMainForm.lblOut1(6).Enabled = False
        fMainForm.cmdOutOn(10).Enabled = False
        fMainForm.cmdOutOff(10).Enabled = False
        fMainForm.lblIn1(34).Visible = False
    End If
    
End Sub

Public Sub OptionsHSK_Click()

    AppToolType = HSK
    
    If AppToolType = HSK Then
        ReDim AutomationStatus(3, 10)
    ElseIf AppToolType = Drill Then
        ReDim AutomationStatus(3, 12)
    ElseIf AppToolType = Round Then
        ReDim AutomationStatus(3, 12)
    End If
    
    
    fMainForm.LabelApp(0).Caption = "HSK"
    fMainForm.txtWorkPiece(7) = "HSK"
    fMainForm.FrameToolOffset.Caption = "HSK Machine Offset:"
    ''ModAutomationStatus.SetAutomationCurrentTool (HSK)
    ModIni.WriteAppToolType ("1")
    
    fMainForm.BtnDisplayPocketStatus_Click
    fMainForm.BtnViewPocketsLocations_Click
    fMainForm.SliderAutoSpeed.Value = 50
    Call WriteOneCommByte(52, tooltype.HSK)
    
    fMainForm.TablePocketsLocations.Height = 3840
    fMainForm.TablePocketsLocations.Rows = 11
    fMainForm.TablePocketsLocations.Width = 8080

    fMainForm.TxtSingleToolDiameter.Visible = False
    fMainForm.UpDown1.Visible = False
    fMainForm.UpDownMyDiameter.Visible = False
    
    fMainForm.SSFrame10.Visible = False
    fMainForm.SSFrame14.Visible = False
    fMainForm.Label13.Visible = False
    fMainForm.Label11.Visible = False
    fMainForm.UpDnDiameter.Visible = False
    fMainForm.TextToolsDiameter.Visible = False
    fMainForm.TextAllPocketsDiameter.Visible = False
    fMainForm.Label30.Visible = False

    fMainForm.TxtViewPocket.Visible = False
    fMainForm.Label18.Visible = False
    fMainForm.UpDnViewPocket.Visible = False
    
    fMainForm.UpDownTestPocket(0).Max = TotalHSK
    fMainForm.UpDownTestPocket(1).Max = TotalHSK
    fMainForm.UpDownLoadUnloadPocket.Max = TotalHSK
    fMainForm.UpDnPocketIndex.Max = TotalHSK
    fMainForm.UpDownPocket.Max = TotalHSK
    fMainForm.UpDnViewPocket.Max = TotalHSK
    fMainForm.UpDown2.Max = TotalHSK
    
    fMainForm.textToolOffset(0).Enabled = False
    fMainForm.textToolOffset(1).Enabled = False
    fMainForm.textToolOffset(2).Enabled = False
    fMainForm.textToolOffset(3).Enabled = False
    fMainForm.textToolOffset(4).Enabled = False
    fMainForm.textToolOffset(5).Enabled = False
    
    fMainForm.textToolOffset(0).Text = " - - "
    fMainForm.textToolOffset(1).Text = " - - "
    fMainForm.textToolOffset(2).Text = " - - "
    fMainForm.textToolOffset(3).Text = " - - "
    fMainForm.textToolOffset(4).Text = " - - "
    fMainForm.textToolOffset(5).Text = " - - "
    
    fMainForm.Label29(0).Enabled = False
    fMainForm.Label29(1).Enabled = False
    fMainForm.Label29(2).Enabled = False
    fMainForm.Label29(3).Enabled = False
    fMainForm.Label29(4).Enabled = False
    fMainForm.Label29(5).Enabled = False
    
    fMainForm.Label32.Enabled = False
    fMainForm.Label33.Enabled = False
    fMainForm.Label34.Enabled = False
    fMainForm.Label35.Enabled = False
    fMainForm.Label36.Enabled = False
    fMainForm.Label37.Enabled = False
    
    fMainForm.CmdSetOffset.Enabled = False
    fMainForm.CmdRestoreDefult.Enabled = False
    
    ''change status for single pocket
    fMainForm.UpDown1.Visible = False
    fMainForm.TxtSingleToolDiameter.Visible = False

    fMainForm.lblOut1(6).Enabled = False
    fMainForm.cmdOutOn(10).Enabled = False
    fMainForm.cmdOutOff(10).Enabled = False
    fMainForm.lblIn1(34).Visible = False
    fMainForm.LabelInput(20).Visible = False
    
End Sub


Public Sub OptionsRound_Click()

    AppToolType = Round
    
        
    If AppToolType = HSK Then
        ReDim AutomationStatus(3, 10)
    ElseIf AppToolType = Drill Then
        ReDim AutomationStatus(3, 12)
    ElseIf AppToolType = Round Then
        ReDim AutomationStatus(3, 12)
    End If
    
    
    
    fMainForm.LabelApp(0) = "Round"
    fMainForm.txtWorkPiece(7) = "Round"
    fMainForm.FrameToolOffset.Caption = "Round Machine Offset:"
    ''ModAutomationStatus.SetAutomationCurrentTool (Round)
    ModIni.WriteAppToolType ("3")
    
    fMainForm.BtnDisplayPocketStatus_Click
    fMainForm.BtnViewPocketsLocations_Click
    fMainForm.SliderAutoSpeed.Value = 50
    Call WriteOneCommByte(52, tooltype.Round)
    
    fMainForm.TablePocketsLocations.Height = 2800
    fMainForm.TablePocketsLocations.Rows = 8
    fMainForm.TablePocketsLocations.Width = 8070
    
    fMainForm.TxtSingleToolDiameter.Visible = True
    fMainForm.UpDown1.Visible = True
    fMainForm.UpDownMyDiameter.Visible = True
    
    fMainForm.SSFrame10.Visible = True
    fMainForm.SSFrame14.Visible = True
    fMainForm.Label11.Visible = True
    fMainForm.Label13.Visible = True
    fMainForm.UpDnDiameter.Visible = True
    fMainForm.TextToolsDiameter.Visible = True
    fMainForm.TextAllPocketsDiameter.Visible = True
    fMainForm.Label30.Visible = True

    fMainForm.TxtViewPocket.Visible = True
    fMainForm.Label18.Visible = True
    fMainForm.UpDnViewPocket.Visible = True
    
    fMainForm.UpDownTestPocket(0).Max = TotalROUND
    fMainForm.UpDownTestPocket(1).Max = TotalROUND
    fMainForm.UpDownLoadUnloadPocket.Max = TotalROUND
    fMainForm.UpDnPocketIndex.Max = TotalROUND
    fMainForm.UpDownPocket.Max = TotalROUND
    fMainForm.UpDnViewPocket.Max = TotalROUND
    fMainForm.UpDown2.Max = TotalROUND
    fMainForm.UpDown5.Max = TotalROUND
    
    fMainForm.textToolOffset(0).Enabled = False
    fMainForm.textToolOffset(1).Enabled = False
    fMainForm.textToolOffset(2).Enabled = True
    fMainForm.textToolOffset(3).Enabled = True
    fMainForm.textToolOffset(4).Enabled = True
    fMainForm.textToolOffset(5).Enabled = True
    
    fMainForm.textToolOffset(0).Text = " - - "
    fMainForm.textToolOffset(1).Text = " - - "
    ''fMainForm.textToolOffset(2).Text = " - - "
    ''fMainForm.textToolOffset(3).Text = " - - "
    ''fMainForm.textToolOffset(4).Text = " - - "
    ''fMainForm.textToolOffset(5).Text = " - - "
    
    fMainForm.Label29(0).Enabled = False
    fMainForm.Label29(1).Enabled = False
    fMainForm.Label29(2).Enabled = True
    fMainForm.Label29(3).Enabled = True
    fMainForm.Label29(4).Enabled = True
    fMainForm.Label29(5).Enabled = True
    
    fMainForm.Label32.Enabled = False
    fMainForm.Label33.Enabled = False
    fMainForm.Label34.Enabled = True
    fMainForm.Label35.Enabled = True
    fMainForm.Label36.Enabled = True
    fMainForm.Label37.Enabled = True
    
    fMainForm.CmdSetOffset.Enabled = True
    fMainForm.CmdRestoreDefult.Enabled = True

    ''change status for single pocket
    fMainForm.UpDown1.Visible = False
    fMainForm.TxtSingleToolDiameter.Visible = False
    
    fMainForm.lblOut1(6).Enabled = False
    fMainForm.cmdOutOn(10).Enabled = False
    fMainForm.cmdOutOff(10).Enabled = False
    fMainForm.lblIn1(34).Visible = False
    fMainForm.LabelInput(20).Visible = False

     
End Sub

Private Sub OptionsTab_Click(PreviousTab As Integer)

    IntTab = OptionsTab.Tab + 1
    
    If IntTab = 1 Then ''information
        Call frmOptions.CmdReadInfo_Click
        
    ElseIf IntTab = 2 Then ''back up
    
    ElseIf IntTab = 3 Then ''configuration
    
    ElseIf IntTab = 4 Then ''more options
        TextMoreOptions.Text = ""
        
    ElseIf IntTab = 5 Then ''gripper
        If AppToolType = HSK Then
            SelectGripper(5).Value = True
        ElseIf AppToolType = Round Then
            SelectGripper(6).Value = True
        End If
    
    ElseIf IntTab = 6 Then ''all locations
        Call CmdAllLocationsTableDisplay_Click
        
    End If
    
End Sub

Private Sub SelectGripper_Click(Index As Integer)
    
    LabelSelectGripper.ForeColor = &H8000&
    TextSelectGripper.Text = CStr(Index)
    LabelSelectGripper.Caption = "...Select gripper..."

End Sub


