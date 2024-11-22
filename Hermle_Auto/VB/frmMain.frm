VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "threed32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm mdiMain 
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "Hermle"
   ClientHeight    =   11100
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   15195
   Icon            =   "frmMain.frx":0000
   LinkMode        =   1  'Source
   LinkTopic       =   "MDIForm1"
   ScrollBars      =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   11295
      Left            =   0
      ScaleHeight     =   11235
      ScaleWidth      =   15135
      TabIndex        =   1
      Top             =   0
      Width           =   15195
      Begin VB.Timer tmrQuesyStatus 
         Interval        =   500
         Left            =   10800
         Top             =   120
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   9945
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   16777215
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   12632256
         UseMaskColor    =   0   'False
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   17
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":27CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2AE6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F23
               Key             =   ""
               Object.Tag             =   "drill"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8904
               Key             =   ""
               Object.Tag             =   "hsk"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12C04
               Key             =   ""
               Object.Tag             =   "Manu"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":15C56
               Key             =   ""
               Object.Tag             =   "Semi"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":18CA8
               Key             =   ""
               Object.Tag             =   "Auto"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1A7FA
               Key             =   ""
               Object.Tag             =   "TwinCat"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":27678
               Key             =   ""
               Object.Tag             =   "Exit"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28702
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":29714
               Key             =   ""
               Object.Tag             =   "ESrelease"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A5EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":729A0
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7366C
               Key             =   ""
               Object.Tag             =   "Shafir"
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A36BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A5B80
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar TopToolBar 
         Height          =   690
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   14910
         _ExtentX        =   26300
         _ExtentY        =   1217
         ButtonWidth     =   2090
         ButtonHeight    =   1164
         AllowCustomize  =   0   'False
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         DisabledImageList=   "ImageList1"
         HotImageList    =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Shafir"
               Key             =   "Shafir"
               ImageIndex      =   15
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Automat"
               Key             =   "Automat"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Semi"
               Key             =   "Semi"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Manual"
               Key             =   "Manual"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "R.E.Stop"
               Key             =   "R.E.Stop"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Options"
               Key             =   "Options"
               ImageIndex      =   3
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Exit"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Communication"
               Key             =   "Communication"
               ImageIndex      =   2
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   5
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Key             =   "Exit Hermle"
                     Text            =   "Exit Hermle"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
                  BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit"
               Key             =   "Exit"
               ImageIndex      =   10
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin VB.Label Label14 
            Caption         =   "Label14"
            Height          =   15
            Left            =   11925
            TabIndex        =   3
            Top             =   450
            Width           =   15
         End
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   6105
         Left            =   360
         TabIndex        =   4
         Top             =   3240
         Width           =   12600
         _ExtentX        =   22225
         _ExtentY        =   10769
         _Version        =   393216
         Tabs            =   10
         Tab             =   1
         TabsPerRow      =   10
         TabHeight       =   882
         BackColor       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Automat"
         TabPicture(0)   =   "frmMain.frx":A5F19
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "CmdSysOperation(3)"
         Tab(0).Control(1)=   "SSFrame6"
         Tab(0).Control(2)=   "SSFrame1(4)"
         Tab(0).Control(3)=   "SSFrame1(5)"
         Tab(0).Control(4)=   "SSFrame1(8)"
         Tab(0).ControlCount=   5
         TabCaption(1)   =   "Tools"
         TabPicture(1)   =   "frmMain.frx":A5F35
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "SSFrame3"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "SSFrame1(9)"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "SSFrame2"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).Control(3)=   "FrameSimulator"
         Tab(1).Control(3).Enabled=   0   'False
         Tab(1).ControlCount=   4
         TabCaption(2)   =   "Work  Piece"
         TabPicture(2)   =   "frmMain.frx":A5F51
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "fraMain"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Manual"
         TabPicture(3)   =   "frmMain.frx":A5F6D
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "SSFrame1(11)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "Pocket Status"
         TabPicture(4)   =   "frmMain.frx":A5F89
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "Frame6"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "Teach"
         TabPicture(5)   =   "frmMain.frx":A5FA5
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "SSTab3"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Diagnostic"
         TabPicture(6)   =   "frmMain.frx":A5FC1
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "Frame4"
         Tab(6).Control(1)=   "Frame2"
         Tab(6).Control(2)=   "Frame7"
         Tab(6).ControlCount=   3
         TabCaption(7)   =   "Operation"
         TabPicture(7)   =   "frmMain.frx":A5FDD
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "Frame11"
         Tab(7).Control(1)=   "FrameManualOperations"
         Tab(7).Control(2)=   "SSFrame1(1200)"
         Tab(7).ControlCount=   3
         TabCaption(8)   =   "Tests"
         TabPicture(8)   =   "frmMain.frx":A5FF9
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "SSTabTests"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "Info"
         TabPicture(9)   =   "frmMain.frx":A6015
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "SSTab5"
         Tab(9).ControlCount=   1
         Begin VB.Frame Frame11 
            Caption         =   "Not used "
            Height          =   2535
            Left            =   -67575
            TabIndex        =   344
            Top             =   5085
            Visible         =   0   'False
            Width           =   4155
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   20
               Left            =   1890
               Style           =   1  'Graphical
               TabIndex        =   357
               Top             =   1935
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   20
               Left            =   2790
               Style           =   1  'Graphical
               TabIndex        =   356
               Top             =   1935
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   5
               Left            =   2805
               Style           =   1  'Graphical
               TabIndex        =   351
               Top             =   945
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   5
               Left            =   1905
               Style           =   1  'Graphical
               TabIndex        =   350
               Top             =   945
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   8
               Left            =   2775
               Style           =   1  'Graphical
               TabIndex        =   348
               Top             =   1440
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   8
               Left            =   1890
               Style           =   1  'Graphical
               TabIndex        =   347
               Top             =   1440
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   2
               Left            =   2790
               Style           =   1  'Graphical
               TabIndex        =   346
               Top             =   405
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   2
               Left            =   1935
               Style           =   1  'Graphical
               TabIndex        =   345
               Top             =   405
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Safty Relay"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   5
               Left            =   225
               TabIndex        =   358
               Top             =   1890
               Width           =   1620
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Close Gripper"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   4
               Left            =   225
               TabIndex        =   352
               Top             =   900
               Width           =   1665
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Kiosk Valve close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   1
               Left            =   225
               TabIndex        =   349
               Top             =   405
               Width           =   1665
            End
         End
         Begin TabDlg.SSTab SSTab5 
            Height          =   5265
            Left            =   -74760
            TabIndex        =   327
            Top             =   690
            Width           =   12180
            _ExtentX        =   21484
            _ExtentY        =   9287
            _Version        =   393216
            TabHeight       =   420
            TabCaption(0)   =   "Robot Information"
            TabPicture(0)   =   "frmMain.frx":A6031
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSFrame13"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Loger"
            TabPicture(1)   =   "frmMain.frx":A604D
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame8(1)"
            Tab(1).Control(1)=   "Frame8(0)"
            Tab(1).Control(2)=   "lstReceivedFromRobot"
            Tab(1).Control(3)=   "lstTransmitToRobot"
            Tab(1).Control(4)=   "imgQuesryTimer"
            Tab(1).Control(5)=   "Label39"
            Tab(1).Control(6)=   "lblUpdateTime(2)"
            Tab(1).Control(7)=   "lblUpdateTime(1)"
            Tab(1).Control(8)=   "lblUpdateTime(0)"
            Tab(1).Control(9)=   "lblTimerQuery"
            Tab(1).Control(10)=   "lblTimerStatus"
            Tab(1).Control(11)=   "lblDllWriteError"
            Tab(1).Control(12)=   "Label38"
            Tab(1).Control(13)=   "lblDllReadError"
            Tab(1).Control(14)=   "Label2"
            Tab(1).Control(15)=   "imgPlcUpdate(1)"
            Tab(1).Control(16)=   "lblCycleMessage(1)"
            Tab(1).Control(17)=   "lblCycleMessage(0)"
            Tab(1).Control(18)=   "imgPlcUpdate(0)"
            Tab(1).ControlCount=   19
            TabCaption(2)   =   "HMI information"
            TabPicture(2)   =   "frmMain.frx":A6069
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Label24"
            Tab(2).Control(1)=   "ListHMIInfo"
            Tab(2).Control(2)=   "CmdClearHMIList"
            Tab(2).Control(3)=   "CheckListHMI"
            Tab(2).Control(4)=   "TextListHmiCounter"
            Tab(2).ControlCount=   5
            Begin VB.TextBox TextListHmiCounter 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   -73800
               TabIndex        =   386
               Text            =   "32000"
               Top             =   2040
               Width           =   975
            End
            Begin VB.CheckBox CheckListHMI 
               BackColor       =   &H8000000C&
               Caption         =   "display  HMI information"
               Height          =   375
               Left            =   -74880
               TabIndex        =   381
               Top             =   1560
               Value           =   1  'Checked
               Width           =   2150
            End
            Begin VB.CommandButton CmdClearHMIList 
               Caption         =   "Clear HMI  List"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   615
               Left            =   -74880
               TabIndex        =   380
               Top             =   840
               Width           =   2150
            End
            Begin VB.ListBox ListHMIInfo 
               Height          =   3570
               Left            =   -72600
               TabIndex        =   379
               Top             =   840
               Width           =   7695
            End
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   240
               Index           =   1
               Left            =   -63444
               MouseIcon       =   "frmMain.frx":A6085
               TabIndex        =   334
               Top             =   1116
               Width           =   345
               Begin VB.Image imgGrLampSource 
                  Height          =   240
                  Index           =   1
                  Left            =   0
                  Picture         =   "frmMain.frx":A638F
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   255
               End
               Begin VB.Image imgRedLampSource 
                  Height          =   240
                  Index           =   1
                  Left            =   0
                  Picture         =   "frmMain.frx":A64D9
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   252
               End
            End
            Begin VB.Frame Frame8 
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   240
               Index           =   0
               Left            =   -63444
               MouseIcon       =   "frmMain.frx":A6623
               TabIndex        =   333
               Top             =   756
               Width           =   345
               Begin VB.Image imgRedLampSource 
                  Height          =   240
                  Index           =   0
                  Left            =   0
                  Picture         =   "frmMain.frx":A692D
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   252
               End
               Begin VB.Image imgGrLampSource 
                  Height          =   240
                  Index           =   0
                  Left            =   0
                  Picture         =   "frmMain.frx":A6A77
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.ListBox lstReceivedFromRobot 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   3735
               ItemData        =   "frmMain.frx":A6BC1
               Left            =   -74640
               List            =   "frmMain.frx":A6BC3
               TabIndex        =   330
               Top             =   720
               Width           =   3924
            End
            Begin VB.ListBox lstTransmitToRobot 
               Appearance      =   0  'Flat
               BackColor       =   &H8000000F&
               Height          =   3735
               ItemData        =   "frmMain.frx":A6BC5
               Left            =   -70680
               List            =   "frmMain.frx":A6BC7
               TabIndex        =   329
               Top             =   720
               Width           =   3924
            End
            Begin Threed.SSFrame SSFrame13 
               Height          =   3792
               Left            =   180
               TabIndex        =   328
               Top             =   504
               Width           =   11820
               _Version        =   65536
               _ExtentX        =   20849
               _ExtentY        =   6689
               _StockProps     =   14
               Caption         =   "Cycle Info"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               Font3D          =   4
               Begin VB.CommandButton CmdClearRobotList 
                  Caption         =   "Clear Robot List"
                  Height          =   615
                  Left            =   480
                  TabIndex        =   406
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.ListBox ListRobotInfo 
                  Height          =   2790
                  Left            =   3000
                  TabIndex        =   405
                  Top             =   480
                  Width           =   8055
               End
            End
            Begin VB.Label Label24 
               Caption         =   "Lines : "
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Left            =   -74640
               TabIndex        =   387
               Top             =   2040
               Width           =   735
            End
            Begin VB.Image imgQuesryTimer 
               Height          =   240
               Left            =   -63768
               Picture         =   "frmMain.frx":A6BC9
               Stretch         =   -1  'True
               Top             =   2160
               Width           =   255
            End
            Begin VB.Label Label39 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Robot Query :"
               ForeColor       =   &H80000008&
               Height          =   345
               Left            =   -66350
               TabIndex        =   385
               Top             =   2160
               Width           =   1490
            End
            Begin VB.Label lblUpdateTime 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               DataField       =   "InsQToPress"
               DataMember      =   "Simot"
               DataSource      =   "DorstDEnv"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Index           =   2
               Left            =   -64800
               TabIndex        =   384
               Top             =   2160
               Width           =   855
            End
            Begin VB.Label lblUpdateTime 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   1
               Left            =   -64800
               TabIndex        =   342
               Top             =   1035
               Width           =   855
            End
            Begin VB.Label lblUpdateTime 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               DataField       =   "InsQToPress"
               DataMember      =   "Simot"
               DataSource      =   "DorstDEnv"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Index           =   0
               Left            =   -64800
               TabIndex        =   341
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblTimerQuery 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "tmrRobotQuery"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   -66350
               TabIndex        =   340
               Top             =   720
               Width           =   1490
            End
            Begin VB.Label lblTimerStatus 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "tmrRobotStatus"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   300
               Left            =   -66350
               TabIndex        =   339
               Top             =   1035
               Width           =   1490
            End
            Begin VB.Label lblDllWriteError 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   -64800
               TabIndex        =   338
               Top             =   1710
               Width           =   855
            End
            Begin VB.Label Label38 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dll write error"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   -66350
               TabIndex        =   337
               Top             =   1710
               Width           =   1490
            End
            Begin VB.Label lblDllReadError 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00C0C0C0&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "0"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   -64800
               TabIndex        =   336
               Top             =   1350
               Width           =   855
            End
            Begin VB.Label Label2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dll read error"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   345
               Left            =   -66350
               TabIndex        =   335
               Top             =   1350
               Width           =   1490
            End
            Begin VB.Image imgPlcUpdate 
               Height          =   240
               Index           =   1
               Left            =   -63768
               Picture         =   "frmMain.frx":A6D13
               Stretch         =   -1  'True
               Top             =   1128
               Width           =   252
            End
            Begin VB.Label lblCycleMessage 
               Alignment       =   2  'Center
               Caption         =   "DLL Write"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   -69636
               TabIndex        =   332
               Top             =   432
               Width           =   1860
            End
            Begin VB.Label lblCycleMessage 
               Alignment       =   2  'Center
               Caption         =   "DLL read"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   -73956
               TabIndex        =   331
               Top             =   432
               Width           =   1860
            End
            Begin VB.Image imgPlcUpdate 
               Height          =   240
               Index           =   0
               Left            =   -63768
               Picture         =   "frmMain.frx":A6E5D
               Stretch         =   -1  'True
               Top             =   768
               Width           =   252
            End
         End
         Begin VB.Frame FrameSimulator 
            Caption         =   "Simulator"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   177
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4815
            Left            =   11280
            TabIndex        =   289
            Top             =   720
            Width           =   1215
            Begin VB.CommandButton CmdSimChuckToPocket 
               Caption         =   "Chuck To Pocket"
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
               Left            =   120
               TabIndex        =   296
               Top             =   2520
               Width           =   975
            End
            Begin VB.CommandButton CmdSimPocketToChuck 
               Caption         =   "Pocket To Chuck"
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
               Left            =   120
               TabIndex        =   295
               Top             =   1800
               Width           =   975
            End
            Begin VB.TextBox TextSimulatorIndex 
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   294
               Top             =   4320
               Width           =   975
            End
            Begin VB.TextBox TextSimulatorIndex 
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   293
               Top             =   3840
               Width           =   975
            End
            Begin VB.TextBox TextSimulatorIndex 
               Height          =   375
               Index           =   0
               Left            =   120
               TabIndex        =   292
               Top             =   3360
               Width           =   975
            End
            Begin VB.CommandButton CmdSimPocketToKiosk 
               Caption         =   "Pocket To Kiosk"
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
               Left            =   120
               TabIndex        =   291
               Top             =   1080
               Width           =   975
            End
            Begin VB.CommandButton CmdSimKioskToPocket 
               Caption         =   "Kiosk To Pocket"
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
               Left            =   120
               TabIndex        =   290
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.Frame FrameManualOperations 
            Caption         =   "Manual Operation"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   4230
            Left            =   -67575
            TabIndex        =   121
            Top             =   765
            Width           =   4935
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Open"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   10
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   354
               Top             =   3660
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   10
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   353
               Top             =   3660
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   9
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   133
               Top             =   3100
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   7
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   132
               Top             =   2573
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   6
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   131
               Top             =   2046
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   4
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   130
               Top             =   1530
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Off"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   3
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   129
               Top             =   992
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   9
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   128
               Top             =   3100
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   7
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   127
               Top             =   2573
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   6
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   126
               Top             =   2046
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Open"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   4
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   125
               Top             =   1530
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "On"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   3
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   124
               Top             =   992
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOff 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Open"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   1
               Left            =   2580
               Style           =   1  'Graphical
               TabIndex        =   123
               Top             =   465
               Width           =   855
            End
            Begin VB.CommandButton cmdOutOn 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   400
               Index           =   1
               Left            =   3480
               Style           =   1  'Graphical
               TabIndex        =   122
               Top             =   465
               Width           =   855
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Gripper 2 Open/Close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   6
               Left            =   195
               TabIndex        =   355
               Top             =   3660
               Width           =   2340
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Gripper Open/Close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   9
               Left            =   195
               TabIndex        =   139
               Top             =   1519
               Width           =   2340
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Door Interlocck"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   8
               Left            =   195
               TabIndex        =   138
               Top             =   2573
               Width           =   2340
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cell Light"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   7
               Left            =   195
               TabIndex        =   137
               Top             =   2046
               Width           =   2340
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Indicator  User Ack"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   3
               Left            =   195
               TabIndex        =   136
               Top             =   992
               Width           =   2340
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Interlock Hermle ByPass"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   2
               Left            =   195
               TabIndex        =   135
               Top             =   3100
               Width           =   2340
            End
            Begin VB.Label lblOut1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Kiosk Valve Open/Close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Index           =   0
               Left            =   195
               TabIndex        =   134
               Top             =   465
               Width           =   2340
            End
         End
         Begin VB.Frame fraMain 
            Height          =   5445
            Left            =   -74856
            TabIndex        =   97
            Top             =   585
            Width           =   12315
            Begin VB.CommandButton BtnDisplayWPTable 
               Caption         =   "Refresh Table"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   915
               Left            =   9120
               Picture         =   "frmMain.frx":A6FA7
               Style           =   1  'Graphical
               TabIndex        =   98
               Top             =   4410
               Width           =   2535
            End
            Begin TabDlg.SSTab SSTab1 
               CausesValidation=   0   'False
               Height          =   5055
               Left            =   135
               TabIndex        =   99
               Top             =   240
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   8916
               _Version        =   393216
               TabHeight       =   520
               TabCaption(0)   =   "Add New WorkPiece"
               TabPicture(0)   =   "frmMain.frx":A7330
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "SSFrame1(10)"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).ControlCount=   1
               TabCaption(1)   =   "WP Options"
               TabPicture(1)   =   "frmMain.frx":A734C
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Frame5"
               Tab(1).Control(1)=   "Frame9"
               Tab(1).Control(2)=   "TextLineNumber"
               Tab(1).Control(3)=   "UpDnLine"
               Tab(1).Control(4)=   "Label4"
               Tab(1).ControlCount=   5
               TabCaption(2)   =   "Offsets"
               TabPicture(2)   =   "frmMain.frx":A7368
               Tab(2).ControlEnabled=   0   'False
               Tab(2).Control(0)=   "FrameToolOffset"
               Tab(2).ControlCount=   1
               Begin VB.Frame Frame5 
                  Caption         =   "Priority"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   177
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   2655
                  Left            =   -74805
                  TabIndex        =   104
                  Top             =   360
                  Width           =   1530
                  Begin VB.CommandButton cmdMove 
                     Caption         =   "Up"
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1000
                     Index           =   6
                     Left            =   120
                     MouseIcon       =   "frmMain.frx":A7384
                     MousePointer    =   99  'Custom
                     Picture         =   "frmMain.frx":A74D6
                     Style           =   1  'Graphical
                     TabIndex        =   106
                     ToolTipText     =   "(+)"
                     Top             =   360
                     Width           =   1290
                  End
                  Begin VB.CommandButton cmdMove 
                     Caption         =   "Down"
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1000
                     Index           =   7
                     Left            =   120
                     MouseIcon       =   "frmMain.frx":A7918
                     MousePointer    =   99  'Custom
                     Picture         =   "frmMain.frx":A7A6A
                     Style           =   1  'Graphical
                     TabIndex        =   105
                     ToolTipText     =   "(-)"
                     Top             =   1440
                     Width           =   1290
                  End
               End
               Begin VB.Frame Frame9 
                  Caption         =   "Delete Line"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   177
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   1530
                  Left            =   -73200
                  TabIndex        =   102
                  Top             =   360
                  Width           =   1728
                  Begin VB.CommandButton BtnDelOrder 
                     Caption         =   "Delete"
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   1000
                     Left            =   120
                     Picture         =   "frmMain.frx":A7EAC
                     Style           =   1  'Graphical
                     TabIndex        =   103
                     Top             =   360
                     Width           =   1500
                  End
               End
               Begin VB.TextBox TextLineNumber 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   -70668
                  TabIndex        =   101
                  Text            =   "1"
                  Top             =   990
                  Width           =   700
               End
               Begin MSComCtl2.UpDown UpDnLine 
                  Height          =   735
                  Left            =   -69960
                  TabIndex        =   100
                  Top             =   960
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1296
                  _Version        =   393216
                  Value           =   1
                  Max             =   50
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin Threed.SSFrame SSFrame1 
                  Height          =   4530
                  Index           =   10
                  Left            =   210
                  TabIndex        =   107
                  Top             =   360
                  Width           =   5265
                  _Version        =   65536
                  _ExtentX        =   9287
                  _ExtentY        =   7990
                  _StockProps     =   14
                  Caption         =   "Add/Edit Work Piece"
                  ForeColor       =   64
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Font3D          =   1
                  ShadowStyle     =   1
                  Begin VB.TextBox txtWorkPiece 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     Enabled         =   0   'False
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   480
                     Index           =   7
                     Left            =   2700
                     MaxLength       =   6
                     TabIndex        =   287
                     Top             =   600
                     Width           =   2100
                  End
                  Begin VB.TextBox txtWorkPiece 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   480
                     Index           =   6
                     Left            =   2700
                     MaxLength       =   2
                     TabIndex        =   112
                     Top             =   3100
                     Width           =   2100
                  End
                  Begin VB.TextBox txtWorkPiece 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   480
                     Index           =   5
                     Left            =   2700
                     MaxLength       =   3
                     TabIndex        =   111
                     Top             =   2600
                     Width           =   2100
                  End
                  Begin VB.TextBox txtWorkPiece 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   480
                     Index           =   4
                     Left            =   2700
                     MaxLength       =   3
                     TabIndex        =   110
                     Top             =   2100
                     Width           =   2100
                  End
                  Begin VB.TextBox txtWorkPiece 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   480
                     Index           =   3
                     Left            =   2700
                     MaxLength       =   9
                     TabIndex        =   109
                     Top             =   1600
                     Width           =   2100
                  End
                  Begin VB.TextBox txtWorkPiece 
                     Alignment       =   2  'Center
                     BackColor       =   &H00FFFFC0&
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   14.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   480
                     Index           =   2
                     Left            =   2700
                     MaxLength       =   3
                     TabIndex        =   108
                     Top             =   1125
                     Width           =   2100
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   480
                     Index           =   8
                     Left            =   330
                     TabIndex        =   113
                     Top             =   3100
                     Width           =   2100
                     _Version        =   65536
                     _ExtentX        =   3704
                     _ExtentY        =   847
                     _StockProps     =   15
                     Caption         =   "Line Number"
                     BackColor       =   12640511
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BevelOuter      =   1
                     BevelInner      =   1
                     Font3D          =   2
                  End
                  Begin Threed.SSCommand BtnAddWrokPiece 
                     Height          =   660
                     Left            =   1200
                     TabIndex        =   114
                     Top             =   3720
                     Width           =   2565
                     _Version        =   65536
                     _ExtentX        =   4530
                     _ExtentY        =   1164
                     _StockProps     =   78
                     Caption         =   "Add Work Piece"
                     ForeColor       =   12582912
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
                     Picture         =   "frmMain.frx":A8926
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   480
                     Index           =   4
                     Left            =   330
                     TabIndex        =   115
                     Top             =   1600
                     Width           =   2100
                     _Version        =   65536
                     _ExtentX        =   3704
                     _ExtentY        =   847
                     _StockProps     =   15
                     Caption         =   "NC Program"
                     ForeColor       =   -2147483630
                     BackColor       =   12640511
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BevelOuter      =   1
                     BevelInner      =   1
                     Font3D          =   2
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   480
                     Index           =   5
                     Left            =   330
                     TabIndex        =   116
                     Top             =   1100
                     Width           =   2100
                     _Version        =   65536
                     _ExtentX        =   3704
                     _ExtentY        =   847
                     _StockProps     =   15
                     Caption         =   "Work Piece"
                     ForeColor       =   -2147483630
                     BackColor       =   12640511
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BevelOuter      =   1
                     BevelInner      =   1
                     Font3D          =   2
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   480
                     Index           =   6
                     Left            =   330
                     TabIndex        =   117
                     Top             =   2100
                     Width           =   2100
                     _Version        =   65536
                     _ExtentX        =   3704
                     _ExtentY        =   847
                     _StockProps     =   15
                     Caption         =   "Tool Diameter"
                     ForeColor       =   -2147483630
                     BackColor       =   12640511
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BevelOuter      =   1
                     BevelInner      =   1
                     Font3D          =   2
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   480
                     Index           =   7
                     Left            =   330
                     TabIndex        =   118
                     Top             =   2600
                     Width           =   2100
                     _Version        =   65536
                     _ExtentX        =   3704
                     _ExtentY        =   847
                     _StockProps     =   15
                     Caption         =   "Tool Amount"
                     ForeColor       =   -2147483630
                     BackColor       =   12640511
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BevelOuter      =   1
                     BevelInner      =   1
                     Font3D          =   2
                  End
                  Begin Threed.SSPanel SSPanel 
                     Height          =   480
                     Index           =   2
                     Left            =   330
                     TabIndex        =   283
                     Top             =   600
                     Width           =   2100
                     _Version        =   65536
                     _ExtentX        =   3704
                     _ExtentY        =   847
                     _StockProps     =   15
                     Caption         =   "Tool Type"
                     BackColor       =   12640511
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     BevelOuter      =   1
                     BevelInner      =   1
                     Font3D          =   2
                  End
                  Begin VB.Image ImgFdBkWPiece 
                     Height          =   360
                     Left            =   4080
                     Picture         =   "frmMain.frx":A8942
                     Top             =   3960
                     Width           =   360
                  End
               End
               Begin Threed.SSFrame FrameToolOffset 
                  Height          =   4455
                  Left            =   -74760
                  TabIndex        =   305
                  Top             =   360
                  Width           =   5295
                  _Version        =   65536
                  _ExtentX        =   9340
                  _ExtentY        =   7858
                  _StockProps     =   14
                  Caption         =   "Tool Offset"
                  ForeColor       =   128
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ShadowStyle     =   1
                  Begin VB.TextBox textToolOffset 
                     Alignment       =   2  'Center
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
                     Index           =   5
                     Left            =   2625
                     TabIndex        =   323
                     Text            =   "1"
                     Top             =   2880
                     Width           =   1000
                  End
                  Begin VB.CommandButton CmdRestoreDefult 
                     Caption         =   "Read Values"
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   975
                     Left            =   120
                     TabIndex        =   322
                     Top             =   3360
                     Width           =   2350
                  End
                  Begin VB.TextBox textToolOffset 
                     Alignment       =   2  'Center
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
                     Index           =   4
                     Left            =   2625
                     TabIndex        =   319
                     Text            =   "1"
                     Top             =   2400
                     Width           =   1000
                  End
                  Begin VB.TextBox textToolOffset 
                     Alignment       =   2  'Center
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
                     Index           =   3
                     Left            =   2625
                     TabIndex        =   314
                     Text            =   "247"
                     Top             =   1920
                     Width           =   1000
                  End
                  Begin VB.TextBox textToolOffset 
                     Alignment       =   2  'Center
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
                     Index           =   1
                     Left            =   2625
                     TabIndex        =   309
                     Text            =   "20"
                     Top             =   960
                     Width           =   1000
                  End
                  Begin VB.TextBox textToolOffset 
                     Alignment       =   2  'Center
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
                     Index           =   0
                     Left            =   2625
                     TabIndex        =   308
                     Text            =   "10"
                     Top             =   480
                     Width           =   1000
                  End
                  Begin VB.CommandButton CmdSetOffset 
                     Caption         =   "Send  Offsets"
                     BeginProperty Font 
                        Name            =   "Microsoft Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   975
                     Left            =   2640
                     TabIndex        =   307
                     Top             =   3360
                     Width           =   2350
                  End
                  Begin VB.TextBox textToolOffset 
                     Alignment       =   2  'Center
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
                     Index           =   2
                     Left            =   2625
                     TabIndex        =   306
                     Text            =   "135"
                     Top             =   1440
                     Width           =   1000
                  End
                  Begin VB.Label Label37 
                     Caption         =   "(D66)"
                     Height          =   255
                     Left            =   3720
                     TabIndex        =   325
                     Top             =   2925
                     Width           =   900
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Kiosk Stopper Offset"
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
                     Index           =   5
                     Left            =   600
                     TabIndex        =   324
                     Top             =   2880
                     Width           =   1950
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Pocket Stopper offset"
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
                     Index           =   4
                     Left            =   600
                     TabIndex        =   321
                     Top             =   2400
                     Width           =   1950
                  End
                  Begin VB.Label Label36 
                     Caption         =   "(D67)"
                     Height          =   255
                     Left            =   3720
                     TabIndex        =   320
                     Top             =   2445
                     Width           =   900
                  End
                  Begin VB.Label Label35 
                     Caption         =   "(D69)"
                     Height          =   255
                     Left            =   3705
                     TabIndex        =   318
                     Top             =   1965
                     Width           =   900
                  End
                  Begin VB.Label Label34 
                     Caption         =   "(D68)"
                     Height          =   255
                     Left            =   3705
                     TabIndex        =   317
                     Top             =   1485
                     Width           =   900
                  End
                  Begin VB.Label Label33 
                     Caption         =   "(D12)"
                     Height          =   255
                     Left            =   3705
                     TabIndex        =   316
                     Top             =   1005
                     Width           =   900
                  End
                  Begin VB.Label Label32 
                     Caption         =   "(D11)"
                     Height          =   255
                     Left            =   3705
                     TabIndex        =   315
                     Top             =   525
                     Width           =   900
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Chuck Depth"
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
                     Index           =   3
                     Left            =   1100
                     TabIndex        =   313
                     Top             =   1920
                     Width           =   1350
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Above Chuck"
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
                     Index           =   1
                     Left            =   1100
                     TabIndex        =   312
                     Top             =   960
                     Width           =   1350
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Above Pocket"
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
                     Index           =   0
                     Left            =   1100
                     TabIndex        =   311
                     Top             =   480
                     Width           =   1350
                  End
                  Begin VB.Label Label29 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Chuck Stopper"
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
                     Index           =   2
                     Left            =   1100
                     TabIndex        =   310
                     Top             =   1440
                     Width           =   1350
                  End
               End
               Begin VB.Label Label4 
                  Caption         =   "Line "
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   300
                  Left            =   -70545
                  TabIndex        =   119
                  Top             =   750
                  Width           =   540
               End
            End
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
               Height          =   4020
               Left            =   6120
               TabIndex        =   120
               Top             =   240
               Width           =   6000
               _ExtentX        =   10583
               _ExtentY        =   7091
               _Version        =   393216
               Rows            =   51
               Cols            =   6
               RowHeightMin    =   396
               BackColor       =   16777215
               BackColorSel    =   12632319
               ScrollTrack     =   -1  'True
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
         End
         Begin VB.Frame Frame6 
            Height          =   5205
            Left            =   -74880
            TabIndex        =   75
            Top             =   720
            Width           =   12192
            Begin Threed.SSCommand CmdResetStatusTable 
               Height          =   465
               Left            =   135
               TabIndex        =   343
               Top             =   4590
               Width           =   3435
               _Version        =   65536
               _ExtentX        =   6059
               _ExtentY        =   820
               _StockProps     =   78
               Caption         =   "Reset Table"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
            End
            Begin VB.Frame Frame1 
               Caption         =   "Work Piece Pockets"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   2175
               Left            =   7512
               TabIndex        =   89
               Top             =   120
               Width           =   4560
               Begin VB.TextBox textChangeWorkPiece 
                  Alignment       =   2  'Center
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  Left            =   2560
                  MaxLength       =   2
                  TabIndex        =   92
                  Top             =   900
                  Width           =   1770
               End
               Begin VB.ComboBox ComboAllStatus 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  ItemData        =   "frmMain.frx":AADF4
                  Left            =   2565
                  List            =   "frmMain.frx":AADF6
                  Style           =   2  'Dropdown List
                  TabIndex        =   91
                  Top             =   360
                  Width           =   1770
               End
               Begin Threed.SSCommand BtnAllStatus 
                  Height          =   510
                  Left            =   90
                  TabIndex        =   90
                  Top             =   1485
                  Width           =   4245
                  _Version        =   65536
                  _ExtentX        =   7488
                  _ExtentY        =   900
                  _StockProps     =   78
                  Caption         =   "Change Status for WorkPiece Pockets"
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
               Begin VB.Label Label8 
                  Caption         =   "Enter WorkPiece (1-50)"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   180
                  TabIndex        =   94
                  Top             =   990
                  Width           =   2265
               End
               Begin VB.Label Label7 
                  Caption         =   "Pick Status From List"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   180
                  TabIndex        =   93
                  Top             =   450
                  Width           =   2220
               End
            End
            Begin VB.TextBox TextShelf 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Left            =   6300
               TabIndex        =   88
               Text            =   "1"
               Top             =   2475
               Width           =   700
            End
            Begin VB.Frame Frame3 
               Caption         =   "Single Pocket"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   2715
               Left            =   7512
               TabIndex        =   78
               Top             =   2376
               Width           =   4560
               Begin VB.ComboBox ComboSingleStatus 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  ItemData        =   "frmMain.frx":AADF8
                  Left            =   2250
                  List            =   "frmMain.frx":AADFA
                  TabIndex        =   82
                  Text            =   "Single Status"
                  Top             =   360
                  Width           =   2085
               End
               Begin VB.TextBox TxtSinglePocketNumber 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   855
                  TabIndex        =   80
                  Text            =   "101"
                  Top             =   1185
                  Width           =   930
               End
               Begin VB.TextBox TxtSingleToolDiameter 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   2430
                  TabIndex        =   79
                  Text            =   "1"
                  Top             =   1170
                  Width           =   700
               End
               Begin Threed.SSCommand BtnSingleStatus 
                  Height          =   510
                  Left            =   120
                  TabIndex        =   81
                  Top             =   2070
                  Width           =   4245
                  _Version        =   65536
                  _ExtentX        =   7479
                  _ExtentY        =   900
                  _StockProps     =   78
                  Caption         =   "Change Status for Single Pocket"
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
               Begin MSComCtl2.UpDown UpDnPocketIndex 
                  Height          =   690
                  Left            =   1755
                  TabIndex        =   83
                  Top             =   1200
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1217
                  _Version        =   393216
                  Value           =   1
                  Max             =   12
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown UpDown1 
                  Height          =   675
                  Left            =   3120
                  TabIndex        =   84
                  Top             =   1170
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1191
                  _Version        =   393216
                  Value           =   1
                  Max             =   7
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label12 
                  Caption         =   "Pick Status From List"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   225
                  TabIndex        =   87
                  Top             =   405
                  Width           =   2175
               End
               Begin VB.Label Label20 
                  Caption         =   "Pocket"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   177
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   240
                  Left            =   1080
                  TabIndex        =   86
                  Top             =   945
                  Width           =   765
               End
               Begin VB.Label Label13 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  Caption         =   "Drill Code"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   330
                  Left            =   2430
                  TabIndex        =   85
                  Top             =   945
                  Width           =   915
               End
            End
            Begin VB.CommandButton BtnDisplayPocketStatus 
               Caption         =   "Refresh Table"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1100
               Left            =   6165
               Picture         =   "frmMain.frx":AADFC
               Style           =   1  'Graphical
               TabIndex        =   76
               Top             =   3360
               Width           =   1200
            End
            Begin MSComCtl2.UpDown UpDnShelfList 
               Height          =   690
               Left            =   7020
               TabIndex        =   77
               Top             =   2475
               Width           =   255
               _ExtentX        =   423
               _ExtentY        =   1217
               _Version        =   393216
               Value           =   1
               Max             =   3
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
               Height          =   3795
               Left            =   120
               TabIndex        =   95
               Top             =   240
               Width           =   5805
               _ExtentX        =   10239
               _ExtentY        =   6694
               _Version        =   393216
               Rows            =   13
               Cols            =   5
               RowHeightMin    =   250
               AllowBigSelection=   0   'False
               FocusRect       =   0
               HighLight       =   0
               GridLines       =   3
               ScrollBars      =   0
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
            Begin VB.Label LabelShelfType 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   6120
               TabIndex        =   286
               Top             =   360
               Width           =   1335
            End
            Begin VB.Label Label9 
               Caption         =   "Shelf"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   6360
               TabIndex        =   96
               Top             =   2190
               Width           =   705
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "Cabinet"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   2655
            Left            =   -70680
            TabIndex        =   60
            Top             =   585
            Width           =   3735
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1. emergency stop 1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   200
               TabIndex        =   74
               Top             =   350
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2. emergency stop 2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   4
               Left            =   200
               TabIndex        =   73
               Top             =   600
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3. Door interlock 1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   5
               Left            =   200
               TabIndex        =   72
               Top             =   850
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4. Door interlock 2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   6
               Left            =   200
               TabIndex        =   71
               Top             =   1100
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5. PCdoor interlock 1"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   7
               Left            =   200
               TabIndex        =   70
               Top             =   1350
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "6. PCdoor interlock 2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   8
               Left            =   200
               TabIndex        =   69
               Top             =   1600
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "7."
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   200
               TabIndex        =   68
               Top             =   1850
               Visible         =   0   'False
               Width           =   2505
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   255
               Index           =   21
               Left            =   2800
               TabIndex        =   67
               Top             =   350
               Width           =   825
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   255
               Index           =   22
               Left            =   2800
               TabIndex        =   66
               Top             =   600
               Width           =   825
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   23
               Left            =   2800
               TabIndex        =   65
               Top             =   850
               Width           =   825
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   24
               Left            =   2800
               TabIndex        =   64
               Top             =   1100
               Width           =   825
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   25
               Left            =   2800
               TabIndex        =   63
               Top             =   1350
               Width           =   825
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   26
               Left            =   2800
               TabIndex        =   62
               Top             =   1600
               Width           =   825
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   201
               Left            =   2800
               TabIndex        =   61
               Top             =   1850
               Visible         =   0   'False
               Width           =   825
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Kiosk"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   2670
            Left            =   -74820
            TabIndex        =   45
            Top             =   585
            Width           =   3735
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "6. tool found"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   18
               Left            =   180
               TabIndex        =   59
               Top             =   1560
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "5. holder direction"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   19
               Left            =   180
               TabIndex        =   58
               Top             =   1320
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "4. valves close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   20
               Left            =   180
               TabIndex        =   57
               Top             =   1080
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3. user ack"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   21
               Left            =   180
               TabIndex        =   56
               Top             =   840
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2.door is close"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   22
               Left            =   180
               TabIndex        =   55
               Top             =   600
               Width           =   2505
            End
            Begin VB.Label LabelInputkiosk 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1.door is open"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   23
               Left            =   180
               TabIndex        =   54
               Top             =   360
               Width           =   2505
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   10
               Left            =   2830
               TabIndex        =   53
               Top             =   360
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   11
               Left            =   2830
               TabIndex        =   52
               Top             =   600
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   12
               Left            =   2830
               TabIndex        =   51
               Top             =   840
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   13
               Left            =   2830
               TabIndex        =   50
               Top             =   1080
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               ForeColor       =   &H00000000&
               Height          =   250
               Index           =   14
               Left            =   2830
               TabIndex        =   49
               Top             =   1320
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   15
               Left            =   2830
               TabIndex        =   48
               Top             =   1560
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   255
               Index           =   230
               Left            =   2835
               TabIndex        =   47
               Top             =   1800
               Visible         =   0   'False
               Width           =   735
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "7. "
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   39
               Left            =   180
               TabIndex        =   46
               Top             =   1800
               Visible         =   0   'False
               Width           =   2505
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Robot"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   2640
            Left            =   -66480
            TabIndex        =   38
            Top             =   600
            Width           =   3735
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "3. "
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   33
               Left            =   180
               TabIndex        =   44
               Top             =   840
               Visible         =   0   'False
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "2. Gripper 2 open "
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   34
               Left            =   180
               TabIndex        =   43
               Top             =   600
               Width           =   2505
            End
            Begin VB.Label lblIn1 
               BackStyle       =   0  'Transparent
               BorderStyle     =   1  'Fixed Single
               Caption         =   "1. Gripper open"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   35
               Left            =   180
               TabIndex        =   42
               Top             =   360
               Width           =   2505
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   255
               Index           =   16
               Left            =   2830
               TabIndex        =   41
               Top             =   360
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   20
               Left            =   2830
               TabIndex        =   40
               Top             =   600
               Width           =   735
            End
            Begin VB.Label LabelInput 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "---"
               Height          =   250
               Index           =   220
               Left            =   2830
               TabIndex        =   39
               Top             =   840
               Visible         =   0   'False
               Width           =   735
            End
         End
         Begin VB.CommandButton CmdSysOperation 
            Caption         =   "Start Automat"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Index           =   3
            Left            =   -69720
            Picture         =   "frmMain.frx":AB185
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   2280
            Width           =   2070
         End
         Begin TabDlg.SSTab SSTabTests 
            Height          =   5055
            Left            =   -74760
            TabIndex        =   5
            Top             =   840
            Width           =   12120
            _ExtentX        =   21378
            _ExtentY        =   8916
            _Version        =   393216
            TabHeight       =   520
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            TabCaption(0)   =   "Pockets"
            TabPicture(0)   =   "frmMain.frx":AC04F
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "SSFrame11"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).Control(1)=   "SSFrame8"
            Tab(0).Control(1).Enabled=   0   'False
            Tab(0).Control(2)=   "SSFrame10"
            Tab(0).Control(2).Enabled=   0   'False
            Tab(0).Control(3)=   "SSFrame7"
            Tab(0).Control(3).Enabled=   0   'False
            Tab(0).ControlCount=   4
            TabCaption(1)   =   "Load / Unload"
            TabPicture(1)   =   "frmMain.frx":AC06B
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "FrameLoadUnload"
            Tab(1).Control(1)=   "txtTestPocket(2)"
            Tab(1).Control(2)=   "BtnStopLoadUnload"
            Tab(1).Control(3)=   "BtnStartLoadUnload"
            Tab(1).Control(4)=   "UpDownTestPocket(2)"
            Tab(1).Control(5)=   "UpDown2"
            Tab(1).Control(6)=   "SSFrame14"
            Tab(1).Control(7)=   "Label1"
            Tab(1).ControlCount=   8
            TabCaption(2)   =   "Saw"
            TabPicture(2)   =   "frmMain.frx":AC087
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "FrameSaw"
            Tab(2).ControlCount=   1
            Begin Threed.SSFrame FrameLoadUnload 
               Height          =   3735
               Left            =   -74280
               TabIndex        =   368
               Top             =   720
               Width           =   6135
               _Version        =   65536
               _ExtentX        =   10821
               _ExtentY        =   6588
               _StockProps     =   14
               Caption         =   "Load  / Unload Tests"
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
               Font3D          =   1
               ShadowStyle     =   1
               Begin Threed.SSPanel SSPanel2 
                  Height          =   2505
                  Left            =   4200
                  TabIndex        =   373
                  Top             =   720
                  Width           =   1095
                  _Version        =   65536
                  _ExtentX        =   1931
                  _ExtentY        =   4419
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
                  BevelOuter      =   0
                  Font3D          =   1
                  Begin VB.OptionButton OptionLoop 
                     Caption         =   "Loop"
                     Height          =   375
                     Index           =   3
                     Left            =   0
                     TabIndex        =   377
                     Top             =   2040
                     Width           =   975
                  End
                  Begin VB.OptionButton OptionLoop 
                     Caption         =   "Loop"
                     Height          =   375
                     Index           =   2
                     Left            =   0
                     TabIndex        =   376
                     Top             =   1365
                     Width           =   975
                  End
                  Begin VB.OptionButton OptionLoop 
                     Caption         =   "Loop"
                     Height          =   375
                     Index           =   1
                     Left            =   0
                     TabIndex        =   375
                     Top             =   750
                     Width           =   975
                  End
                  Begin VB.OptionButton OptionLoop 
                     Caption         =   "Loop"
                     Height          =   375
                     Index           =   0
                     Left            =   0
                     TabIndex        =   374
                     Top             =   120
                     Width           =   975
                  End
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Chuck To Pocket"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   3
                  Left            =   720
                  Style           =   1  'Graphical
                  TabIndex        =   372
                  Top             =   2610
                  Width           =   3400
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Pocket To Chuck"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   2
                  Left            =   720
                  Style           =   1  'Graphical
                  TabIndex        =   371
                  Top             =   1980
                  Width           =   3400
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Pocket To Kiosk"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   1
                  Left            =   720
                  Style           =   1  'Graphical
                  TabIndex        =   370
                  Top             =   1350
                  Width           =   3400
               End
               Begin VB.OptionButton Option1 
                  Caption         =   "Kiosk To Pocket"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   0
                  Left            =   720
                  Style           =   1  'Graphical
                  TabIndex        =   369
                  Top             =   720
                  Width           =   3400
               End
            End
            Begin Threed.SSFrame FrameSaw 
               Height          =   2175
               Left            =   -74280
               TabIndex        =   360
               Top             =   1680
               Width           =   10455
               _Version        =   65536
               _ExtentX        =   18441
               _ExtentY        =   3836
               _StockProps     =   14
               Caption         =   "Saw Cycle"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   1
               Begin VB.OptionButton OptionSaw 
                  Caption         =   "Station to spindle"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   1
                  Left            =   1560
                  Style           =   1  'Graphical
                  TabIndex        =   364
                  Top             =   480
                  Width           =   3400
               End
               Begin VB.OptionButton OptionSaw 
                  Caption         =   "Spindle to station"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   600
                  Index           =   2
                  Left            =   1560
                  Style           =   1  'Graphical
                  TabIndex        =   363
                  Top             =   1200
                  Width           =   3400
               End
               Begin VB.TextBox TextStationNumber 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   5160
                  TabIndex        =   361
                  Text            =   "1"
                  Top             =   840
                  Width           =   915
               End
               Begin MSComCtl2.UpDown UpDownStationNumber 
                  Height          =   735
                  Left            =   6120
                  TabIndex        =   362
                  Top             =   840
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1296
                  _Version        =   393216
                  Value           =   1
                  Max             =   2
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin Threed.SSCommand CmdStartSawTest 
                  Height          =   630
                  Left            =   6600
                  TabIndex        =   365
                  Top             =   840
                  Width           =   2595
                  _Version        =   65536
                  _ExtentX        =   4586
                  _ExtentY        =   1111
                  _StockProps     =   78
                  Caption         =   "Start"
                  ForeColor       =   12582912
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Arial"
                     Size            =   14.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Font3D          =   2
               End
               Begin VB.Label Label19 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Station"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   5160
                  TabIndex        =   367
                  Top             =   525
                  Width           =   915
               End
            End
            Begin VB.TextBox txtTestPocket 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Index           =   2
               Left            =   -67365
               TabIndex        =   6
               Text            =   "101"
               Top             =   1575
               Width           =   915
            End
            Begin Threed.SSCommand BtnStopLoadUnload 
               Height          =   630
               Left            =   -65925
               TabIndex        =   7
               Top             =   1935
               Width           =   2595
               _Version        =   65536
               _ExtentX        =   4586
               _ExtentY        =   1111
               _StockProps     =   78
               Caption         =   "Stop Loop"
               ForeColor       =   12582912
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   2
            End
            Begin Threed.SSCommand BtnStartLoadUnload 
               Height          =   630
               Left            =   -65925
               TabIndex        =   8
               Top             =   1245
               Width           =   2595
               _Version        =   65536
               _ExtentX        =   4586
               _ExtentY        =   1111
               _StockProps     =   78
               Caption         =   "Start"
               ForeColor       =   12582912
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   2
            End
            Begin Threed.SSFrame SSFrame7 
               Height          =   2670
               Left            =   240
               TabIndex        =   9
               Top             =   500
               Width           =   5505
               _Version        =   65536
               _ExtentX        =   9710
               _ExtentY        =   4710
               _StockProps     =   14
               Caption         =   "Pocket To Pocket"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin VB.CheckBox CheckLoopP2P 
                  Caption         =   "Loop"
                  Height          =   615
                  Left            =   3960
                  TabIndex        =   378
                  Top             =   1920
                  Width           =   855
               End
               Begin VB.TextBox txtTestPocket 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Index           =   1
                  Left            =   2900
                  TabIndex        =   11
                  Text            =   "101"
                  Top             =   720
                  Width           =   920
               End
               Begin VB.TextBox txtTestPocket 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Index           =   0
                  Left            =   1080
                  TabIndex        =   10
                  Text            =   "101"
                  Top             =   720
                  Width           =   920
               End
               Begin Threed.SSCommand BtnTestPocketToPocket 
                  Height          =   630
                  Left            =   1185
                  TabIndex        =   12
                  Top             =   1920
                  Width           =   2655
                  _Version        =   65536
                  _ExtentX        =   4678
                  _ExtentY        =   1101
                  _StockProps     =   78
                  Caption         =   "Start Test"
                  ForeColor       =   12582912
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
               Begin MSComCtl2.UpDown UpDownTestPocket 
                  Height          =   690
                  Index           =   0
                  Left            =   2040
                  TabIndex        =   13
                  Top             =   720
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1217
                  _Version        =   393216
                  Value           =   1
                  Max             =   12
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown UpDownTestPocket 
                  Height          =   675
                  Index           =   1
                  Left            =   3825
                  TabIndex        =   14
                  Top             =   720
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1191
                  _Version        =   393216
                  Value           =   1
                  Max             =   12
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown UpDownTestShelf1 
                  Height          =   690
                  Left            =   840
                  TabIndex        =   15
                  Top             =   720
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1217
                  _Version        =   393216
                  Value           =   1
                  Max             =   3
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown UpDownTestShelf2 
                  Height          =   675
                  Left            =   2625
                  TabIndex        =   16
                  Top             =   720
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1191
                  _Version        =   393216
                  Value           =   1
                  Max             =   3
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label28 
                  Caption         =   "S"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   2625
                  TabIndex        =   20
                  Top             =   1395
                  Width           =   285
               End
               Begin VB.Label Label27 
                  Caption         =   "S"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   915
                  TabIndex        =   19
                  Top             =   1395
                  Width           =   195
               End
               Begin VB.Label Label16 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Pocket"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   18
                  Top             =   420
                  Width           =   920
               End
               Begin VB.Label Label22 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Pocket"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   2900
                  TabIndex        =   17
                  Top             =   420
                  Width           =   920
               End
            End
            Begin Threed.SSFrame SSFrame10 
               Height          =   1530
               Left            =   225
               TabIndex        =   21
               Top             =   3200
               Width           =   5505
               _Version        =   65536
               _ExtentX        =   9710
               _ExtentY        =   2699
               _StockProps     =   14
               Caption         =   "Drill Code"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin VB.TextBox TextToolsDiameter2 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   3315
                  TabIndex        =   22
                  Text            =   "1"
                  Top             =   480
                  Width           =   700
               End
               Begin MSComCtl2.UpDown UpDown5 
                  Height          =   675
                  Left            =   4035
                  TabIndex        =   23
                  Top             =   480
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1191
                  _Version        =   393216
                  Value           =   1
                  Max             =   7
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label23 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  Caption         =   "Select Drill Code :"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Left            =   360
                  TabIndex        =   24
                  Top             =   600
                  Width           =   2835
               End
            End
            Begin Threed.SSFrame SSFrame8 
               Height          =   2670
               Left            =   6050
               TabIndex        =   25
               Top             =   500
               Width           =   5500
               _Version        =   65536
               _ExtentX        =   9710
               _ExtentY        =   4710
               _StockProps     =   14
               Caption         =   "All Pockets"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin VB.CheckBox CheckLoopAllPockets 
                  Caption         =   "Loop"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   177
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   4320
                  TabIndex        =   326
                  Top             =   2040
                  Width           =   945
               End
               Begin VB.TextBox TextAllPocketsDiameter 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   3240
                  TabIndex        =   298
                  Text            =   "1"
                  Top             =   675
                  Width           =   700
               End
               Begin VB.TextBox TextShelfNumber2 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   1965
                  TabIndex        =   26
                  Text            =   "1"
                  Top             =   675
                  Width           =   700
               End
               Begin MSComCtl2.UpDown UpDownTestShelf 
                  Height          =   690
                  Left            =   1680
                  TabIndex        =   27
                  Top             =   660
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1217
                  _Version        =   393216
                  Value           =   1
                  Max             =   3
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin Threed.SSCommand BtnTestAllPockets 
                  Height          =   630
                  Left            =   1700
                  TabIndex        =   28
                  Top             =   1920
                  Width           =   2500
                  _Version        =   65536
                  _ExtentX        =   4410
                  _ExtentY        =   1111
                  _StockProps     =   78
                  Caption         =   "Start Test"
                  ForeColor       =   12582912
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
               Begin MSComCtl2.UpDown UpDownMyDiameter 
                  Height          =   690
                  Left            =   3960
                  TabIndex        =   300
                  Top             =   675
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1217
                  _Version        =   393216
                  Value           =   1
                  Max             =   8
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label30 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Diameter"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   3240
                  TabIndex        =   299
                  Top             =   360
                  Width           =   950
               End
               Begin VB.Label lblShelf 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  BackColor       =   &H80000005&
                  BorderStyle     =   1  'Fixed Single
                  Caption         =   "Shelf"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   11.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   285
                  Left            =   1680
                  TabIndex        =   29
                  Top             =   360
                  Width           =   1005
               End
            End
            Begin MSComCtl2.UpDown UpDownTestPocket 
               Height          =   675
               Index           =   2
               Left            =   -67635
               TabIndex        =   30
               Top             =   1575
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   1191
               _Version        =   393216
               Value           =   1
               Max             =   3
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDown2 
               Height          =   675
               Left            =   -66420
               TabIndex        =   31
               Top             =   1575
               Width           =   255
               _ExtentX        =   423
               _ExtentY        =   1191
               _Version        =   393216
               Value           =   1
               Max             =   12
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin Threed.SSFrame SSFrame14 
               Height          =   1530
               Left            =   -67560
               TabIndex        =   301
               Top             =   2920
               Width           =   4305
               _Version        =   65536
               _ExtentX        =   7594
               _ExtentY        =   2699
               _StockProps     =   14
               Caption         =   "Drill Code"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin VB.TextBox TextDrillCode 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   2955
                  TabIndex        =   302
                  Text            =   "1"
                  Top             =   480
                  Width           =   700
               End
               Begin MSComCtl2.UpDown UpDown3 
                  Height          =   675
                  Left            =   3675
                  TabIndex        =   303
                  Top             =   480
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1191
                  _Version        =   393216
                  Value           =   1
                  Max             =   8
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label31 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  Caption         =   "Select Drill Code :"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   375
                  Left            =   195
                  TabIndex        =   304
                  Top             =   615
                  Width           =   2580
               End
            End
            Begin Threed.SSFrame SSFrame11 
               Height          =   1530
               Left            =   6050
               TabIndex        =   401
               Top             =   3200
               Width           =   5500
               _Version        =   65536
               _ExtentX        =   9710
               _ExtentY        =   2699
               _StockProps     =   14
               Caption         =   "Reset Tests"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin MSComctlLib.ProgressBar ProgressBarResetTests 
                  Height          =   255
                  Left            =   1800
                  TabIndex        =   403
                  Top             =   360
                  Visible         =   0   'False
                  Width           =   2535
                  _ExtentX        =   4471
                  _ExtentY        =   450
                  _Version        =   393216
                  Appearance      =   1
                  Min             =   1
                  Max             =   7
               End
               Begin Threed.SSCommand CmdResetAllTests 
                  Height          =   630
                  Left            =   1800
                  TabIndex        =   402
                  Top             =   600
                  Width           =   2505
                  _Version        =   65536
                  _ExtentX        =   4410
                  _ExtentY        =   1111
                  _StockProps     =   78
                  Caption         =   "Reset All Tests"
                  ForeColor       =   12582912
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
            Begin VB.Label Label1 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pocket 2"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   -67365
               TabIndex        =   366
               Top             =   1275
               Width           =   915
            End
         End
         Begin Threed.SSFrame SSFrame6 
            Height          =   2055
            Left            =   -66240
            TabIndex        =   33
            Top             =   740
            Width           =   3615
            _Version        =   65536
            _ExtentX        =   6376
            _ExtentY        =   3625
            _StockProps     =   14
            Caption         =   "Tools"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   1
            Begin Threed.SSPanel TextAmountLeft 
               Height          =   492
               Left            =   1700
               TabIndex        =   34
               Top             =   1320
               Width           =   1300
               _Version        =   65536
               _ExtentX        =   2293
               _ExtentY        =   868
               _StockProps     =   15
               Caption         =   "---"
               BackColor       =   13160660
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
            Begin Threed.SSPanel TextAmount 
               Height          =   492
               Left            =   1700
               TabIndex        =   35
               Top             =   720
               Width           =   1300
               _Version        =   65536
               _ExtentX        =   2293
               _ExtentY        =   868
               _StockProps     =   15
               Caption         =   "---"
               BackColor       =   13160660
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               Caption         =   "Amount Left"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   204
               TabIndex        =   37
               Top             =   1404
               Width           =   1380
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               Caption         =   "Amount"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   168
               TabIndex        =   36
               Top             =   768
               Width           =   1056
            End
         End
         Begin Threed.SSFrame SSFrame2 
            Height          =   1215
            Left            =   1245
            TabIndex        =   140
            Top             =   3120
            Width           =   10005
            _Version        =   65536
            _ExtentX        =   17648
            _ExtentY        =   2143
            _StockProps     =   14
            Caption         =   "LOADING"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Font3D          =   4
            Begin Threed.SSPanel PanelK2P 
               Height          =   570
               Index           =   1
               Left            =   840
               TabIndex        =   141
               Top             =   360
               Visible         =   0   'False
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   1005
               _StockProps     =   15
               Caption         =   "Step"
               ForeColor       =   -2147483630
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
            Begin Threed.SSPanel PanelK2P 
               Height          =   570
               Index           =   3
               Left            =   1600
               TabIndex        =   279
               Top             =   360
               Width           =   7060
               _Version        =   65536
               _ExtentX        =   12462
               _ExtentY        =   1005
               _StockProps     =   15
               Caption         =   "step"
               ForeColor       =   -2147483630
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
         End
         Begin TabDlg.SSTab SSTab3 
            Height          =   5160
            Left            =   -74820
            TabIndex        =   142
            Top             =   765
            Width           =   12210
            _ExtentX        =   21537
            _ExtentY        =   9102
            _Version        =   393216
            Tab             =   1
            TabHeight       =   617
            TabCaption(0)   =   "Teach General Locations."
            TabPicture(0)   =   "frmMain.frx":AC0A3
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "MSFlexGrid3"
            Tab(0).Control(1)=   "FrameTeachlocation"
            Tab(0).Control(2)=   "BtnShowGeneralLocations"
            Tab(0).ControlCount=   3
            TabCaption(1)   =   "Teach Shelves Locations."
            TabPicture(1)   =   "frmMain.frx":AC0BF
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Label5"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).Control(1)=   "LabelShelfType2"
            Tab(1).Control(1).Enabled=   0   'False
            Tab(1).Control(2)=   "SSFrame12"
            Tab(1).Control(2).Enabled=   0   'False
            Tab(1).Control(3)=   "SSFrame4"
            Tab(1).Control(3).Enabled=   0   'False
            Tab(1).Control(4)=   "UpDnShelf"
            Tab(1).Control(4).Enabled=   0   'False
            Tab(1).Control(5)=   "TextShelfNum"
            Tab(1).Control(5).Enabled=   0   'False
            Tab(1).ControlCount=   6
            TabCaption(2)   =   "View Pockets Locations"
            TabPicture(2)   =   "frmMain.frx":AC0DB
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "SSFrame5"
            Tab(2).Control(1)=   "TablePocketsLocations"
            Tab(2).ControlCount=   2
            Begin VB.TextBox TextShelfNum 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   24
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   700
               Left            =   7515
               TabIndex        =   158
               Text            =   "1"
               Top             =   1725
               Width           =   700
            End
            Begin VB.CommandButton BtnShowGeneralLocations 
               Caption         =   "Show Locations"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1320
               Left            =   -68400
               Picture         =   "frmMain.frx":AC0F7
               Style           =   1  'Graphical
               TabIndex        =   157
               Top             =   840
               Width           =   1455
            End
            Begin MSComCtl2.UpDown UpDnShelf 
               Height          =   705
               Left            =   8235
               TabIndex        =   143
               Top             =   1725
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   1244
               _Version        =   393216
               Value           =   1
               Max             =   3
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin Threed.SSFrame SSFrame5 
               Height          =   4020
               Left            =   -66240
               TabIndex        =   144
               Top             =   600
               Width           =   3225
               _Version        =   65536
               _ExtentX        =   5689
               _ExtentY        =   7091
               _StockProps     =   14
               Caption         =   "Show Locations"
               ForeColor       =   128
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   1
               Begin VB.TextBox TxtViewPocket 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   1650
                  TabIndex        =   147
                  Text            =   "101"
                  Top             =   1440
                  Width           =   1020
               End
               Begin VB.CommandButton BtnViewPocketsLocations 
                  Caption         =   "Show Locations"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1215
                  Left            =   720
                  Picture         =   "frmMain.frx":AC480
                  Style           =   1  'Graphical
                  TabIndex        =   146
                  Top             =   2640
                  Width           =   1935
               End
               Begin VB.TextBox TextShelfNumber 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   345
                  TabIndex        =   145
                  Text            =   "1"
                  Top             =   1440
                  Width           =   700
               End
               Begin MSComCtl2.UpDown UpDownViewShelf 
                  Height          =   705
                  Left            =   1065
                  TabIndex        =   148
                  Top             =   1440
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1244
                  _Version        =   393216
                  Value           =   1
                  Max             =   3
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown UpDnViewPocket 
                  Height          =   705
                  Left            =   2760
                  TabIndex        =   149
                  Top             =   1440
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1244
                  _Version        =   393216
                  Value           =   1
                  Max             =   12
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin VB.Label Label18 
                  Caption         =   "Pocket"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   1920
                  TabIndex        =   152
                  Top             =   1140
                  Width           =   780
               End
               Begin VB.Label LabelShelvs 
                  Caption         =   "Display pockets Coordinates"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   570
                  Left            =   840
                  TabIndex        =   151
                  Top             =   480
                  Width           =   1815
               End
               Begin VB.Label Label3 
                  Caption         =   "Shelf"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   330
                  Left            =   480
                  TabIndex        =   150
                  Top             =   1125
                  Width           =   645
               End
            End
            Begin Threed.SSFrame FrameTeachlocation 
               Height          =   3930
               Left            =   -66810
               TabIndex        =   153
               Top             =   690
               Width           =   2535
               _Version        =   65536
               _ExtentX        =   4471
               _ExtentY        =   6932
               _StockProps     =   14
               Caption         =   " Teach Location"
               ForeColor       =   128
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
               Begin VB.ComboBox ComboLocations 
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   420
                  ItemData        =   "frmMain.frx":AC809
                  Left            =   135
                  List            =   "frmMain.frx":AC80B
                  TabIndex        =   154
                  Text            =   "Select Location"
                  Top             =   1935
                  Width           =   2250
               End
               Begin Threed.SSCommand BtnTeachPosition 
                  Height          =   945
                  Left            =   135
                  TabIndex        =   155
                  Top             =   2760
                  Width           =   2250
                  _Version        =   65536
                  _ExtentX        =   3969
                  _ExtentY        =   1676
                  _StockProps     =   78
                  Caption         =   "Teach Position"
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
                  Font3D          =   2
               End
               Begin VB.Label Label17 
                  Caption         =   "Read Current Location From Robot as General Position."
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1185
                  Left            =   120
                  TabIndex        =   156
                  Top             =   600
                  Width           =   2235
               End
            End
            Begin MSFlexGridLib.MSFlexGrid MSFlexGrid3 
               Height          =   2125
               Left            =   -74640
               TabIndex        =   159
               Top             =   840
               Width           =   5800
               _ExtentX        =   10239
               _ExtentY        =   3757
               _Version        =   393216
               Rows            =   6
               Cols            =   5
               RowHeightMin    =   250
               AllowBigSelection=   0   'False
               FocusRect       =   0
               HighLight       =   0
               GridLines       =   3
               ScrollBars      =   2
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
            Begin MSFlexGridLib.MSFlexGrid TablePocketsLocations 
               Height          =   4185
               Left            =   -74640
               TabIndex        =   160
               Top             =   720
               Width           =   8080
               _ExtentX        =   14261
               _ExtentY        =   7382
               _Version        =   393216
               Rows            =   11
               Cols            =   7
               RowHeightMin    =   350
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
            Begin Threed.SSFrame SSFrame4 
               Height          =   4350
               Left            =   8775
               TabIndex        =   161
               Top             =   495
               Width           =   2985
               _Version        =   65536
               _ExtentX        =   5265
               _ExtentY        =   7673
               _StockProps     =   14
               Caption         =   "Single Pocket"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   1
               Begin VB.TextBox TextToolsDiameter 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   1665
                  TabIndex        =   163
                  Text            =   "1"
                  Top             =   2565
                  Width           =   700
               End
               Begin VB.TextBox TextPocketNumber 
                  Alignment       =   2  'Center
                  BackColor       =   &H0080FF80&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   24
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   700
                  Left            =   270
                  TabIndex        =   162
                  Text            =   "101"
                  Top             =   2565
                  Width           =   930
               End
               Begin MSComCtl2.UpDown UpDnDiameter 
                  Height          =   675
                  Left            =   2355
                  TabIndex        =   164
                  Top             =   2565
                  Width           =   240
                  _ExtentX        =   423
                  _ExtentY        =   1191
                  _Version        =   393216
                  Value           =   1
                  Max             =   7
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin MSComCtl2.UpDown UpDownPocket 
                  Height          =   690
                  Left            =   1185
                  TabIndex        =   165
                  Top             =   2565
                  Width           =   255
                  _ExtentX        =   423
                  _ExtentY        =   1217
                  _Version        =   393216
                  Value           =   1
                  Max             =   12
                  Min             =   1
                  Enabled         =   -1  'True
               End
               Begin Threed.SSCommand BtnTeachSingle 
                  Height          =   600
                  Left            =   270
                  TabIndex        =   166
                  Top             =   3555
                  Width           =   2490
                  _Version        =   65536
                  _ExtentX        =   4392
                  _ExtentY        =   1058
                  _StockProps     =   78
                  Caption         =   "Teach Single Pocket"
                  ForeColor       =   12582912
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Font3D          =   2
               End
               Begin VB.Label Label6 
                  Caption         =   "Read Current Location From Robot as single pocket."
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1185
                  Left            =   360
                  TabIndex        =   281
                  Top             =   585
                  Width           =   2235
               End
               Begin VB.Label Label11 
                  Alignment       =   2  'Center
                  Appearance      =   0  'Flat
                  Caption         =   "Drill Code"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   330
                  Left            =   1665
                  TabIndex        =   168
                  Top             =   2340
                  Width           =   915
               End
               Begin VB.Label Label10 
                  Caption         =   "Pocket"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   405
                  TabIndex        =   167
                  Top             =   2340
                  Width           =   915
               End
            End
            Begin Threed.SSFrame SSFrame12 
               Height          =   4485
               Left            =   135
               TabIndex        =   169
               Top             =   405
               Width           =   6585
               _Version        =   65536
               _ExtentX        =   11615
               _ExtentY        =   7911
               _StockProps     =   14
               Caption         =   "Multi Pockets"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   1
               Begin VB.CommandButton BtnShowFirstLastPockets 
                  Caption         =   "Refresh Table"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   960
                  Left            =   3510
                  Picture         =   "frmMain.frx":AC80D
                  Style           =   1  'Graphical
                  TabIndex        =   170
                  Top             =   2340
                  Width           =   2775
               End
               Begin MSFlexGridLib.MSFlexGrid TableTeach 
                  Height          =   945
                  Left            =   2250
                  TabIndex        =   171
                  Top             =   3375
                  Width           =   4065
                  _ExtentX        =   7170
                  _ExtentY        =   1667
                  _Version        =   393216
                  Rows            =   3
                  Cols            =   4
                  ScrollBars      =   0
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
               Begin Threed.SSCommand BtnTeachFirst 
                  Height          =   735
                  Left            =   120
                  TabIndex        =   172
                  Top             =   3000
                  Width           =   1545
                  _Version        =   65536
                  _ExtentX        =   2725
                  _ExtentY        =   1296
                  _StockProps     =   78
                  Caption         =   "Teach First"
                  ForeColor       =   12582912
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Font3D          =   2
               End
               Begin Threed.SSCommand BtnTeachLast 
                  Height          =   735
                  Left            =   3420
                  TabIndex        =   173
                  Top             =   480
                  Width           =   1365
                  _Version        =   65536
                  _ExtentX        =   2408
                  _ExtentY        =   1296
                  _StockProps     =   78
                  Caption         =   "Teach Last"
                  ForeColor       =   12582912
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Font3D          =   2
               End
               Begin Threed.SSCommand BtnCalcPoints 
                  Height          =   740
                  Left            =   4920
                  TabIndex        =   174
                  Top             =   480
                  Width           =   1500
                  _Version        =   65536
                  _ExtentX        =   2646
                  _ExtentY        =   1305
                  _StockProps     =   78
                  Caption         =   "Calculate"
                  ForeColor       =   12582912
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Font3D          =   2
               End
               Begin VB.Image ImgToolType 
                  Height          =   2445
                  Left            =   180
                  Picture         =   "frmMain.frx":ACB96
                  Stretch         =   -1  'True
                  Top             =   450
                  Width           =   3165
               End
            End
            Begin VB.Label LabelShelfType2 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Label2"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   14.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   375
               Left            =   6840
               TabIndex        =   288
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "Shelf"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   7605
               TabIndex        =   175
               Top             =   1440
               Width           =   570
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   1905
            Index           =   4
            Left            =   -70800
            TabIndex        =   176
            Top             =   3960
            Width           =   8025
            _Version        =   65536
            _ExtentX        =   14155
            _ExtentY        =   3360
            _StockProps     =   14
            Caption         =   "Robot Status"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Font3D          =   4
            ShadowStyle     =   1
            Begin VB.ListBox ListAutomatMessage 
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   915
               Left            =   360
               TabIndex        =   404
               Top             =   480
               Width           =   7335
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   1188
            Index           =   5
            Left            =   -70824
            TabIndex        =   177
            Top             =   760
            Width           =   4416
            _Version        =   65536
            _ExtentX        =   7789
            _ExtentY        =   2096
            _StockProps     =   14
            Caption         =   "Speed"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   4
            ShadowStyle     =   1
            Begin MSComctlLib.Slider SliderAutoSpeed 
               Height          =   465
               Left            =   240
               TabIndex        =   178
               Top             =   480
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   820
               _Version        =   393216
               Min             =   1
               Max             =   100
               SelStart        =   1
               TickStyle       =   1
               Value           =   1
               TextPosition    =   1
            End
            Begin VB.Label lblProccentTag 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               Caption         =   "%"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   300
               Index           =   1
               Left            =   3645
               TabIndex        =   180
               Top             =   585
               Width           =   210
            End
            Begin VB.Label LableAutoSpeed 
               Caption         =   "100"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000080&
               Height          =   285
               Left            =   3240
               TabIndex        =   179
               Top             =   585
               Width           =   420
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   5145
            Index           =   8
            Left            =   -74595
            TabIndex        =   181
            Top             =   760
            Width           =   3555
            _Version        =   65536
            _ExtentX        =   6271
            _ExtentY        =   9075
            _StockProps     =   14
            Caption         =   "Auto Mode"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   4
            ShadowStyle     =   1
            Begin Threed.SSCommand CmdResetWPiece 
               Height          =   855
               Left            =   240
               TabIndex        =   407
               Top             =   3960
               Width           =   3015
               _Version        =   65536
               _ExtentX        =   5318
               _ExtentY        =   1508
               _StockProps     =   78
               Caption         =   "Reset Work Piece"
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
               Font3D          =   2
            End
            Begin Threed.SSFrame SSFrame1 
               Height          =   1185
               Index           =   2
               Left            =   225
               TabIndex        =   182
               Top             =   2640
               Width           =   3045
               _Version        =   65536
               _ExtentX        =   5371
               _ExtentY        =   2090
               _StockProps     =   14
               Caption         =   "Night Mode"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin VB.OptionButton optNightMode 
                  Caption         =   "On"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   550
                  Index           =   0
                  Left            =   250
                  Style           =   1  'Graphical
                  TabIndex        =   184
                  Top             =   450
                  Width           =   1100
               End
               Begin VB.OptionButton optNightMode 
                  Caption         =   "Off"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   550
                  Index           =   1
                  Left            =   1692
                  Style           =   1  'Graphical
                  TabIndex        =   183
                  Top             =   450
                  Value           =   -1  'True
                  Width           =   1100
               End
            End
            Begin Threed.SSFrame SSFrame1 
               Height          =   1095
               Index           =   3
               Left            =   225
               TabIndex        =   185
               Top             =   1380
               Width           =   3045
               _Version        =   65536
               _ExtentX        =   5380
               _ExtentY        =   1940
               _StockProps     =   14
               Caption         =   "One Tool"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Font3D          =   4
               ShadowStyle     =   1
               Begin VB.OptionButton optOneTool 
                  Caption         =   "Off"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   550
                  Index           =   1
                  Left            =   1710
                  Style           =   1  'Graphical
                  TabIndex        =   187
                  Top             =   432
                  Value           =   -1  'True
                  Width           =   1100
               End
               Begin VB.OptionButton optOneTool 
                  Caption         =   "On"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   550
                  Index           =   0
                  Left            =   255
                  Style           =   1  'Graphical
                  TabIndex        =   186
                  Top             =   435
                  Width           =   1100
               End
            End
            Begin VB.Label LabelMode 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Work Piece"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   450
               Left            =   270
               TabIndex        =   188
               Top             =   630
               Width           =   3045
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   2415
            Index           =   9
            Left            =   1245
            TabIndex        =   189
            Top             =   675
            Width           =   9975
            _Version        =   65536
            _ExtentX        =   17595
            _ExtentY        =   4260
            _StockProps     =   14
            Caption         =   "Load/Unload Tool"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   4
            ShadowStyle     =   1
            Begin VB.TextBox txtWorkPiece 
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
               Height          =   620
               Index           =   1
               Left            =   7450
               TabIndex        =   192
               Text            =   "1"
               ToolTipText     =   "(1-100)"
               Top             =   590
               Width           =   900
            End
            Begin VB.TextBox txtLoadUnloadPocket 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   7450
               TabIndex        =   191
               Text            =   "101"
               Top             =   1485
               Width           =   900
            End
            Begin VB.TextBox txtLoadUnloadShelf 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   20.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   620
               Left            =   8700
               TabIndex        =   190
               Text            =   "1"
               Top             =   1485
               Width           =   900
            End
            Begin Threed.SSCommand cmdUnloadTool 
               Height          =   615
               Left            =   405
               TabIndex        =   194
               Top             =   1485
               Width           =   3495
               _Version        =   65536
               _ExtentX        =   6165
               _ExtentY        =   1094
               _StockProps     =   78
               Caption         =   "Unload Tool From Pocket"
               ForeColor       =   12582912
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
            Begin Threed.SSPanel SSPanel 
               Height          =   615
               Index           =   0
               Left            =   5190
               TabIndex        =   195
               Top             =   1485
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   1080
               _StockProps     =   15
               Caption         =   "Pocket"
               ForeColor       =   -2147483630
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   14.19
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
            Begin Threed.SSPanel SSPanel 
               Height          =   620
               Index           =   1
               Left            =   5190
               TabIndex        =   196
               Top             =   590
               Width           =   2100
               _Version        =   65536
               _ExtentX        =   3704
               _ExtentY        =   1080
               _StockProps     =   15
               Caption         =   "Work Piece"
               ForeColor       =   -2147483630
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   14.58
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
            Begin MSComCtl2.UpDown UpDownLoadUnloadPocket 
               Height          =   615
               Left            =   8355
               TabIndex        =   197
               Top             =   1485
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   1085
               _Version        =   393216
               Value           =   1
               Max             =   12
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin MSComCtl2.UpDown UpDownLoadUnloadShelf 
               Height          =   615
               Left            =   9600
               TabIndex        =   198
               Top             =   1485
               Width           =   240
               _ExtentX        =   423
               _ExtentY        =   1085
               _Version        =   393216
               Value           =   1
               Max             =   3
               Min             =   1
               Enabled         =   -1  'True
            End
            Begin Threed.SSCommand cmdLoadTool 
               Height          =   615
               Left            =   435
               TabIndex        =   193
               Top             =   585
               Width           =   3450
               _Version        =   65536
               _ExtentX        =   6085
               _ExtentY        =   1085
               _StockProps     =   78
               Caption         =   "Load Tool To Pocket"
               ForeColor       =   12582912
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
            Begin VB.Label LabelLoad 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Run"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   3960
               TabIndex        =   382
               Top             =   690
               Width           =   765
            End
            Begin VB.Label labelUnload 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Stop"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   405
               Left            =   3960
               TabIndex        =   383
               Top             =   1590
               Width           =   765
            End
            Begin VB.Label Label25 
               Caption         =   "Pocket"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   7560
               TabIndex        =   200
               Top             =   1200
               Width           =   735
            End
            Begin VB.Label Label26 
               Alignment       =   2  'Center
               Caption         =   "Shelf"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Left            =   8760
               TabIndex        =   199
               Top             =   1200
               Width           =   765
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   3900
            Index           =   1200
            Left            =   -74760
            TabIndex        =   201
            Top             =   765
            Width           =   5355
            _Version        =   65536
            _ExtentX        =   9446
            _ExtentY        =   6879
            _StockProps     =   14
            Caption         =   "SemiAutomat "
            ForeColor       =   128
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowStyle     =   1
            Begin Threed.SSCommand cmdGoParkingPos 
               Height          =   660
               Left            =   636
               TabIndex        =   202
               Top             =   972
               Width           =   4236
               _Version        =   65536
               _ExtentX        =   7451
               _ExtentY        =   1164
               _StockProps     =   78
               Caption         =   " Parking Position"
               ForeColor       =   12582912
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
            Begin Threed.SSCommand cmdGoExchangePos 
               Height          =   660
               Left            =   636
               TabIndex        =   203
               Top             =   1848
               Width           =   4260
               _Version        =   65536
               _ExtentX        =   7514
               _ExtentY        =   1164
               _StockProps     =   78
               Caption         =   " Exchange Gripper Position"
               ForeColor       =   12582912
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
            Begin Threed.SSCommand cmdGoZeroPos 
               Height          =   660
               Left            =   600
               TabIndex        =   204
               Top             =   2715
               Width           =   4260
               _Version        =   65536
               _ExtentX        =   7514
               _ExtentY        =   1164
               _StockProps     =   78
               Caption         =   "Retract Position."
               ForeColor       =   12582912
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
         Begin Threed.SSFrame SSFrame3 
            Height          =   1335
            Left            =   1245
            TabIndex        =   205
            Top             =   4440
            Width           =   10005
            _Version        =   65536
            _ExtentX        =   17648
            _ExtentY        =   2355
            _StockProps     =   14
            Caption         =   "UNLOADING"
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Font3D          =   4
            Begin Threed.SSPanel PanelP2k 
               Height          =   570
               Index           =   1
               Left            =   795
               TabIndex        =   206
               Top             =   480
               Visible         =   0   'False
               Width           =   705
               _Version        =   65536
               _ExtentX        =   1244
               _ExtentY        =   1005
               _StockProps     =   15
               Caption         =   "Step"
               ForeColor       =   -2147483630
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
            Begin Threed.SSPanel PanelP2k 
               Height          =   570
               Index           =   3
               Left            =   1600
               TabIndex        =   280
               Top             =   480
               Width           =   7060
               _Version        =   65536
               _ExtentX        =   12462
               _ExtentY        =   1005
               _StockProps     =   15
               Caption         =   "Step"
               ForeColor       =   -2147483630
               BackColor       =   12640511
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   1
               BevelInner      =   1
               Font3D          =   2
            End
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   4965
            Index           =   11
            Left            =   -74730
            TabIndex        =   207
            Top             =   720
            Width           =   12015
            _Version        =   65536
            _ExtentX        =   21193
            _ExtentY        =   8758
            _StockProps     =   14
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   2
            Font3D          =   4
            ShadowStyle     =   1
            Begin VB.Frame fraJog 
               Caption         =   "Jog Robot"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   3600
               Index           =   0
               Left            =   360
               TabIndex        =   208
               Top             =   240
               Width           =   6840
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Rx+"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   6
                  Left            =   3432
                  MouseIcon       =   "frmMain.frx":C4D96
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C4EE8
                  Style           =   1  'Graphical
                  TabIndex        =   399
                  ToolTipText     =   "(+)"
                  Top             =   555
                  Width           =   1005
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "X+"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   0
                  Left            =   120
                  MouseIcon       =   "frmMain.frx":C564A
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C579C
                  Style           =   1  'Graphical
                  TabIndex        =   398
                  ToolTipText     =   "(+)"
                  Top             =   555
                  Width           =   1000
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Z+"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   4
                  Left            =   2328
                  MouseIcon       =   "frmMain.frx":C5BDE
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C5D30
                  Style           =   1  'Graphical
                  TabIndex        =   397
                  ToolTipText     =   "(+)"
                  Top             =   555
                  Width           =   1000
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Y+"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   2
                  Left            =   1224
                  MouseIcon       =   "frmMain.frx":C6172
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C62C4
                  Style           =   1  'Graphical
                  TabIndex        =   396
                  ToolTipText     =   "(+)"
                  Top             =   555
                  Width           =   1000
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Rz-"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   11
                  Left            =   5650
                  MouseIcon       =   "frmMain.frx":C6706
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C6858
                  Style           =   1  'Graphical
                  TabIndex        =   395
                  ToolTipText     =   "(+)"
                  Top             =   1755
                  Width           =   1005
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Ry-"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   9
                  Left            =   4541
                  MouseIcon       =   "frmMain.frx":C6FBA
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C710C
                  Style           =   1  'Graphical
                  TabIndex        =   394
                  ToolTipText     =   "(+)"
                  Top             =   1755
                  Width           =   1005
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Rx-"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   7
                  Left            =   3432
                  MouseIcon       =   "frmMain.frx":C786E
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C79C0
                  Style           =   1  'Graphical
                  TabIndex        =   393
                  ToolTipText     =   "(+)"
                  Top             =   1755
                  Width           =   1005
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Z-"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   5
                  Left            =   2328
                  MouseIcon       =   "frmMain.frx":C8122
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C8274
                  Style           =   1  'Graphical
                  TabIndex        =   392
                  ToolTipText     =   "(-)"
                  Top             =   1755
                  Width           =   1000
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Y-"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   3
                  Left            =   1224
                  MouseIcon       =   "frmMain.frx":C86B6
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C8808
                  Style           =   1  'Graphical
                  TabIndex        =   391
                  ToolTipText     =   "(-)"
                  Top             =   1755
                  Width           =   1000
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "X-"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   1
                  Left            =   120
                  MouseIcon       =   "frmMain.frx":C8C4A
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C8D9C
                  Style           =   1  'Graphical
                  TabIndex        =   390
                  ToolTipText     =   "(-)"
                  Top             =   1755
                  Width           =   1000
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Rz+"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   10
                  Left            =   5650
                  MouseIcon       =   "frmMain.frx":C91DE
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C9330
                  Style           =   1  'Graphical
                  TabIndex        =   389
                  ToolTipText     =   "(+)"
                  Top             =   555
                  Width           =   1005
               End
               Begin VB.CommandButton BtnJog 
                  Caption         =   "Ry+"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1000
                  Index           =   8
                  Left            =   4541
                  MouseIcon       =   "frmMain.frx":C9A92
                  MousePointer    =   99  'Custom
                  Picture         =   "frmMain.frx":C9BE4
                  Style           =   1  'Graphical
                  TabIndex        =   388
                  ToolTipText     =   "(+)"
                  Top             =   555
                  Width           =   1005
               End
            End
            Begin Threed.SSFrame SSFrame1 
               Height          =   1545
               Index           =   0
               Left            =   8955
               TabIndex        =   209
               Top             =   240
               Width           =   2835
               _Version        =   65536
               _ExtentX        =   5001
               _ExtentY        =   2725
               _StockProps     =   14
               Caption         =   "Speed"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               ShadowStyle     =   1
               Begin MSComctlLib.Slider SliderManuSpeed 
                  Height          =   585
                  Left            =   240
                  TabIndex        =   210
                  Top             =   720
                  Width           =   2310
                  _ExtentX        =   4075
                  _ExtentY        =   1032
                  _Version        =   393216
                  Min             =   1
                  Max             =   100
                  SelStart        =   1
                  TickStyle       =   1
                  Value           =   1
               End
               Begin VB.Label LabelPercent 
                  Caption         =   "100"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   375
                  Left            =   765
                  TabIndex        =   212
                  Top             =   450
                  Width           =   420
               End
               Begin VB.Label lblProccentTag 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  Caption         =   "%"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00000080&
                  Height          =   300
                  Index           =   2
                  Left            =   1620
                  TabIndex        =   211
                  Top             =   450
                  Width           =   225
               End
            End
            Begin Threed.SSFrame fraStep 
               Height          =   3600
               Index           =   6
               Left            =   7320
               TabIndex        =   213
               Top             =   240
               Width           =   1485
               _Version        =   65536
               _ExtentX        =   2619
               _ExtentY        =   6350
               _StockProps     =   14
               Caption         =   "Step(mm)"
               ForeColor       =   64
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Alignment       =   2
               ShadowStyle     =   1
               Begin VB.OptionButton IncRobotStep 
                  Caption         =   "0.1"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   430
                  Index           =   0
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   218
                  Top             =   520
                  Value           =   -1  'True
                  Width           =   1000
               End
               Begin VB.OptionButton IncRobotStep 
                  Caption         =   "1.0"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   430
                  Index           =   1
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   217
                  Top             =   1080
                  Width           =   1000
               End
               Begin VB.OptionButton IncRobotStep 
                  Caption         =   "2.0"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   430
                  Index           =   2
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   216
                  Top             =   1640
                  Width           =   1000
               End
               Begin VB.OptionButton IncRobotStep 
                  Caption         =   "5.0"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   430
                  Index           =   3
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   215
                  Top             =   2200
                  Width           =   1000
               End
               Begin VB.OptionButton IncRobotStep 
                  Caption         =   "10.0"
                  BeginProperty Font 
                     Name            =   "Microsoft Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   430
                  Index           =   4
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   214
                  Top             =   2760
                  Width           =   1000
               End
            End
         End
      End
      Begin Threed.SSPanel pnlGembal 
         Height          =   2415
         Index           =   0
         Left            =   330
         TabIndex        =   219
         Top             =   750
         Width           =   6000
         _Version        =   65536
         _ExtentX        =   10583
         _ExtentY        =   4260
         _StockProps     =   15
         Caption         =   "Current Position"
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   2
         Font3D          =   4
         Alignment       =   6
         Begin VB.Timer tmrUpdateRobotStatus 
            Interval        =   3000
            Left            =   5400
            Top             =   120
         End
         Begin VB.Timer tmrRobotQuery 
            Interval        =   500
            Left            =   120
            Top             =   120
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   300
            Index           =   1
            Left            =   2070
            TabIndex        =   220
            Top             =   410
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Y[mm]"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlPositionData 
            Height          =   525
            Index           =   1
            Left            =   2150
            TabIndex        =   221
            Top             =   700
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "123.456"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
         End
         Begin Threed.SSPanel pnlPositionData 
            Height          =   525
            Index           =   0
            Left            =   225
            TabIndex        =   222
            Top             =   700
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "123.456"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   300
            Index           =   2
            Left            =   4080
            TabIndex        =   223
            Top             =   410
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "Z[mm]"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   300
            Index           =   3
            Left            =   240
            TabIndex        =   224
            Top             =   410
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   529
            _StockProps     =   15
            Caption         =   "X[mm]"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlPositionData 
            Height          =   525
            Index           =   2
            Left            =   4100
            TabIndex        =   225
            ToolTipText     =   "Deg"
            Top             =   700
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "123.456"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   285
            Index           =   9
            Left            =   180
            TabIndex        =   226
            Top             =   1440
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rx[Deg]"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   285
            Index           =   12
            Left            =   1980
            TabIndex        =   227
            Top             =   1440
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Ry[Deg]"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Begin Threed.SSPanel pnlGembalDescp 
               Height          =   285
               Index           =   13
               Left            =   2160
               TabIndex        =   228
               Top             =   1485
               Width           =   1710
               _Version        =   65536
               _ExtentX        =   3016
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "W (deg.)"
               ForeColor       =   16711680
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
            End
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   285
            Index           =   14
            Left            =   4050
            TabIndex        =   229
            Top             =   1440
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "Rz[Deg]"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.76
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Begin Threed.SSPanel pnlGembalDescp 
               Height          =   285
               Index           =   15
               Left            =   2160
               TabIndex        =   230
               Top             =   1485
               Width           =   1710
               _Version        =   65536
               _ExtentX        =   3016
               _ExtentY        =   503
               _StockProps     =   15
               Caption         =   "W (deg.)"
               ForeColor       =   16711680
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
            End
         End
         Begin Threed.SSPanel pnlPositionData 
            Height          =   525
            Index           =   3
            Left            =   225
            TabIndex        =   231
            Top             =   1800
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "123.456"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
         End
         Begin Threed.SSPanel pnlPositionData 
            Height          =   525
            Index           =   4
            Left            =   2150
            TabIndex        =   232
            Top             =   1800
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "123.456"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
         End
         Begin Threed.SSPanel pnlPositionData 
            Height          =   525
            Index           =   5
            Left            =   4100
            TabIndex        =   233
            Top             =   1800
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   926
            _StockProps     =   15
            Caption         =   "123.456"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   1
         End
      End
      Begin Threed.SSPanel pnlMod 
         Height          =   2415
         Left            =   9000
         TabIndex        =   234
         Top             =   750
         Width           =   6000
         _Version        =   65536
         _ExtentX        =   10583
         _ExtentY        =   4260
         _StockProps     =   15
         Caption         =   "Work Piece"
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BevelInner      =   2
         Font3D          =   2
         Alignment       =   6
         Begin Threed.SSPanel pnlMode 
            Height          =   495
            Index           =   2
            Left            =   200
            TabIndex        =   235
            Top             =   700
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "---"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnlServoLoop 
            Height          =   495
            Index           =   2
            Left            =   2000
            TabIndex        =   236
            Top             =   1800
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "gripper"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlMode 
            Height          =   495
            Index           =   1
            Left            =   2000
            TabIndex        =   237
            Top             =   700
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "---"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnlServoLoop 
            Height          =   495
            Index           =   1
            Left            =   3800
            TabIndex        =   238
            Top             =   1800
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Key"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
         End
         Begin Threed.SSPanel pnlMode 
            Height          =   495
            Index           =   0
            Left            =   3800
            TabIndex        =   239
            Top             =   700
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "---"
            ForeColor       =   -2147483630
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel pnlServoStatus 
            Height          =   495
            Index           =   0
            Left            =   200
            TabIndex        =   240
            Top             =   1800
            Width           =   1700
            _Version        =   65536
            _ExtentX        =   2999
            _ExtentY        =   873
            _StockProps     =   15
            Caption         =   "Servo"
            ForeColor       =   0
            BackColor       =   12632256
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            FloodColor      =   49152
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   255
            Index           =   0
            Left            =   165
            TabIndex        =   241
            Top             =   420
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Work Piece"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   285
            Index           =   5
            Left            =   3930
            TabIndex        =   242
            Top             =   420
            Width           =   1605
            _Version        =   65536
            _ExtentX        =   2831
            _ExtentY        =   503
            _StockProps     =   15
            Caption         =   "NC program"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   225
            Index           =   6
            Left            =   2025
            TabIndex        =   243
            Top             =   1545
            Width           =   1710
            _Version        =   65536
            _ExtentX        =   3016
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Gripper"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   225
            Index           =   7
            Left            =   3810
            TabIndex        =   244
            Top             =   1545
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Key State"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   225
            Index           =   8
            Left            =   165
            TabIndex        =   245
            Top             =   1545
            Width           =   1635
            _Version        =   65536
            _ExtentX        =   2884
            _ExtentY        =   397
            _StockProps     =   15
            Caption         =   "Servo"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel pnlGembalDescp 
            Height          =   255
            Index           =   4
            Left            =   1920
            TabIndex        =   285
            Top             =   420
            Width           =   1665
            _Version        =   65536
            _ExtentX        =   2937
            _ExtentY        =   450
            _StockProps     =   15
            Caption         =   "Amount Left"
            ForeColor       =   16711680
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
         End
         Begin Threed.SSPanel SSPanel1 
            Height          =   375
            Left            =   2160
            TabIndex        =   297
            Top             =   1200
            Width           =   1695
            _Version        =   65536
            _ExtentX        =   2990
            _ExtentY        =   661
            _StockProps     =   15
            Caption         =   "Robot"
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   0
            Font3D          =   4
         End
      End
      Begin TabDlg.SSTab SSTab4 
         Height          =   6135
         Left            =   13095
         TabIndex        =   246
         Top             =   3240
         Width           =   1830
         _ExtentX        =   3228
         _ExtentY        =   10821
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   882
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Operation"
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FrameRunMode"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "CmdSysOperation(2)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "ProgressBarReset"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "CmdSysOperation(4)"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).ControlCount=   4
         Begin VB.CommandButton CmdSysOperation 
            Caption         =   "Reset Profibus"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   810
            Index           =   4
            Left            =   360
            Style           =   1  'Graphical
            TabIndex        =   359
            Top             =   5175
            Visible         =   0   'False
            Width           =   1164
         End
         Begin MSComctlLib.ProgressBar ProgressBarReset 
            Height          =   255
            Left            =   330
            TabIndex        =   284
            Top             =   4680
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Min             =   1
            Max             =   39
         End
         Begin VB.CommandButton CmdSysOperation 
            Caption         =   "Reset"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   990
            Index           =   2
            Left            =   360
            Picture         =   "frmMain.frx":CA346
            Style           =   1  'Graphical
            TabIndex        =   247
            Top             =   3600
            Width           =   1164
         End
         Begin Threed.SSPanel SSPanel7 
            Height          =   2685
            Left            =   -74850
            TabIndex        =   248
            Top             =   510
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   4736
            _StockProps     =   15
            Caption         =   "Table Command"
            ForeColor       =   255
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
            Alignment       =   6
            Begin VB.OptionButton optTableCommand 
               Caption         =   "Operate Mode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   5
               Left            =   180
               TabIndex        =   254
               Top             =   1860
               Width           =   2325
            End
            Begin VB.OptionButton optTableCommand 
               Caption         =   "Set Load"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   7
               Left            =   180
               TabIndex        =   253
               Top             =   825
               Width           =   2385
            End
            Begin VB.OptionButton optTableCommand 
               Caption         =   "Mode"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   8
               Left            =   180
               TabIndex        =   252
               Top             =   480
               Width           =   2115
            End
            Begin VB.OptionButton optTableCommand 
               Caption         =   "Servo Loop"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   9
               Left            =   180
               TabIndex        =   251
               Top             =   1515
               Width           =   2145
            End
            Begin VB.OptionButton optTableCommand 
               Caption         =   "Table Limit"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   10
               Left            =   180
               TabIndex        =   250
               Top             =   1170
               Width           =   2175
            End
            Begin VB.OptionButton optTableCommand 
               Caption         =   "Data Command"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   330
               Index           =   11
               Left            =   180
               TabIndex        =   249
               Top             =   2190
               Value           =   -1  'True
               Width           =   2385
            End
         End
         Begin Threed.SSPanel SSPanel8 
            Height          =   1335
            Left            =   -74880
            TabIndex        =   255
            Top             =   3330
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   2355
            _StockProps     =   15
            Caption         =   "Service"
            ForeColor       =   255
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
            Alignment       =   6
            Begin VB.PictureBox HiTime1 
               BackColor       =   &H000000FF&
               Height          =   1000
               Left            =   0
               ScaleHeight     =   945
               ScaleWidth      =   945
               TabIndex        =   258
               Top             =   0
               Width           =   1000
            End
            Begin VB.CheckBox Check1 
               Caption         =   "Host"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   345
               Left            =   60
               Style           =   1  'Graphical
               TabIndex        =   257
               Top             =   60
               Visible         =   0   'False
               Width           =   555
            End
            Begin VB.CheckBox Check2 
               Caption         =   "OFF"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   177
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   1830
               Style           =   1  'Graphical
               TabIndex        =   256
               Top             =   480
               Width           =   645
            End
            Begin Threed.SSPanel pnlScriptFileType 
               Height          =   312
               Index           =   0
               Left            =   780
               TabIndex        =   259
               Top             =   480
               Width           =   924
               _Version        =   65536
               _ExtentX        =   1630
               _ExtentY        =   550
               _StockProps     =   15
               Caption         =   "Record"
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               Font3D          =   2
               Autosize        =   1
            End
            Begin Threed.SSCommand SSCommand1 
               Height          =   285
               Left            =   1800
               TabIndex        =   260
               Top             =   900
               Width           =   645
               _Version        =   65536
               _ExtentX        =   1138
               _ExtentY        =   503
               _StockProps     =   78
            End
            Begin Threed.SSPanel pnlScriptFileType 
               Height          =   435
               Index           =   1
               Left            =   750
               TabIndex        =   261
               Top             =   810
               Width           =   975
               _Version        =   65536
               _ExtentX        =   1720
               _ExtentY        =   767
               _StockProps     =   15
               Caption         =   "Service"
               ForeColor       =   -2147483630
               BackColor       =   -2147483633
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               BevelOuter      =   0
               Font3D          =   2
            End
         End
         Begin Threed.SSCommand SSCommand2 
            Height          =   705
            Left            =   -74850
            TabIndex        =   262
            Top             =   4800
            Width           =   2595
            _Version        =   65536
            _ExtentX        =   4577
            _ExtentY        =   1244
            _StockProps     =   78
            Caption         =   "Execute"
            ForeColor       =   12582912
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   24
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   2
         End
         Begin Threed.SSPanel SSPanel11 
            Height          =   1275
            Left            =   -74910
            TabIndex        =   263
            Top             =   360
            Visible         =   0   'False
            Width           =   2655
            _Version        =   65536
            _ExtentX        =   4683
            _ExtentY        =   2249
            _StockProps     =   15
            Caption         =   "Control Trans"
            ForeColor       =   255
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   15.75
               Charset         =   177
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
            Alignment       =   6
            Begin VB.OptionButton optControMode 
               Caption         =   "Local"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   2
               Left            =   180
               TabIndex        =   265
               Top             =   510
               Value           =   -1  'True
               Width           =   1815
            End
            Begin VB.OptionButton optControMode 
               Caption         =   "Remote"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   13.5
                  Charset         =   177
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   3
               Left            =   180
               TabIndex        =   264
               Top             =   840
               Width           =   2055
            End
         End
         Begin Threed.SSFrame FrameRunMode 
            Height          =   3075
            Left            =   120
            TabIndex        =   266
            Top             =   420
            Width           =   1530
            _Version        =   65536
            _ExtentX        =   2699
            _ExtentY        =   5424
            _StockProps     =   14
            ForeColor       =   64
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Font3D          =   4
            ShadowStyle     =   1
            Begin VB.CommandButton CmdSysOperation 
               Caption         =   "Pause"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Index           =   1
               Left            =   180
               Picture         =   "frmMain.frx":CB210
               Style           =   1  'Graphical
               TabIndex        =   268
               Top             =   465
               Width           =   1164
            End
            Begin VB.CommandButton CmdSysOperation 
               Caption         =   "Resume"
               BeginProperty Font 
                  Name            =   "Microsoft Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Index           =   0
               Left            =   210
               Picture         =   "frmMain.frx":CC0DA
               Style           =   1  'Graphical
               TabIndex        =   267
               Top             =   1755
               Width           =   1164
            End
         End
      End
      Begin Threed.SSFrame SSFrame9 
         Height          =   1215
         Left            =   360
         TabIndex        =   269
         Top             =   9630
         Width           =   14535
         _Version        =   65536
         _ExtentX        =   25638
         _ExtentY        =   2143
         _StockProps     =   14
         Caption         =   "Robot Status"
         ForeColor       =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Alignment       =   2
         Font3D          =   4
         Begin Threed.SSPanel txtCycleName 
            Height          =   435
            Left            =   195
            TabIndex        =   270
            Top             =   645
            Width           =   4515
            _Version        =   65536
            _ExtentX        =   7964
            _ExtentY        =   767
            _StockProps     =   15
            ForeColor       =   -2147483630
            BackColor       =   12640511
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel txtCycleStep 
            Height          =   435
            Left            =   4800
            TabIndex        =   271
            Top             =   645
            Width           =   1260
            _Version        =   65536
            _ExtentX        =   2222
            _ExtentY        =   767
            _StockProps     =   15
            ForeColor       =   -2147483630
            BackColor       =   12640511
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
            Font3D          =   2
         End
         Begin Threed.SSPanel txtCycleMessage 
            Height          =   435
            Index           =   1
            Left            =   6360
            TabIndex        =   278
            Top             =   645
            Width           =   7125
            _Version        =   65536
            _ExtentX        =   12568
            _ExtentY        =   767
            _StockProps     =   15
            Caption         =   "0"
            ForeColor       =   -2147483630
            BackColor       =   12640511
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Microsoft Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BevelOuter      =   1
            BevelInner      =   1
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Caption         =   "Robot Status :"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   8880
            TabIndex        =   277
            Top             =   300
            Width           =   1920
         End
         Begin VB.Label lblStatus 
            Caption         =   "Cycle Name"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   1
            Left            =   240
            TabIndex        =   273
            Top             =   300
            Width           =   1440
         End
         Begin VB.Label lblStatus 
            Alignment       =   2  'Center
            Caption         =   "Line Number"
            BeginProperty Font 
               Name            =   "Microsoft Sans Serif"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   2
            Left            =   4680
            TabIndex        =   272
            Top             =   300
            Width           =   1440
         End
      End
      Begin VB.Label LabelApp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Work Piece"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Index           =   4
         Left            =   6600
         TabIndex        =   400
         Top             =   2700
         Width           =   2160
      End
      Begin VB.Label LabelApp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Pocket"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Index           =   3
         Left            =   6600
         TabIndex        =   282
         Top             =   2225
         Width           =   2160
      End
      Begin VB.Label LabelApp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Tool"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Index           =   0
         Left            =   6600
         TabIndex        =   276
         Top             =   1275
         Width           =   2160
      End
      Begin VB.Label LabelApp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Manual"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Index           =   2
         Left            =   6600
         TabIndex        =   275
         Top             =   1750
         Width           =   2160
      End
      Begin VB.Label LabelApp 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Diameter"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   420
         Index           =   1
         Left            =   6600
         TabIndex        =   274
         Top             =   800
         Width           =   2160
      End
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   1368
      Top             =   1692
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Align           =   1  'Align Top
      Height          =   0
      Left            =   0
      ScaleHeight     =   0
      ScaleWidth      =   15195
      TabIndex        =   0
      Top             =   0
      Width           =   15195
   End
End
Attribute VB_Name = "mdiMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Any) As Long
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hWnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

''for the incremental move
Dim IncStep As Double
Dim IncVel  As Integer


Private Sub AboutApplication_Click()
    FrmAbout.Show
End Sub

Private Sub BtnAddWrokPiece_Click()
''1.the function add new workpiece to the AllWp() Array.
''2.the f add the WP to the memory and to the CSV file.

Dim WorkPiece As Integer
Dim NCProgram As Long
Dim diameter As Integer
Dim ToolAmount As Integer
Dim ret As Integer
Dim line As Integer
Dim TempWP As WorkPiece
Dim ii As Integer
Dim LegalAmount As Boolean
Dim BoolAllDiameters

On Error GoTo error

    'get info from GUI
    TempWP.WPNumber = CInt(txtWorkPiece(2).Text)
    TempWP.NCProgram = CDbl(txtWorkPiece(3).Text)
    TempWP.ToolDiameter = CInt(txtWorkPiece(4).Text)
    TempWP.ToolAmount = CInt(txtWorkPiece(5).Text)
    TempWP.LineNumber = CInt(txtWorkPiece(6).Text)
    TempWP.WPToolType = (txtWorkPiece(7).Text)
    
    If ((TempWP.LineNumber > 1) And (AllWP(TempWP.LineNumber - 1).WPNumber = 0)) Then
        MsgBox "the line number is wrong.please insert a consecutive number."
        Exit Sub
    End If
    
    If TempWP.WPToolType = "HSK" Then
        ''if the diameter is between 100 to 300
        If ((TempWP.ToolDiameter < 100) Or (TempWP.ToolDiameter > 300)) Then
            GoTo error
        End If
    
    ElseIf TempWP.WPToolType = "Drill" Then
        If TempWP.ToolDiameter < 1 Or _
            TempWP.ToolDiameter > 7 Then
            GoTo error
        End If
        
    ElseIf TempWP.WPToolType = "Round" Then
        If TempWP.ToolDiameter < 1 Or _
            TempWP.ToolDiameter > 8 Then
                GoTo error
        End If
        
    End If
    
    AppDiameter = CInt(TempWP.ToolDiameter)
    LabelApp(1).Caption = "Diameter :" & AppDiameter
    
    If AllWP(1).ToolDiameter <> 0 Then
        If (CInt(TempWP.ToolDiameter)) <> AllWP(1).ToolDiameter Then
            ret = MsgBox("The diameter is wrong" & vbCrLf _
            & "insert another Diameter", vbCritical, _
            "Work Piece diameter ")
            txtWorkPiece(4).Text = ""
            Exit Sub
        End If
    End If
    
    ''the lineNumber is the same as the Index in the Array.
    ''add workpiece to the list in memory and to the csv file.
    Call modAllWorkPiece.WorkPieceEdit(TempWP.LineNumber, TempWP)
    Call ModCsvFile.SaveAllWP ''         save the data array into file
    Call BtnDisplayWPTable_Click   '     refresh the work piece Table

    ImgFdBkWPiece.Visible = True ''feedBack the User
    LegalAmount = modAllWorkPiece.CheckToolAmount()
    If LegalAmount = False Then
'        ret = MsgBox("The total amount is  over the limit" & vbCrLf _
'            & "Press OK to Continue.", vbInformation _
'                , "Tool Amount in Order")
    End If
    
    If AppToolType <> HSK Then
        Call CmdSetOffset_Click
    End If
    
    ModLogFile.LogAddLine ("add workpiece :" & TempWP.WPNumber)
    Call ModGUI.ListHMIUpdate("add workpiece :" & TempWP.WPNumber)
    
    Exit Sub
    
error:

     Call FrmDialog.ShowDialogForm(22, 14, 16, "MdiMain", "BtnAddWrokPiece_Click()", GeneralError)

End Sub

Private Sub BtnAllStatus_Click()
''1.the function change WPnumber/diameter/status/NCprogram for group of pockets.
''2.the function take/return no parameters.
''3.the function is being called from the GUI only.

Dim ret As Integer
Dim MyStatus As Integer
Dim MyWPNumber As Integer
Dim MyWPLine As Integer
Dim kk As Integer
Dim shelf As Integer
Dim column As Integer
Dim LastColumn As Integer
Dim PocketCounter As Integer
Dim BoolWPExist As Boolean
Dim BoolWPAmount As Integer
Dim BoolWPFirstTime As Boolean

    On Error GoTo label
    
    AppDiameter = AllWP(AppWPIndex).ToolDiameter
        
    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If
    
    PocketCounter = 0
    BoolWPAmount = True
    BoolWPFirstTime = True
    
    'chack for inputs not empty
    If textChangeWorkPiece.Text = "" Then
        Call FrmDialog.ShowDialogForm(32, 31, 31, "mdiMain", "BtnAllStatus_Click()", InputIncomplete)
        Exit Sub
    End If
    
    'chack for inputs has choosen
    If ComboAllStatus.Text = "" Then
        Call FrmDialog.ShowDialogForm(32, 31, 31, "mdiMain", "BtnAllStatus_Click()", InputIncomplete)
        Exit Sub
    End If
    
    'chack if the inputs is a number
    If IsNumeric(CInt(textChangeWorkPiece.Text)) = False Then
        Call FrmDialog.ShowDialogForm(32, 31, 31, "mdiMain", "BtnAllStatus_Click()", InputIncomplete)
        Exit Sub
    End If
    
    ''get info from GUI.
    MyStatus = ComboAllStatus.ListIndex + 1
    MyWPNumber = CInt(textChangeWorkPiece.Text)
    
    ''check if the WPiece Exist in the Allwp() array
    For kk = 1 To 50
        If AllWP(kk).WPNumber = MyWPNumber Then
            BoolWPExist = True
            Exit For
        Else
            BoolWPExist = False
            
        End If
    Next
    
    If (BoolWPExist = False) Then
        ret = MsgBox("the Work Piece number does not Exist." & vbCrLf & _
        "Please select a different number." _
        , vbExclamation _
        , "BtnAllStatus_Click()")
        textChangeWorkPiece.Text = ""
        Exit Sub
    End If
    
    
    ''get the current WorkPiece line number
    For kk = 1 To 50
        If AllWP(kk).WPNumber = MyWPNumber Then
            MyWPLine = kk
            Exit For
        End If
    Next
    
    ''check if SetNew/Edit Workpiece.
    For shelf = 1 To 3
        For column = 1 To LastColumn
            If AutomationStatus(shelf, column).WorkPiece = MyWPNumber Then
                BoolWPFirstTime = False
                Exit For
            Else
'                BoolWPFirstTime = True
'                Exit For
            End If
        Next
    Next
    
    ''if the workpiece already exist in the status table then replace parameters.
    If BoolWPFirstTime = False Then
        For shelf = 1 To 3
            If AppShelvs(shelf).ShelfToolType = AppToolType Then
                For column = 1 To LastColumn
                    
                   If AutomationStatus(shelf, column).WorkPiece = MyWPNumber Then
        
                        AutomationStatus(shelf, column).Status = MyStatus
                        AutomationStatus(shelf, column).ProgramNumber = AllWP(MyWPLine).NCProgram
                        AutomationStatus(shelf, column).diameter = AllWP(MyWPLine).ToolDiameter
                        AutomationStatus(shelf, column).WorkPiece = AllWP(MyWPLine).WPNumber
        
                        If AllWP(MyWPLine).WPToolType = "HSK" Then
                            AutomationStatus(shelf, column).CurrentTool = HSK
        
                        ElseIf AllWP(MyWPLine).WPToolType = "Drill" Then
                            AutomationStatus(shelf, column).CurrentTool = Drill
        
                        ElseIf AllWP(MyWPLine).WPToolType = "Round" Then
                            AutomationStatus(shelf, column).CurrentTool = Round
                        End If
                        BoolWPExist = True
        
                    End If
        
                Next
            End If
                
        Next
    
    ''if the WP does not exist in the status table - it is the first time to set WP.
    ElseIf BoolWPFirstTime = True Then
        For shelf = 1 To 3
            If AppShelvs(shelf).ShelfToolType = AppToolType Then
            For column = 1 To LastColumn
                
                If AutomationStatus(shelf, column).WorkPiece = 0 Then
                    If BoolWPAmount = True Then
                        AutomationStatus(shelf, column).WorkPiece = AllWP(MyWPLine).WPNumber
                        AutomationStatus(shelf, column).Status = MyStatus
                        AutomationStatus(shelf, column).ProgramNumber = AllWP(MyWPLine).NCProgram
                        AutomationStatus(shelf, column).diameter = AllWP(MyWPLine).ToolDiameter
                        
                        If AllWP(MyWPLine).WPToolType = "HSK" Then
                            AutomationStatus(shelf, column).CurrentTool = HSK
                        ElseIf AllWP(MyWPLine).WPToolType = "Drill" Then
                            AutomationStatus(shelf, column).CurrentTool = Drill
                        ElseIf AllWP(MyWPLine).WPToolType = "Round" Then
                            AutomationStatus(shelf, column).CurrentTool = Round
                        End If
                        
                        PocketCounter = PocketCounter + 1
                        If ((AppToolType = Drill) Or _
                            (AppToolType = Round) Or _
                            (AppToolType = HSK And AppDiameter <= 150)) Then
                            If PocketCounter = AllWP(MyWPLine).ToolAmount Then
                                BoolWPAmount = False
                            End If
                        End If
                        If (AppToolType = HSK And AppDiameter >= 151) Then
                            If PocketCounter = 2 * AllWP(MyWPLine).ToolAmount Then
                                BoolWPAmount = False
                            End If
                        End If
                    End If  ''BoolWPAmount
                End If ''AutomationStatus
                
            Next column
            End If
        Next shelf
    End If ''the workpiece exist
    
    If ((AppToolType = HSK) And (AppDiameter > 150)) Then
        ModAutomationStatus.SetBigHSKStatus
    End If
    
    Call SaveAutomation
    Call BtnDisplayPocketStatus_Click ''refresh the table
      
  Exit Sub
label:
Call FrmDialog.ShowDialogForm(32, 36, 36, "mdimain", "BtnAllStatus_Click()", GeneralError)

End Sub

Private Sub BtnCalcPoints_Click()
''1.this sub call the interpolation function for each shelf seperatly or 3 shelvs together.
''2.the function check the "AllShelvs" Condition.
''3.the function save data to CSV File.
''4.the function mark TRUE if the shelf has been taught correctly.
''5.the sub get no parameter.
''6.the sub return no parameter.
''7.the sub called from the GUI only.

Dim MyShelf As Integer
Dim ret As Integer

On Error GoTo labelerror

    MyShelf = CInt(TextShelfNum.Text)
       
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(MyShelf).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
   
    If AppShelvs(MyShelf).ShelfToolType = Drill Then
        Call ModMathDrill.DrillShelfCalculation(MyShelf)
        Call SaveArray("DrillLocations")
        
    ElseIf AppShelvs(MyShelf).ShelfToolType = Round Then
        Call ModMathRound.RoundShelfCalculation(MyShelf)
        Call SaveArray("RoundLocations")
        
    ElseIf AppShelvs(MyShelf).ShelfToolType = HSK Then
        Call ModMath.HSKPocketInterpolation(MyShelf)
        Call SaveArray("HSKLocations")
        
    End If
    ModLogFile.LogAddLine ("calculate shelf : " & CStr(MyShelf))
    Call ModGUI.ListHMIUpdate("calculate shelf : " & CStr(MyShelf))
    
    ret = MsgBox("calculate shelf number " & CStr(MyShelf) & " Done." _
        , vbInformation, _
            "BtnCalcPoints_Click()")
    
Exit Sub

labelerror:
    Call FrmDialog.ShowDialogForm(37, 37, 37, "MdiMain", "BtnCalcPoints_Click()", GeneralError)

End Sub



Private Sub BtnDelOrder_Click()

Dim TempBool As Boolean
Dim shelf As Integer
Dim column As Integer
Dim LineNumber As Integer
Dim DeleteWP As Integer


    On Error GoTo error
    
    'chack if the inputs is a number
    TempBool = IsNumeric((TextLineNumber.Text))
    If TempBool = False Then
        GoTo error
        Exit Sub
    End If
    
    LineNumber = CInt(TextLineNumber.Text) ''get info from GUI
    
    If LineNumber = 1 Then
        ModHandShake.SetOneCommByte (28)
    End If
    

    ''send current line Number to erase
    Call modAllWorkPiece.WorkPieceReset(LineNumber)
    
    Call modAllWorkPiece.ReOrderAllWPiece
    Call ModCsvFile.SaveAllWP ''         save the data array into file
    Call BtnDisplayWPTable_Click   '     refresh the work piece Table
    
    ModLogFile.LogAddLine ("Delete workpiece :" & CStr(AllWP(LineNumber).WPNumber))
    Call ModGUI.ListHMIUpdate("Delete workpiece :" & CStr(AllWP(LineNumber).WPNumber))
    
    Exit Sub
error:
    Call FrmDialog.ShowDialogForm(21, 15, 15, "MdiMain", "BtnDelOrder_Click()", GeneralError)
End Sub

Public Sub BtnDisplayPocketStatus_Click()
''1. the function read the status of the pockets from file and display the data on screen.
''2. the function take and return no parameters.
Debug.Print "BtnDisplayPocketStatus_Click()"

Dim shelf As Integer
Dim column As Integer
Dim diameter As Integer
Dim raw As Integer
Dim line As Integer
Dim orient As Integer
Dim number As Integer
Dim ProgNumberO As Long
Dim StatusNumber As Integer
Dim WorkPiece As Integer
   
 
    ModCsvFile.ReadAutomationStatus
    shelf = CInt(TextShelf.Text)
    
    LabelShelfType.Caption = AppShelvs(shelf).ShelfName
    ''''********
    ''''  HSK
    ''''********
    
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        MSFlexGrid2.Height = 650
        Exit Sub
    End If
     
    If AppToolType = HSK Then
   
        MSFlexGrid2.Height = 3500
        
        column = 1
        For line = 1 To TotalHSK
        
            MSFlexGrid2.TextMatrix(line, 0) = CStr(shelf * 100 + line)
            MSFlexGrid2.TextMatrix(line, 1) = AutomationStatus(shelf, column).WorkPiece
            MSFlexGrid2.TextMatrix(line, 2) = AutomationStatus(shelf, column).diameter
            MSFlexGrid2.TextMatrix(line, 4) = AutomationStatus(shelf, column).ProgramNumber
            

                
            If (AutomationStatus(shelf, column).Status = 1) Then
                MSFlexGrid2.TextMatrix(line, 3) = "empty"
            
            ElseIf (AutomationStatus(shelf, column).Status = 2) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Unmachined"
            
            ElseIf (AutomationStatus(shelf, column).Status = 3) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Machined"
            
            ElseIf (AutomationStatus(shelf, column).Status = 4) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Reserved"
            
            ElseIf (AutomationStatus(shelf, column).Status = 5) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Mask"
            
            ElseIf (AutomationStatus(shelf, column).Status = 6) Then
                MSFlexGrid2.TextMatrix(line, 3) = "occupied"
                
            ElseIf (AutomationStatus(shelf, column).Status = 7) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Broken"
            
            ElseIf (AutomationStatus(shelf, column).Status = 8) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Disable"
                                
            End If
            column = column + 1
        Next
    ''''********
    ''''  Drill
    ''''********
    ElseIf AppToolType = Drill Then
    
        MSFlexGrid2.Height = 4050
        column = 1

        For line = 1 To TotalDRILL
            
            MSFlexGrid2.TextMatrix(line, 0) = CStr(shelf * 100 + line)
            MSFlexGrid2.TextMatrix(line, 1) = AutomationStatus(shelf, column).WorkPiece
            MSFlexGrid2.TextMatrix(line, 2) = AutomationStatus(shelf, column).diameter
            MSFlexGrid2.TextMatrix(line, 4) = AutomationStatus(shelf, column).ProgramNumber
            
            If (AutomationStatus(shelf, column).Status = 1) Then
                MSFlexGrid2.TextMatrix(line, 3) = "empty"
            
            ElseIf (AutomationStatus(shelf, column).Status = 2) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Unmachined"
            
            ElseIf (AutomationStatus(shelf, column).Status = 3) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Machined"
            
            ElseIf (AutomationStatus(shelf, column).Status = 4) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Reserved"
            
            ElseIf (AutomationStatus(shelf, column).Status = 5) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Mask"
            
            ElseIf (AutomationStatus(shelf, column).Status = 6) Then
                MSFlexGrid2.TextMatrix(line, 3) = "occupied"
                
            ElseIf (AutomationStatus(shelf, column).Status = 7) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Broken"
            
            ElseIf (AutomationStatus(shelf, column).Status = 8) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Disable"
                                
            End If
            
            column = column + 1

        Next
    ''''********
    ''''  round
    ''''********
        ElseIf AppToolType = Round Then
    
        MSFlexGrid2.Height = 4100
        column = 1

        For line = 1 To TotalROUND
            
            MSFlexGrid2.TextMatrix(line, 0) = CStr(shelf * 100 + line)
            MSFlexGrid2.TextMatrix(line, 1) = AutomationStatus(shelf, column).WorkPiece
            MSFlexGrid2.TextMatrix(line, 2) = AutomationStatus(shelf, column).diameter
            MSFlexGrid2.TextMatrix(line, 4) = AutomationStatus(shelf, column).ProgramNumber
            
            If (AutomationStatus(shelf, column).Status = 1) Then
                MSFlexGrid2.TextMatrix(line, 3) = "empty"
            
            ElseIf (AutomationStatus(shelf, column).Status = 2) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Unmachined"
            
            ElseIf (AutomationStatus(shelf, column).Status = 3) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Machined"
            
            ElseIf (AutomationStatus(shelf, column).Status = 4) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Reserved"
            
            ElseIf (AutomationStatus(shelf, column).Status = 5) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Mask"
            
            ElseIf (AutomationStatus(shelf, column).Status = 6) Then
                MSFlexGrid2.TextMatrix(line, 3) = "occupied"
                
            ElseIf (AutomationStatus(shelf, column).Status = 7) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Broken"
            
            ElseIf (AutomationStatus(shelf, column).Status = 8) Then
                MSFlexGrid2.TextMatrix(line, 3) = "Disable"
                                
            End If
            
            column = column + 1

        Next

    End If

End Sub

Public Sub BtnDisplayWPTable_Click()
''1.the function display AllWP() data on screen.
Debug.Print "BtnDisplayWPTable_Click()"

    
Dim ii As Integer

    For ii = 1 To 50
    
        MSFlexGrid1.TextMatrix(ii, 0) = ii 'AllWP(ii).LineNumber
        MSFlexGrid1.TextMatrix(ii, 1) = AllWP(ii).WPNumber
        MSFlexGrid1.TextMatrix(ii, 2) = AllWP(ii).NCProgram
        MSFlexGrid1.TextMatrix(ii, 3) = AllWP(ii).ToolDiameter
        MSFlexGrid1.TextMatrix(ii, 4) = AllWP(ii).ToolAmount
        MSFlexGrid1.TextMatrix(ii, 5) = AllWP(ii).WPToolType
        
    Next
    
End Sub

Private Sub BtnJog_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Dim ret As Integer
On Error GoTo labelJog

    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(55, 12, 12, "MdiMain", "BtnJog_MouseDown()", KeyPosition)
        Exit Sub
    End If
    
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(55, 11, 26, "mdiMain", "BtnJog_MouseDown()", EmergencyStop)
        Exit Sub
    End If
    

    Call MoveIncremental(Index)
        
    Exit Sub
labelJog:
   ret = MsgBox("The System was unable to Jog the Robot," & vbCrLf _
        & "possible Causes are :" & vbCrLf _
        & "1.Vars mismatch" & vbCrLf _
        & "2.Communication problem with the Controller ," & vbCrLf _
        & "  [MdiMain->BtnJog_MouseDown]", _
        vbInformation, "Error in Jogging Robot.")
        
End Sub


                
    

Public Sub MoveIncremental(Index As Integer)

Dim ret As Integer
Dim CHigh As Integer
Dim CLow As Integer

    CHigh = 6
    CLow = 2
    
        Call SetServo(True)
        Call ClearIncrementalTarget ''clear the entire Array

    Select Case Index

        Case 0
            IncrementalTarget(0) = IncStep
            Call IncrementalMove(CHigh / 100 * AppSpeed, Index)

        Case 1
            IncrementalTarget(0) = -IncStep
            Call IncrementalMove(CHigh / 100 * AppSpeed, Index)
            
        Case 2
            IncrementalTarget(1) = IncStep
            Call IncrementalMove(CHigh / 100 * AppSpeed, Index)
            
        Case 3
            IncrementalTarget(1) = -IncStep
            Call IncrementalMove(CHigh / 100 * AppSpeed, Index)
            
        Case 4
            IncrementalTarget(2) = IncStep
            Call IncrementalMove(CHigh / 100 * AppSpeed, Index)
            
        Case 5
            IncrementalTarget(2) = -IncStep
            Call IncrementalMove(CHigh / 100 * AppSpeed, Index)
            
        Case 6
            IncrementalTarget(3) = IncStep
            Call IncrementalMove(CLow / 100 * AppSpeed, Index)

        Case 7
            IncrementalTarget(3) = -IncStep
            Call IncrementalMove(CLow / 100 * AppSpeed, Index)
            
        Case 8
            IncrementalTarget(4) = IncStep
            Call IncrementalMove(CLow / 100 * AppSpeed, Index)
            
        Case 9
            IncrementalTarget(4) = -IncStep
            Call IncrementalMove(CLow / 100 * AppSpeed, Index)
            
        Case 10
            IncrementalTarget(5) = IncStep
            Call IncrementalMove(CLow / 100 * AppSpeed, Index)
            
        Case 11
            IncrementalTarget(5) = -IncStep
            Call IncrementalMove(CLow / 100 * AppSpeed, Index)

    End Select
    Exit Sub


labelJog:
   ret = MsgBox("The System was unable to Jog the Robot," & vbCrLf _
        & "possible Causes are :" & vbCrLf _
        & "1.Vars mismatch" & vbCrLf _
        & "2.Communication problem with the Controller ," & vbCrLf _
        & "  [MdiMain->MoveIncremental]", _
        vbInformation, "Error in Jogging Robot.")
End Sub


Public Sub BtnShowFirstLastPockets_Click()

'1.the function display on screen the values of the
'   first and the last pocket of the selected shelf.
'2.the shelf taken from the GUI.

Dim shelf As Integer
Dim ret As Integer
Dim Xcord As Double
Dim Ycord As Double
Dim Zcord As Double

On Error GoTo error
    If TextShelfNum.Text = "" Then
        ret = MsgBox("Shelf number can not be empty", vbInformation, "Shelf Number")
        Exit Sub
    End If
    
        
        
        shelf = CInt(TextShelfNum.Text)
        LabelShelfType2.Caption = AppShelvs(shelf).ShelfName
        If AppShelvs(shelf).ShelfToolType = Drill Then
    
            ''display the first pocket
            Xcord = DrillLocations(shelf, 1).diameter(1).X
            TableTeach.TextMatrix(1, 1) = Format(Xcord, "000.000")
            
            Ycord = DrillLocations(shelf, 1).diameter(1).Y
            TableTeach.TextMatrix(1, 2) = Format(Ycord, "000.000")
    
            Zcord = DrillLocations(shelf, 1).diameter(1).z
            TableTeach.TextMatrix(1, 3) = Format(Zcord, "000.000")
            
             ''display the last pocket
            Xcord = DrillLocations(shelf, 12).diameter(1).X
            TableTeach.TextMatrix(2, 1) = Format(Xcord, "000.000")
            
            Ycord = DrillLocations(shelf, 12).diameter(1).Y
            TableTeach.TextMatrix(2, 2) = Format(Ycord, "000.000")
    
            Zcord = DrillLocations(shelf, 12).diameter(1).z
            TableTeach.TextMatrix(2, 3) = Format(Zcord, "000.000")
        
         ElseIf AppShelvs(shelf).ShelfToolType = HSK Then
        
            ''display the first pocket
            Xcord = HSKLocations(shelf, 1).X
            TableTeach.TextMatrix(1, 1) = Format(Xcord, "000.000")
            
            Ycord = HSKLocations(shelf, 1).Y
            TableTeach.TextMatrix(1, 2) = Format(Ycord, "000.000")
    
            Zcord = HSKLocations(shelf, 1).z
            TableTeach.TextMatrix(1, 3) = Format(Zcord, "000.000")
            
             ''display the last pocket
            Xcord = HSKLocations(shelf, 10).X
            TableTeach.TextMatrix(2, 1) = Format(Xcord, "000.000")
            
            Ycord = HSKLocations(shelf, 10).Y
            TableTeach.TextMatrix(2, 2) = Format(Ycord, "000.000")
    
            Zcord = HSKLocations(shelf, 10).z
            TableTeach.TextMatrix(2, 3) = Format(Zcord, "000.000")
            
          ElseIf AppShelvs(shelf).ShelfToolType = Round Then
        
            ''display the first pocket
            Xcord = RoundLocations(shelf, 1).diameter(1).X
            TableTeach.TextMatrix(1, 1) = Format(Xcord, "000.000")
            
            Ycord = RoundLocations(shelf, 1).diameter(1).Y
            TableTeach.TextMatrix(1, 2) = Format(Ycord, "000.000")
    
            Zcord = RoundLocations(shelf, 1).diameter(1).z
            TableTeach.TextMatrix(1, 3) = Format(Zcord, "000.000")
            
             ''display the last pocket
            Xcord = RoundLocations(shelf, 12).diameter(1).X
            TableTeach.TextMatrix(2, 1) = Format(Xcord, "000.000")
            
            Ycord = RoundLocations(shelf, 12).diameter(1).Y
            TableTeach.TextMatrix(2, 2) = Format(Ycord, "000.000")
    
            Zcord = RoundLocations(shelf, 12).diameter(1).z
            TableTeach.TextMatrix(2, 3) = Format(Zcord, "000.000")
            
        End If
        
         
    Exit Sub
                           
error:
       ret = MsgBox("The System was unable to Show The Data," & vbCrLf _
        & "1.Check the Inputs.should not be empty." & vbCrLf _
        & "2.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.BtnShowPockets_Click]", vbInformation, "Error in Show Pockets.")
   

End Sub

Public Sub BtnShowGeneralLocations_Click()
''1.the function display General Location data on the screen.
''2.the f read the data from CSV file
''3.the f display the data on the screen.
Debug.Print "BtnShowGeneralLocations_Click()"

Dim line As Integer
Dim column As Integer

    ModCsvFile.ReadGeneralLocation
    
    
    MSFlexGrid3.TextMatrix(0, 0) = "Number"
    MSFlexGrid3.TextMatrix(0, 1) = "Name"
    MSFlexGrid3.TextMatrix(0, 2) = "X"
    MSFlexGrid3.TextMatrix(0, 3) = "Y"
    MSFlexGrid3.TextMatrix(0, 4) = "Z"
    
    MSFlexGrid3.TextMatrix(1, 0) = "10"
    MSFlexGrid3.TextMatrix(1, 1) = "Spindle"
    MSFlexGrid3.TextMatrix(1, 2) = GeneralLocation(10).X
    MSFlexGrid3.TextMatrix(1, 3) = GeneralLocation(10).Y
    MSFlexGrid3.TextMatrix(1, 4) = GeneralLocation(10).z
    
    MSFlexGrid3.TextMatrix(2, 0) = "11"
    MSFlexGrid3.TextMatrix(2, 1) = "Kiosk"
    MSFlexGrid3.TextMatrix(2, 2) = GeneralLocation(11).X
    MSFlexGrid3.TextMatrix(2, 3) = GeneralLocation(11).Y
    MSFlexGrid3.TextMatrix(2, 4) = GeneralLocation(11).z
     
    MSFlexGrid3.TextMatrix(3, 0) = "12"
    MSFlexGrid3.TextMatrix(3, 1) = "Chuck"
    MSFlexGrid3.TextMatrix(3, 2) = GeneralLocation(12).X
    MSFlexGrid3.TextMatrix(3, 3) = GeneralLocation(12).Y
    MSFlexGrid3.TextMatrix(3, 4) = GeneralLocation(12).z
    
    MSFlexGrid3.TextMatrix(4, 0) = "120"
    MSFlexGrid3.TextMatrix(4, 1) = "Station 1"
    MSFlexGrid3.TextMatrix(4, 2) = GeneralLocation(120).X
    MSFlexGrid3.TextMatrix(4, 3) = GeneralLocation(120).Y
    MSFlexGrid3.TextMatrix(4, 4) = GeneralLocation(120).z
    
    MSFlexGrid3.TextMatrix(5, 0) = "121"
    MSFlexGrid3.TextMatrix(5, 1) = "Station 2"
    MSFlexGrid3.TextMatrix(5, 2) = GeneralLocation(121).X
    MSFlexGrid3.TextMatrix(5, 3) = GeneralLocation(121).Y
    MSFlexGrid3.TextMatrix(5, 4) = GeneralLocation(121).z

End Sub



Private Sub BtnSingleStatus_Click()
''1.this function change status for a single pockets.

Dim ret As Integer
Dim Status As Integer
Dim shelf As Integer
Dim diameter As Integer
Dim PocketNumber As Integer
Dim OldStatus As PocketStatus


    On Error GoTo label
        
     shelf = CInt(fMainForm.TextShelf.Text)
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        ret = MsgBox("the shelf number is incorrect." & vbCrLf _
            & "the shelf type is not the same as the gripper type" _
            , vbCritical _
            , "BtnSingleStatus_Click()")
        Exit Sub
    End If
    
    
    'chack for inputs not empty
    If TxtSinglePocketNumber.Text = "" Then
        GoTo label
    End If

    'chack for inputs not empty (HSK)
    If AppToolType = HSK Then
        If TxtSingleToolDiameter.Text = "" Then
            GoTo label
        End If
    End If

    'chack for inputs has choosen
    If ComboSingleStatus.Text = "Single Status" Then
        GoTo label
    End If
    
    'chack if the inputs is a number
    If IsNumeric(CInt(TxtSinglePocketNumber.Text)) = False Then
        GoTo label
    End If
    
    'chack if the inputs is a number
    If AppToolType = HSK Then
        If IsNumeric(CInt(TxtSingleToolDiameter.Text)) = False Then
            GoTo label
        End If
    End If
    
    'get info from the user
    Status = CInt(ComboSingleStatus.ListIndex) + 1
    shelf = CInt(TextShelf.Text)
    PocketNumber = CInt(TxtSinglePocketNumber.Text)
    
    If AppToolType = HSK Then
        If PocketNumber >= 101 And PocketNumber <= 110 Then
            PocketNumber = PocketNumber - 100
        ElseIf PocketNumber >= 201 And PocketNumber <= 210 Then
            PocketNumber = PocketNumber - 200
        ElseIf PocketNumber >= 301 And PocketNumber <= 310 Then
            PocketNumber = PocketNumber - 300
        End If
        diameter = 0
        
    ElseIf AppToolType = Drill Then
        If PocketNumber >= 101 And PocketNumber <= 112 Then
            PocketNumber = PocketNumber - 100
        ElseIf PocketNumber >= 201 And PocketNumber <= 212 Then
            PocketNumber = PocketNumber - 200
        ElseIf PocketNumber >= 301 And PocketNumber <= 312 Then
            PocketNumber = PocketNumber - 300
        End If
        diameter = CInt(TxtSingleToolDiameter.Text)
        
    ElseIf AppToolType = Round Then
        If PocketNumber >= 101 And PocketNumber <= 112 Then
            PocketNumber = PocketNumber - 100
        ElseIf PocketNumber >= 201 And PocketNumber <= 212 Then
            PocketNumber = PocketNumber - 200
        ElseIf PocketNumber >= 301 And PocketNumber <= 312 Then
            PocketNumber = PocketNumber - 300
        End If
        diameter = CInt(TxtSingleToolDiameter.Text)
        
    End If
    
    OldStatus = GetPocketStatus(shelf, PocketNumber)
    If OldStatus = Mask Then
        ret = MsgBox("the current status is : MASK" & vbCrLf _
        & "Do you want to continue  ? " _
        , vbQuestion + vbYesNo, _
        "Change status for single pocket")
        
        If ret = 7 Then ''     no
            Exit Sub
        ElseIf ret = 6 Then '''yes
            Call SetPocketStatus(Status, CInt(TxtSinglePocketNumber.Text))
        End If
        
    End If
    Call SetPocketStatus(Status, CInt(TxtSinglePocketNumber.Text))
    
    
    
    Call BtnDisplayPocketStatus_Click ''refresh the table
    
  Exit Sub
label:
ret = MsgBox("The System was unable to Change a single Status," & vbCrLf _
    & "1.Check the Inputs.should not be empty." & vbCrLf _
    & "2.Check the Inputs.should not be 'single status'." & vbCrLf _
    & "  [MdinMain.BtnSingleStatus_Click]" _
    , vbInformation _
    , "Error in Change single Status")

End Sub

Public Sub BtnStartLoadUnload_Click()

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim ret As Integer
Dim ii As Integer
Dim bb As Integer
Dim TempBool As Boolean


    On Error GoTo label
    
    AppDiameter = CInt(TextDrillCode.Text)
    
    ''check the key state
    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(19, 12, 12, "MdiMain", "BtnStartLoadUnload_Click()", KeyPosition)
        Exit Sub
    End If
    
    ''get info from GUI
    shelf = CInt(Left(txtTestPocket(2).Text, 1)) ''get info from GUI
    column = CInt(Right(txtTestPocket(2).Text, 2))
    pocket = CInt(TextDrillCode.Text)
     
      
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
    
    If AppToolType = Drill Then
        ArrayPosition(0) = 1    ''1=XYZ
        ArrayPosition(1) = 1 ''1=robot cordinate
        ArrayPosition(2) = DrillLocations(shelf, column).diameter(pocket).X ''assign Values
        ArrayPosition(3) = DrillLocations(shelf, column).diameter(pocket).Y
        ArrayPosition(4) = DrillLocations(shelf, column).diameter(pocket).z
        ArrayPosition(5) = DrillLocations(shelf, column).diameter(pocket).Rx
        ArrayPosition(6) = DrillLocations(shelf, column).diameter(pocket).Ry
        ArrayPosition(7) = DrillLocations(shelf, column).diameter(pocket).Rz
        
        Call WriteOneCommByte(52, tooltype.Drill)
        
    ElseIf AppToolType = HSK Then
        ArrayPosition(0) = 1    ''1=XYZ
        ArrayPosition(1) = 1 ''1=robot cordinate
        ArrayPosition(2) = HSKLocations(shelf, column).X
        ArrayPosition(3) = HSKLocations(shelf, column).Y
        ArrayPosition(4) = HSKLocations(shelf, column).z
        ArrayPosition(5) = HSKLocations(shelf, column).Rx
        ArrayPosition(6) = HSKLocations(shelf, column).Ry
        ArrayPosition(7) = HSKLocations(shelf, column).Rz
        
         Call WriteOneCommByte(52, tooltype.HSK)
        
    ElseIf AppToolType = Round Then
        ArrayPosition(0) = 1 ''  1=XYZ
        ArrayPosition(1) = 1 ''  1=robot cordinate
        ArrayPosition(2) = RoundLocations(shelf, column).diameter(pocket).X ''assign Values
        ArrayPosition(3) = RoundLocations(shelf, column).diameter(pocket).Y
        ArrayPosition(4) = RoundLocations(shelf, column).diameter(pocket).z
        ArrayPosition(5) = RoundLocations(shelf, column).diameter(pocket).Rx
        ArrayPosition(6) = RoundLocations(shelf, column).diameter(pocket).Ry
        ArrayPosition(7) = RoundLocations(shelf, column).diameter(pocket).Rz
        
        Call WriteOneCommByte(52, tooltype.Round)
        
                
        ''correct the Z parameter.
        If pocket <= 4 Then
            PocketDepth = 86
        ElseIf pocket >= 5 Then
            PocketDepth = 92
        Else
            PocketDepth = 92
        End If
        
        ArrayPosition(4) = RoundLocations(shelf, column).diameter(pocket).z - PocketDepth + ChuckDepth - ChuckStopper + ShelfSafty

        
    End If
    
    Call MotoComToolBox.WritePosition(16, ArrayPosition) ''send info to controller
    Call MotoComToolBox.WritePosition(17, ArrayPosition)
    ModHandShake.SetOneCommByte (18)
    
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(19, 11, 26, "mdiMain", "BtnStartLoadUnload_Click()", GeneralError)
        Exit Sub
    End If
    
        
    TempBool = ModHandShake.WriteDrillCode(AppDiameter)
    If TempBool = False Then
        Call FrmDialog.ShowDialogForm(60, 60, 60, "mdiMain", "CmdSysOperation_Click()", GeneralError)
        Exit Sub
    End If
            
    TempBool = ModHandShake.SendToolSensorState()
    If TempBool = False Then
        Call FrmDialog.ShowDialogForm(61, 61, 61, "mdiMain", "CmdSysOperation_Click()", GeneralError)
        Exit Sub
    End If
   
    If AppToolType <> HSK Then
        Call CmdSetOffset_Click ''send offset to controller
    End If
   
    If CurrentTest = "3_TEST_POCKET_TO_CHUCK.JBI" Or _
        CurrentTest = "3_TEST_CHUCK_TO_POCKET.JBI" Or _
        CurrentTest = "3_TEST_KIOSK_TO_CHUCK.JBI" Or _
        CurrentTest = "3_TEST_CHUCK_TO_KIOSK.JBI" _
    Then
        ret = MsgBox("    Make sure the Hermle DOORS are open ." & vbCrLf _
            & "    Make sure the CHUCK is open ." & vbCrLf _
            & "    press OK to continue ." & vbCrLf _
            & "                          " & vbCrLf _
                , vbExclamation + vbOKCancel _
                , "BtnStartLoadUnload_Click()")
                
            If ret = 1 Then ''     1=OK
                ''do nothing
            ElseIf ret = 2 Then  ' 2=Cancel
                Exit Sub
            End If
                
                
    End If
            

    RunTestEndless = False
        
    TempBool = StartJob(CurrentTest)
    If TempBool = False Then
        Call FrmDialog.ShowDialogForm(19, 32, 32, "mdimain", "BtnStartLoadUnload_Click()", InputIncomplete)
        Exit Sub
    End If
    
     ModLogFile.LogAddLine ("Start Test : " & CurrentTest)
     Call ModGUI.ListHMIUpdate(" Start Test : " & CurrentTest)
        
Exit Sub
label:
   ret = FrmDialog.ShowDialogForm(19, 11, 11, "MdiMain", "BtnStartLoadUnload_Click()", GeneralError)
End Sub

Public Sub BtnStopLoadUnload_Click()

Dim jj As Integer


    RunTestEndless = False
    
    ''disable the pockets names buttons.
    For jj = 0 To 3
        Option1(jj).Value = False
    Next
    
    ''disable the OptionsLoops RadioButtons
    For jj = 0 To 3
        OptionLoop(jj).Value = vbUnchecked
    Next jj
    
     ''disable the Point-To-Point Loop RadioButtons
    fMainForm.CheckLoopP2P.Value = vbUnchecked
    
    CurrentTest = ""
    
    Call CmdResetAllTests_Click
    
    ModLogFile.LogAddLine (" Stop Loop cycle.")
    Call ModGUI.ListHMIUpdate("Stop Loop cycle.")
     
End Sub

Private Sub BtnTeachFirst_Click()

Dim shelf As Integer
Dim ret As Integer
On Error GoTo LabelTeachFirst

    shelf = CInt(fMainForm.TextShelfNum.Text)
    
    If HermleCommState = OffLine Then
        ret = MsgBox("the application is OffLine !" _
        & vbCrLf & "set communication to OnLine." _
        , vbExclamation, "can not teach first pocket")
        Exit Sub
    End If
    
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
    
    Call ReadCurrentPosition(TempPosition)
    
    If AppShelvs(shelf).ShelfToolType = Drill Then
        DrillLocations(shelf, 1).diameter(1).X = TempPosition(0)
        DrillLocations(shelf, 1).diameter(1).Y = TempPosition(1)
        DrillLocations(shelf, 1).diameter(1).z = TempPosition(2) - JigHeight
        DrillLocations(shelf, 1).diameter(1).Rx = TempPosition(3)
        DrillLocations(shelf, 1).diameter(1).Ry = TempPosition(4)
        DrillLocations(shelf, 1).diameter(1).Rz = TempPosition(5)
        Call SaveArray("DrillLocations")
    
    ElseIf AppShelvs(shelf).ShelfToolType = HSK Then
        HSKLocations(shelf, 1).X = TempPosition(0)
        HSKLocations(shelf, 1).Y = TempPosition(1)
        HSKLocations(shelf, 1).z = TempPosition(2)
        HSKLocations(shelf, 1).Rx = TempPosition(3)
        HSKLocations(shelf, 1).Ry = TempPosition(4)
        HSKLocations(shelf, 1).Rz = TempPosition(5)
        Call SaveArray("HSKLocations")
        
    ElseIf AppShelvs(shelf).ShelfToolType = Round Then
        RoundLocations(shelf, 1).diameter(1).X = TempPosition(0)
        RoundLocations(shelf, 1).diameter(1).Y = TempPosition(1)
        RoundLocations(shelf, 1).diameter(1).z = TempPosition(2) - JigHeight
        RoundLocations(shelf, 1).diameter(1).Rx = TempPosition(3)
        RoundLocations(shelf, 1).diameter(1).Ry = TempPosition(4)
        RoundLocations(shelf, 1).diameter(1).Rz = TempPosition(5)
        RoundLocations(shelf, 1).diameter(1).name = CStr(shelf) & "01.1"
        
        Call SaveArray("RoundLocations")
    
    End If
    
        
    ''save the location in the memory .
    If shelf = 1 Then
        GeneralLocation(21).X = TempPosition(0)
        GeneralLocation(21).Y = TempPosition(1)
        GeneralLocation(21).z = TempPosition(2)
        GeneralLocation(21).Rx = TempPosition(3)
        GeneralLocation(21).Ry = TempPosition(4)
        GeneralLocation(21).Rz = TempPosition(5)
        ''Call WritePosition(21, TempPosition) ''run over the value in the controller
       
    ElseIf shelf = 2 Then
        GeneralLocation(23).X = TempPosition(0)
        GeneralLocation(23).Y = TempPosition(1)
        GeneralLocation(23).z = TempPosition(2)
        GeneralLocation(23).Rx = TempPosition(3)
        GeneralLocation(23).Ry = TempPosition(4)
        GeneralLocation(23).Rz = TempPosition(5)
        ''Call WritePosition(23, TempPosition) ''run over the value in the controller
        
    ElseIf shelf = 3 Then
        GeneralLocation(25).X = TempPosition(0)
        GeneralLocation(25).Y = TempPosition(1)
        GeneralLocation(25).z = TempPosition(2)
        GeneralLocation(25).Rx = TempPosition(3)
        GeneralLocation(25).Ry = TempPosition(4)
        GeneralLocation(25).Rz = TempPosition(5)
        ''Call WritePosition(25, TempPosition) ''run over the value in the controller
        
    End If
          
    Call WriteGeneralLocations ''save the location in the TextFile.
    Call BtnShowFirstLastPockets_Click
    
    ModLogFile.LogAddLine (" teach first pocket in shelf : " & CStr(shelf))
    Call ModGUI.ListHMIUpdate("teach first pocket in shelf : " & CStr(shelf))
    
    Exit Sub
LabelTeachFirst:
    Call FrmDialog.ShowDialogForm(59, 59, 59, "mdimain", "BtnTeachFirst_Click()", GeneralError)

End Sub

Private Sub BtnTeachLast_Click()
''1.this function sample the current location from robot.
''2.the function save data in the memory
''3.the f save data in the CSV file.

Dim shelf As Integer
Dim ret As Integer
On Error GoTo LabelTeachLast

    shelf = CInt(fMainForm.TextShelfNum.Text)
       
           
    If HermleCommState = OffLine Then
        ret = MsgBox("the application is OffLine !" _
        & vbCrLf & "set communication to OnLine." _
        , vbExclamation, "can not teach last pocket")
        Exit Sub
    End If
    
    
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
    
    Call ReadCurrentPosition(TempPosition)

    If AppShelvs(shelf).ShelfToolType = Drill Then
        DrillLocations(shelf, 12).diameter(1).X = TempPosition(0)
        DrillLocations(shelf, 12).diameter(1).Y = TempPosition(1)
        DrillLocations(shelf, 12).diameter(1).z = TempPosition(2) - JigHeight
        DrillLocations(shelf, 12).diameter(1).Rx = TempPosition(3)
        DrillLocations(shelf, 12).diameter(1).Ry = TempPosition(4)
        DrillLocations(shelf, 12).diameter(1).Rz = TempPosition(5)
        Call SaveArray("DrillLocations")
    
    ElseIf AppShelvs(shelf).ShelfToolType = HSK Then
        HSKLocations(shelf, 10).X = TempPosition(0)
        HSKLocations(shelf, 10).Y = TempPosition(1)
        HSKLocations(shelf, 10).z = TempPosition(2)
        HSKLocations(shelf, 10).Rx = TempPosition(3)
        HSKLocations(shelf, 10).Ry = TempPosition(4)
        HSKLocations(shelf, 10).Rz = TempPosition(5)

        Call SaveArray("HSKLocations")
        
    ElseIf AppShelvs(shelf).ShelfToolType = Round Then
        RoundLocations(shelf, 12).diameter(1).X = TempPosition(0)
        RoundLocations(shelf, 12).diameter(1).Y = TempPosition(1)
        RoundLocations(shelf, 12).diameter(1).z = TempPosition(2) - JigHeight
        RoundLocations(shelf, 12).diameter(1).Rx = TempPosition(3)
        RoundLocations(shelf, 12).diameter(1).Ry = TempPosition(4)
        RoundLocations(shelf, 12).diameter(1).Rz = TempPosition(5)
          RoundLocations(shelf, 12).diameter(1).name = CStr(shelf) & "12.1"
        Call SaveArray("RoundLocations")
        
    End If

    ''save the location in file .
    If shelf = 1 Then
        GeneralLocation(22).X = TempPosition(0)
        GeneralLocation(22).Y = TempPosition(1)
        GeneralLocation(22).z = TempPosition(2)
        GeneralLocation(22).Rx = TempPosition(3)
        GeneralLocation(22).Ry = TempPosition(4)
        GeneralLocation(22).Rz = TempPosition(5)
        ''Call WritePosition(22, TempPosition) ''run over the value in the controller
        
    ElseIf shelf = 2 Then
        GeneralLocation(24).X = TempPosition(0)
        GeneralLocation(24).Y = TempPosition(1)
        GeneralLocation(24).z = TempPosition(2)
        GeneralLocation(24).Rx = TempPosition(3)
        GeneralLocation(24).Ry = TempPosition(4)
        GeneralLocation(24).Rz = TempPosition(5)
        ''Call WritePosition(24, TempPosition) ''run over the value in the controller
        
    ElseIf shelf = 3 Then
        GeneralLocation(26).X = TempPosition(0)
        GeneralLocation(26).Y = TempPosition(1)
        GeneralLocation(26).z = TempPosition(2)
        GeneralLocation(26).Rx = TempPosition(3)
        GeneralLocation(26).Ry = TempPosition(4)
        GeneralLocation(26).Rz = TempPosition(5)
        ''Call WritePosition(26, TempPosition) ''run over the value in the controller
        
    End If
    
    Call WriteGeneralLocations ''          save the location in a TextFile.
    Call BtnShowFirstLastPockets_Click
    ModLogFile.LogAddLine (" teach last pocket in shelf : " & CStr(shelf))
    Call ModGUI.ListHMIUpdate("teach last pocket in shelf : " & CStr(shelf))
    Exit Sub
    
LabelTeachLast:
    Call FrmDialog.ShowDialogForm(59, 59, 59, "mdimain", "BtnTeachFirst_Click()", GeneralError)

End Sub

Private Sub BtnTeachPosition_Click()

Dim ret As Integer

On Error GoTo LabelTeach

   
    Call ReadCurrentPosition(TempPosition) ''read the current position into TempPosition
     
    If ComboLocations.Text = "Kiosk" Then ''save position to the memory of PC.
        GeneralLocation(11).X = TempPosition(0)
        GeneralLocation(11).Y = TempPosition(1)
        GeneralLocation(11).z = TempPosition(2)
        GeneralLocation(11).Rx = TempPosition(3)
        GeneralLocation(11).Ry = TempPosition(4)
        GeneralLocation(11).Rz = TempPosition(5)
        Call MotoComToolBox.WritePosition(11, TempPosition)
        Call MotoComToolBox.StartJob("calc_kiosk_points.JBI")
        
    ElseIf ComboLocations.Text = "Chuck" Then
        GeneralLocation(12).X = TempPosition(0)
        GeneralLocation(12).Y = TempPosition(1)
        GeneralLocation(12).z = TempPosition(2)
        GeneralLocation(12).Rx = TempPosition(3)
        GeneralLocation(12).Ry = TempPosition(4)
        GeneralLocation(12).Rz = TempPosition(5)
        Call MotoComToolBox.WritePosition(12, TempPosition)
        Call MotoComToolBox.StartJob("calc_chuck_points.JBI")
        
    ElseIf ComboLocations.Text = "Spindle" Then
        GeneralLocation(10).X = TempPosition(0)
        GeneralLocation(10).Y = TempPosition(1)
        GeneralLocation(10).z = TempPosition(2)
        GeneralLocation(10).Rx = TempPosition(3)
        GeneralLocation(10).Ry = TempPosition(4)
        GeneralLocation(10).Rz = TempPosition(5)
        Call MotoComToolBox.WritePosition(10, TempPosition)

    ElseIf ComboLocations.Text = "Station 1" Then
        GeneralLocation(120).X = TempPosition(0)
        GeneralLocation(120).Y = TempPosition(1)
        GeneralLocation(120).z = TempPosition(2)
        GeneralLocation(120).Rx = TempPosition(3)
        GeneralLocation(120).Ry = TempPosition(4)
        GeneralLocation(120).Rz = TempPosition(5)
        Call MotoComToolBox.WritePosition(120, TempPosition)

    ElseIf ComboLocations.Text = "Station 2" Then
        GeneralLocation(121).X = TempPosition(0)
        GeneralLocation(121).Y = TempPosition(1)
        GeneralLocation(121).z = TempPosition(2)
        GeneralLocation(121).Rx = TempPosition(3)
        GeneralLocation(121).Ry = TempPosition(4)
        GeneralLocation(121).Rz = TempPosition(5)
        Call MotoComToolBox.WritePosition(121, TempPosition)

    ElseIf ComboLocations.Text = "Select Location" Then
        Call FrmDialog.ShowDialogForm(39, 39, 39, "MdiMain", "BtnTeachPosition_Click()", InputIncomplete)

   End If
      
    Delay 500
    
    Call WriteGeneralLocations
    Call fMainForm.BtnShowGeneralLocations_Click

     
Exit Sub

LabelTeach:
   ret = MsgBox("The System was unable to Teach general Location," & vbCrLf _
        & "1.Check the Inputs.should not be empty." & vbCrLf _
        & "2.Check the Inputs.should not be 'Select Location'." & vbCrLf _
        & "3.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.BtnTeachPosition_Click]", vbInformation, "Error in Teaching General Location.")
   
End Sub

Private Sub BtnTeachSingle_Click()

Dim shelf As Integer
Dim PocketNumber As Integer
Dim ToolDiameter As Integer
Dim ret As Integer
Dim column As Integer

    On Error GoTo labelsingle
    shelf = CInt(TextShelfNum.Text)
    
           
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
    

    If AppToolType = Drill Then
        ''**********
        ''**   Drill
        '***********
        ''check correct input from user: 'fields should be full
        If TextPocketNumber.Text = "" Or TextToolsDiameter.Text = "" Or TextShelfNum.Text = "" Then
            GoTo labelsingle
        End If
        
        ''check correct input from user:'fields should be integers
        If IsNumeric(TextPocketNumber.Text) = False Or IsNumeric(TextToolsDiameter.Text) = False Or IsNumeric(TextShelfNum.Text) = False Then
            GoTo labelsingle
        End If
        
        shelf = CInt(TextShelfNum.Text)
        
        PocketNumber = CInt(Right(TextPocketNumber.Text, 2))
        ToolDiameter = CInt(TextToolsDiameter.Text)
        
        Call ReadCurrentPosition(TempPosition)
        
        DrillLocations(shelf, PocketNumber).diameter(ToolDiameter).X = TempPosition(0)
        DrillLocations(shelf, PocketNumber).diameter(ToolDiameter).Y = TempPosition(1)
        DrillLocations(shelf, PocketNumber).diameter(ToolDiameter).z = TempPosition(2) - JigHeight
        DrillLocations(shelf, PocketNumber).diameter(ToolDiameter).Rx = TempPosition(3)
        DrillLocations(shelf, PocketNumber).diameter(ToolDiameter).Ry = TempPosition(4)
        DrillLocations(shelf, PocketNumber).diameter(ToolDiameter).Rz = TempPosition(5)
        Call SaveArray("DrillLocations")
    
    ElseIf AppToolType = HSK Then
        ''**********
        ''    HSK
        '***********
        ''check correct input from user: 'fields should be full
        If TextPocketNumber.Text = "" Or TextShelfNum.Text = "" Then
            GoTo labelsingle
        End If
        
        ''check correct input from user:'fields should be integers
        If IsNumeric(TextPocketNumber.Text) = False Or IsNumeric(TextShelfNum.Text) = False Then
            GoTo labelsingle
        End If
        
        shelf = CInt(Left(TextPocketNumber, 1))
        column = CInt(Right(TextPocketNumber, 2))
                
        Call ReadCurrentPosition(TempPosition)

        HSKLocations(shelf, column).X = TempPosition(0)
        HSKLocations(shelf, column).Y = TempPosition(1)
        HSKLocations(shelf, column).z = TempPosition(2)
        HSKLocations(shelf, column).Rx = TempPosition(3)
        HSKLocations(shelf, column).Ry = TempPosition(4)
        HSKLocations(shelf, column).Rz = TempPosition(5)
        Call SaveArray("HSKLocations")
        
    ElseIf AppToolType = Round Then
        '''***************
        '''  round tool
        '''***************
        
        ''check correct input from user: 'fields should be full
        If TextPocketNumber.Text = "" Or TextToolsDiameter.Text = "" Or TextShelfNum.Text = "" Then
            GoTo labelsingle
        End If
        
        ''check correct input from user:'fields should be integers
        If IsNumeric(TextPocketNumber.Text) = False Or IsNumeric(TextToolsDiameter.Text) = False Or IsNumeric(TextShelfNum.Text) = False Then
            GoTo labelsingle
        End If
        
        shelf = CInt(TextShelfNum.Text)
        
        PocketNumber = CInt(Right(TextPocketNumber.Text, 2))
        ToolDiameter = CInt(TextToolsDiameter.Text)
        
        Call ReadCurrentPosition(TempPosition)
        
        RoundLocations(shelf, PocketNumber).diameter(ToolDiameter).X = TempPosition(0)
        RoundLocations(shelf, PocketNumber).diameter(ToolDiameter).Y = TempPosition(1)
        RoundLocations(shelf, PocketNumber).diameter(ToolDiameter).z = TempPosition(2) - JigHeight
        RoundLocations(shelf, PocketNumber).diameter(ToolDiameter).Rx = TempPosition(3)
        RoundLocations(shelf, PocketNumber).diameter(ToolDiameter).Ry = TempPosition(4)
        RoundLocations(shelf, PocketNumber).diameter(ToolDiameter).Rz = TempPosition(5)
        Call SaveArray("RoundLocations")
        
    End If
    
    ''refresh the table :View shelfs locations.
    BtnViewPocketsLocations_Click
    
    ModLogFile.LogAddLine (" Teach single pocket : ")
    Call ModGUI.ListHMIUpdate("Teach single pocket : ")
    
    Exit Sub
    
labelsingle:
   ret = MsgBox("The System was unable to Teach Single Pocket," & vbCrLf _
        & "1.Check the Inputs.should not be empty." & vbCrLf _
        & "2.Check the Inputs.should be numbers only." & vbCrLf _
        & "3.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.BtnTeachSingle_Click]", vbInformation, "Error in Teaching Single Pocket.")
   
    
End Sub

Private Sub BtnTestAllPockets_Click()

Dim ret As Integer
Dim bb As Boolean
Dim shelf As Integer
Dim MyDiameter As Integer
    
    AppDiameter = CInt(TextAllPocketsDiameter.Text)
    CurrentTest = "3_TEST_ALL_POCKETS.JBI"
    shelf = CInt(TextShelfNumber2.Text)
    MyDiameter = CInt(TextAllPocketsDiameter.Text)
    
    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(12, 12, 12, "MdiMain", "BtnTestAllPockets_Click()", KeyPosition)
        Exit Sub
    End If
    
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf).ShelfToolType <> AppToolType Then
        ret = MsgBox("the shelf is incorrect ! " & vbCrLf & _
            "the shelf type is not the same as the gripper type" _
                , vbCritical, "BtnTestAllPockets_Click()")
        Exit Sub
    End If
    
    If AppToolType <> HSK Then
        If OffsetSent = False Then
            ret = MsgBox("Please send Offset to robot", _
                            vbInformation, _
                                "BtnTestPocketToPocket_Click()")
        End If
    End If
    
    ModHandShake.AllPocketsShelf = shelf
    ModHandShake.AllPocketsColumn = 1
    ModHandShake.AllPocketsDiameter = MyDiameter
    
    ''chkEndless.Value = vbChecked ''  set endless loop.
     
    If AppToolType = Drill Then
        Call ModHandShake.BuildArrayPosition(shelf, 1, MyDiameter, Drill)    '''        build the position
        Call MotoComToolBox.WritePosition(16, ArrayPosition) '''        send the position
        Call ModHandShake.WriteSensorLocation(shelf, 1, MyDiameter, 13)
        
        Call ModHandShake.BuildArrayPosition(shelf, 2, MyDiameter, Drill) '''        build the position
        Call MotoComToolBox.WritePosition(17, ArrayPosition) '''        send the position
        Call ModHandShake.WriteSensorLocation(shelf, 2, MyDiameter, 20)
        
        Call WriteOneCommByte(52, tooltype.Drill)
        
    ElseIf AppToolType = HSK Then
        Call ModHandShake.BuildArrayPosition(shelf, 1, MyDiameter, HSK) '''    build the position
        Call MotoComToolBox.WritePosition(16, ArrayPosition) '''               send the position
        
        Call ModHandShake.BuildArrayPosition(shelf, 2, MyDiameter, HSK) '''    build the position
        Call MotoComToolBox.WritePosition(17, ArrayPosition) '''               send the position
        
        Call WriteOneCommByte(52, tooltype.HSK)
        
    ElseIf AppToolType = Round Then
        Call ModHandShake.BuildArrayPosition(shelf, 1, MyDiameter, Round) '''    build the position
        Call MotoComToolBox.WritePosition(16, ArrayPosition) '''                 send the position
        Call ModHandShake.WriteSensorLocation(shelf, 1, MyDiameter, 13)
        
        Call ModHandShake.BuildArrayPosition(shelf, 2, MyDiameter, Round) '''        build the position
        Call MotoComToolBox.WritePosition(17, ArrayPosition) '''        send the position
        Call ModHandShake.WriteSensorLocation(shelf, 1, MyDiameter, 20)
        
        
        Call WriteOneCommByte(52, tooltype.Round)
        
    End If

    If AppToolType <> HSK Then
        Call CmdSetOffset_Click
    End If
    
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(12, 11, 26, "mdiMain", "BtnTestAllPockets_Click()", EmergencyStop)
        Exit Sub
    End If
    
        
    bb = ModHandShake.WriteDrillCode(AppDiameter)
    If bb = False Then
        Call FrmDialog.ShowDialogForm(12, 60, 60, "mdiMain", "CmdSysOperation_Click()", GeneralError)
        Exit Sub
    End If
            
    bb = ModHandShake.SendToolSensorState()
    If bb = False Then
        Call FrmDialog.ShowDialogForm(12, 61, 61, "mdiMain", "CmdSysOperation_Click()", GeneralError)
        Exit Sub
    End If
   
    bb = StartJob(CurrentTest)
    If bb = False Then
        Call FrmDialog.ShowDialogForm(12, 32, 32, "mdimain", " BtnTestAllPockets_Click()", InputIncomplete)
        Exit Sub
    End If
     
     ModLogFile.LogAddLine (" Start Test : All Pockets.")
     Call ModGUI.ListHMIUpdate("Start Test : All Pockets.")
     
End Sub


Private Sub BtnTestPocketToPocket_Click()

Dim shelf_from As Integer
Dim column_from As Integer
Dim pocket_from As Integer

Dim shelf_to As Integer
Dim column_to As Integer
Dim pocket_to As Integer
Dim LineNumber As Integer

Dim ret As Integer
Dim bb As Boolean

    AppDiameter = CInt(TextToolsDiameter2.Text)
    CurrentTest = "3_TEST_POCKET_TO_POCKET.JBI"
    ''Call WriteOneCommByte(52, tooltype.HSK)
    
    shelf_from = CInt(Left(txtTestPocket(0).Text, 1))
    column_from = CInt(Right(txtTestPocket(0).Text, 2))
    pocket_from = CInt(TextToolsDiameter2.Text)
    
    shelf_to = CInt(Left(txtTestPocket(1).Text, 1))
    column_to = CInt(Right(txtTestPocket(1).Text, 2))
    pocket_to = CInt(TextToolsDiameter2.Text)
        
    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(11, 12, 12, "MdiMain", "BtnTestPocketToPocket_Click()", KeyPosition)
        Exit Sub
    End If
 
    ''check if the shelf tooltype is the same as the gripper tool type.
    If AppShelvs(shelf_from).ShelfToolType <> AppToolType Then
        ret = MsgBox("the shelf is incorrect.the shelf type is not the same as the gripper type", _
            vbCritical, "BtnTestPocketToPocket_Click()")
        Exit Sub
    End If
    
    If AppToolType <> HSK Then
        If OffsetSent = False Then
            ret = MsgBox("Please send Offset to robot", _
                            vbInformation, _
                                "BtnTestPocketToPocket_Click()")
        End If
    End If
    
    ''If AppToolType = drill Then
    If AppShelvs(shelf_from).ShelfToolType = Drill Then
    
        Call BuildArrayPosition(shelf_from, column_from, pocket_from, Drill) '''build the position
        Call MotoComToolBox.WritePosition(16, ArrayPosition)                  '''send the position
        Call ModHandShake.WriteSensorLocation(shelf_from, column_from, pocket_from, 13)
        
        Call BuildArrayPosition(shelf_to, column_to, pocket_to, Drill)
        Call MotoComToolBox.WritePosition(17, ArrayPosition)
        Call ModHandShake.WriteSensorLocation(shelf_to, column_to, pocket_to, 20)
        
        Call WriteOneCommByte(52, tooltype.Drill)
    
    ''ElseIf AppToolType = hsk Then
    ElseIf AppShelvs(shelf_from).ShelfToolType = HSK Then
    
        Call BuildArrayPosition(shelf_from, column_from, pocket_from, HSK)
        Call MotoComToolBox.WritePosition(16, ArrayPosition)
        Call ModHandShake.WriteSensorLocation(shelf_from, column_from, pocket_from, 13)
        
        Call BuildArrayPosition(shelf_to, column_to, pocket_to, HSK)
        Call MotoComToolBox.WritePosition(17, ArrayPosition)
        Call ModHandShake.WriteSensorLocation(shelf_to, column_to, pocket_to, 20)
        
        Call WriteOneCommByte(52, tooltype.HSK)
        
    '''ElseIf AppToolType = Round Then
    ElseIf AppShelvs(shelf_from).ShelfToolType = Round Then

        Call BuildArrayPosition(shelf_from, column_from, pocket_from, Round) ''              builed the position
        ArrayPosition(4) = RoundLocations(shelf_from, column_from).diameter(pocket_from).z - PocketDepth + ChuckDepth - ChuckStopper + ShelfSafty
        Call MotoComToolBox.WritePosition(16, ArrayPosition)
        Call ModHandShake.WriteSensorLocation(shelf_from, column_from, pocket_from, 13)
        
        Call BuildArrayPosition(shelf_to, column_to, pocket_to, Round) ''                     builed the position
        ArrayPosition(4) = RoundLocations(shelf_to, column_to).diameter(pocket_to).z - PocketDepth + ChuckDepth - ChuckStopper + ShelfSafty
        Call MotoComToolBox.WritePosition(17, ArrayPosition)
        Call ModHandShake.WriteSensorLocation(shelf_to, column_to, pocket_to, 20)
        
        Call WriteOneCommByte(52, tooltype.Round)
            
    End If
    
    If AppToolType <> HSK Then
        Call CmdSetOffset_Click
    End If
    
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(19, 11, 26, "mdiMain", " BtnTestPocketToPocket_Click()", GeneralError)
        Exit Sub
    End If
    
    bb = ModHandShake.WriteDrillCode(AppDiameter)
    If bb = False Then
        Call FrmDialog.ShowDialogForm(60, 60, 60, "mdiMain", " BtnTestPocketToPocket_Click()", GeneralError)
        Exit Sub
    End If
            
    bb = ModHandShake.SendToolSensorState()
    If bb = False Then
        Call FrmDialog.ShowDialogForm(61, 61, 61, "mdiMain", " BtnTestPocketToPocket_Click()", GeneralError)
        Exit Sub
    End If
        

    RunTestEndless = False
    
    bb = StartJob(CurrentTest)
    If bb = False Then
        Call FrmDialog.ShowDialogForm(19, 32, 32, "mdimain", " BtnTestPocketToPocket_Click()", InputIncomplete)
        Exit Sub
    End If

    ModLogFile.LogAddLine (" Start Test : Pocket To Pocket.")
    Call ModGUI.ListHMIUpdate(" Start Test : Pocket To Pocket.")
    
End Sub

Public Sub BtnViewPocketsLocations_Click()
''the function read the pockets location from the memory.
''the function display the location (coordinates) in table.
Debug.Print "BtnViewPocketsLocations_Click()"
''
''
Dim shelf As Integer
Dim column As Integer
Dim orient As Integer
Dim Xpos As Double
Dim Ypos As Double
Dim Zpos As Double
Dim RxPos As Double
Dim RyPos As Double
Dim RzPos As Double
Dim line As Integer
Dim ret As Integer

    line = 1
    
    shelf = CInt(TextShelfNumber.Text)
    On Error GoTo error
     
    ''If AppToolType = hsk Then
   If AppShelvs(shelf).ShelfToolType = HSK Then
        ModCsvFile.ReadPocketsLocations ("HSKLocations") ''read the pockets location from file.
   
        If AppShelvs(shelf).ShelfToolType <> AppToolType Then
            TablePocketsLocations.Height = 650
            TablePocketsLocations.Enabled = False
        ElseIf AppShelvs(shelf).ShelfToolType = AppToolType Then
            TablePocketsLocations.Height = 3840
            TablePocketsLocations.Enabled = True
        End If
            
        TablePocketsLocations.Rows = 11
        TablePocketsLocations.Width = 8080
        
        shelf = CInt(TextShelfNumber.Text)
        For line = 1 To 10
        
            ''header:the most left column.
            TablePocketsLocations.TextMatrix(line, 0) = shelf * 100 + line

            'X position
            Xpos = HSKLocations(shelf, line).X
            TablePocketsLocations.TextMatrix(line, 1) = Format(Xpos, "000.000")

            'Y position
            Ypos = HSKLocations(shelf, line).Y
            TablePocketsLocations.TextMatrix(line, 2) = Format(Ypos, "000.000")

            'Z position
            Zpos = HSKLocations(shelf, line).z
            TablePocketsLocations.TextMatrix(line, 3) = Format(Zpos, "000.000")
             
            'Rx position
            RxPos = HSKLocations(shelf, line).Rx
            TablePocketsLocations.TextMatrix(line, 4) = Format(RxPos, "000.000")

            'Ry position
            RyPos = HSKLocations(shelf, line).Ry
            TablePocketsLocations.TextMatrix(line, 5) = Format(RyPos, "000.000")

            'Rz position
            RzPos = HSKLocations(shelf, line).Rz
            TablePocketsLocations.TextMatrix(line, 6) = Format(RzPos, "000.000")
            

        Next
    End If
    
    '''If AppToolType = drill Then
    If AppShelvs(shelf).ShelfToolType = Drill Then
    
        If AppShelvs(shelf).ShelfToolType <> AppToolType Then
            TablePocketsLocations.Height = 650
            TablePocketsLocations.Enabled = False
        ElseIf AppShelvs(shelf).ShelfToolType = AppToolType Then
            TablePocketsLocations.Height = 3840
            TablePocketsLocations.Enabled = True
        End If
        
        TablePocketsLocations.Rows = 8
        TablePocketsLocations.Width = 8070
        shelf = CInt(TextShelfNumber.Text)
        column = CInt(Right(TxtViewPocket.Text, 2))
        For orient = 1 To 7
            
            ''header:the most left column.
            TablePocketsLocations.TextMatrix(orient, 0) = orient
            
            'X position
            Xpos = DrillLocations(shelf, column).diameter(orient).X
            TablePocketsLocations.TextMatrix(orient, 1) = Format(Xpos, "000.000")
            
            'Y position
            Ypos = DrillLocations(shelf, column).diameter(orient).Y
            TablePocketsLocations.TextMatrix(orient, 2) = Format(Ypos, "000.000")
            
            'Z position
            Zpos = DrillLocations(shelf, column).diameter(orient).z
            TablePocketsLocations.TextMatrix(orient, 3) = Format(Zpos, "000.000")
            
            ''rx
            Xpos = DrillLocations(shelf, column).diameter(orient).Rx
            TablePocketsLocations.TextMatrix(orient, 4) = Format(Xpos, "000.000")
            
            'rY position
            Ypos = DrillLocations(shelf, column).diameter(orient).Ry
            TablePocketsLocations.TextMatrix(orient, 5) = Format(Ypos, "000.000")
            
            'rZ position
            Zpos = DrillLocations(shelf, column).diameter(orient).Rz
            TablePocketsLocations.TextMatrix(orient, 6) = Format(Zpos, "000.000")
        
            If AppShelvs(shelf).ShelfToolType <> AppToolType Then
                TablePocketsLocations.Height = 650
                TablePocketsLocations.Enabled = False
            End If
            
        Next
        End If
        
   ''' If AppToolType = Round Then
   If AppShelvs(shelf).ShelfToolType = Round Then
   
        If AppShelvs(shelf).ShelfToolType <> AppToolType Then
            TablePocketsLocations.Height = 650
            TablePocketsLocations.Enabled = False
        ElseIf AppShelvs(shelf).ShelfToolType = AppToolType Then
            TablePocketsLocations.Height = 3840
            TablePocketsLocations.Enabled = True
        End If
        
        TablePocketsLocations.Rows = 9
        TablePocketsLocations.Width = 8070
        shelf = CInt(TextShelfNumber.Text)
        column = CInt(Right(TxtViewPocket.Text, 2))
        For orient = 1 To 8
            
            ''header:the most left column.
            TablePocketsLocations.TextMatrix(orient, 0) = orient
            
            'X position
            Xpos = RoundLocations(shelf, column).diameter(orient).X
            TablePocketsLocations.TextMatrix(orient, 1) = Format(Xpos, "000.000")
            
            'Y position
            Ypos = RoundLocations(shelf, column).diameter(orient).Y
            TablePocketsLocations.TextMatrix(orient, 2) = Format(Ypos, "000.000")
            
            'Z position
            Zpos = RoundLocations(shelf, column).diameter(orient).z
            TablePocketsLocations.TextMatrix(orient, 3) = Format(Zpos, "000.000")
            
            'Rx position
            RxPos = RoundLocations(shelf, column).diameter(orient).Rx
            TablePocketsLocations.TextMatrix(orient, 4) = Format(RxPos, "000.000")

            'Ry position
            RyPos = RoundLocations(shelf, column).diameter(orient).Ry
            TablePocketsLocations.TextMatrix(orient, 5) = Format(RyPos, "000.000")

            'Rz position
            RzPos = RoundLocations(shelf, column).diameter(orient).Rz
            TablePocketsLocations.TextMatrix(orient, 6) = Format(RzPos, "000.000")
            
            If AppShelvs(shelf).ShelfToolType <> AppToolType Then
                TablePocketsLocations.Height = 650
                TablePocketsLocations.Enabled = False
            End If
        
        Next orient

    End If
    
    Exit Sub
    
error:
       ret = MsgBox("error while display pocket location" & vbCrLf _
       & "the error is : " & Err.Description _
       , vbExclamation _
       , " BtnViewPocketsLocations_Click()")

            
End Sub









Private Sub CmdClearHMIList_Click()

    ListHMIInfo.Clear
    
End Sub

Private Sub CmdClearRobotList_Click()

    ListRobotInfo.Clear
    
End Sub

Private Sub cmdGoExchangePos_Click()

Dim ret As Integer

    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(12, 12, 12, "MdiMain", "BtnTestAllPockets_Click()", 1)
        Exit Sub
    End If
    
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(19, 11, 26, "mdiMain", "BtnStartLoadUnload_Click()", 1)
        Exit Sub
    End If
    
    RunTestEndless = False
    
    Call StartJob("4_EXCHANGE_GRIPPER.JBI")
    
End Sub


Private Sub cmdGoParkingPos_Click()
Dim ret As Integer

    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(12, 12, 12, "MdiMain", "BtnTestAllPockets_Click()", 1)
        Exit Sub
    End If

        

    RunTestEndless = False
    
    Call SetServo(True)
    Call StartJob("4_PARKING.JBI")
    
End Sub

Private Sub cmdGoZeroPos_Click()

Dim ret As Integer

    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(12, 12, 12, "MdiMain", "BtnTestAllPockets_Click()", 1)
        Exit Sub
    End If
    
    RunTestEndless = False
    
    Call SetServo(True)
    Call StartJob("5_CURRENT_RETRACT.JBI")
    
End Sub

Private Sub cmdLoadTool_Click()
''1.the function run RobotJob that take tool from the kiosk to the right pocket.
''2.the function being called from the GUI.
''3.the function take no parameters.
''4.the f return no parameters.


Dim ret As Integer
Dim iWPiece As Integer
Dim iPocketNumber As Integer
Dim shelf_to As Integer
Dim column_to As Integer
Dim pocket_to As Integer
Dim MyPocket As String
Dim bb As Boolean

On Error GoTo LabelLoad
    
    

    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(56, 12, 12, "MdiMain", "cmdLoadTool_Click()", KeyPosition)
        Exit Sub
    End If
    
    If GripperStatus = 0 Then
        ret = MsgBox("The Gripper is Close." & vbCrLf _
        & "Pleae Open Gripper.", vbExclamation, "gripper state")
        Exit Sub
    End If
    
    
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(56, 11, 26, "mdiMain", "cmdLoadTool_Click()", EmergencyStop)
        Exit Sub
    End If
        
    ' check correct input from user:'fields should be integers
    If IsNumeric(txtLoadUnloadPocket.Text) = False Or txtLoadUnloadPocket.Text = "" Then
        ret = MsgBox("Wrong Pocket number ", vbInformation, "Error in Loading Tool To Pocket.")
        Exit Sub 'GoTo labelLoad
    End If
    
    ''check if the shelf tooltype is the same as the application tool type.
    shelf_to = CInt(fMainForm.txtLoadUnloadShelf.Text)
    If AppShelvs(shelf_to).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
    
    ' check correct input from user:'fields should be integers
    If IsNumeric(txtWorkPiece(1).Text) = False Or txtWorkPiece(1).Text = "" Then
        ret = MsgBox("Wrong Workpiece number ", vbInformation, "Error in Loading Tool To Pocket.")
        Exit Sub 'GoTo labelLoad
    End If
''
     'check if Work Piece is (1-100)
    If CInt(txtWorkPiece(1).Text) > 100 Or CInt(txtWorkPiece(1).Text) < 1 Then
        ret = MsgBox("Wrong Workpiece number ", vbInformation, "Error in Loading Tool To Pocket.")
        Exit Sub 'GoTo labelLoad
    End If

    If AppToolType = HSK Then ''check if pocket number is between the bounds.

        If (CInt(txtLoadUnloadPocket) >= 101 And CInt(txtLoadUnloadPocket) <= 110) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 201 And CInt((txtLoadUnloadPocket.Text)) <= 210) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 301 And CInt((txtLoadUnloadPocket.Text)) <= 310) _
            Then
        Else
            ret = MsgBox("Wrong Pocket number", vbInformation, "Error in Loading Tool To Pocket.")
            Exit Sub
        End If
        Call WriteOneCommByte(52, tooltype.HSK)

     ElseIf AppToolType = Drill Then  ''check if pocket number is between the bounds.
     
        If (CInt((txtLoadUnloadPocket.Text)) >= 101 And CInt((txtLoadUnloadPocket.Text)) <= 112) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 201 And CInt((txtLoadUnloadPocket.Text)) <= 212) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 301 And CInt((txtLoadUnloadPocket.Text)) <= 312) _
           Then
        Else
            ret = MsgBox("Wrong Pocket number" & vbCrLf & _
                "Error in Loading Tool To Pocket.", _
                vbExclamation, " cmdLoadTool_Click()")
            Exit Sub
        End If
        Call WriteOneCommByte(52, tooltype.Drill)
        
    ElseIf AppToolType = Round Then  ''check if pocket number is between the bounds.
     
        If (CInt((txtLoadUnloadPocket.Text)) >= 101 And CInt((txtLoadUnloadPocket.Text)) <= 112) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 201 And CInt((txtLoadUnloadPocket.Text)) <= 212) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 301 And CInt((txtLoadUnloadPocket.Text)) <= 312) _
           Then
        Else
            ret = MsgBox("Wrong Pocket number", vbInformation, "Error in Loading Tool To Pocket.")
            Exit Sub
        End If
        Call WriteOneCommByte(52, tooltype.Round)
        
     End If
    
    ret = SetServo(True) ''make sure the servo is ON.
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(56, 11, 26, "mdiMain", "cmdLoadTool_Click()", 1)
        Exit Sub
    End If
    
    
    bb = ModHandShake.WriteDrillCode(AppDiameter)
    If bb = False Then
        Call FrmDialog.ShowDialogForm(56, 60, 60, "mdiMain", "cmdLoadTool_Click()", GeneralError)
        Exit Sub
    End If
            
    bb = ModHandShake.SendToolSensorState()
    If bb = False Then
        Call FrmDialog.ShowDialogForm(56, 61, 61, "mdiMain", "cmdLoadTool_Click()", GeneralError)
        Exit Sub
    End If
   
    If Jobs.AutoStart(0) = JobState.RUN Then
        SetOneCommByte (24) ''rise request for run the load cycle
    ElseIf Jobs.AutoStart(0) <> JobState.RUN Then
        bb = StartJob("3_KIOSK_TO_POCKET.JBI") ''start the load cycle.
        If bb = False Then
            Call FrmDialog.ShowDialogForm(56, 32, 32, "MdiMain", "cmdLoadTool_Click()", InputIncomplete)
            Exit Sub
        End If
    End If
    
    

   Exit Sub
    
LabelLoad:
   ret = MsgBox("The System was unable to Load Tool To  Pocket," & vbCrLf _
        & "1.Check the Inputs.should not be empty." & vbCrLf _
        & "2.Check the Inputs.should be numbers only." & vbCrLf _
        & "3.WorkPiece Should be between (1-100)." & vbCrLf _
        & "4.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.cmdLoadTool_Click]", vbInformation, "Error in Loading Tool To Pocket.")
        

''b24 Request Load autoStart job
''b25 Done  Load autoStart job
''b26 Request UnLoad autoStart job
''b27 Done  UnLoad autoStart job
End Sub

Private Sub cmdMove_Click(Index As Integer)
''1.the function change the priority of one WP in the AllWP()Array.
''2.the f change the GUI too by calling the function:BtnRefreshOrder_Click
''3.Legend:
        'index 6 -> up priority
        'index 7 -> down priority (increse number)
'

Dim line As Integer
Dim TempLine As Integer
Dim ret As Integer
Dim TempWP As WorkPiece
Dim TempBool As Boolean

On Error GoTo error
    
    'chack if the inputs is a number
    TempBool = IsNumeric((TextLineNumber.Text))
    If TempBool = False Then
        Call FrmDialog.ShowDialogForm(23, 26, 17, "MdiMain", "cmdMove_Click()", GeneralError)
        TextLineNumber.Text = ""
        Exit Sub
    End If

   
    line = CInt(TextLineNumber.Text)
    
    If line = 1 Then
        ModHandShake.SetOneCommByte (28)
    End If
    
    If Index = 7 Then
        ''save the current workpiece into the TempWP.
        TempWP.LineNumber = AllWP(line).LineNumber
        TempWP.NCProgram = AllWP(line).NCProgram
        TempWP.ToolAmount = AllWP(line).ToolAmount
        TempWP.ToolAmountLeft = AllWP(line).ToolAmountLeft
        TempWP.ToolDiameter = AllWP(line).ToolDiameter
        TempWP.WPNumber = AllWP(line).WPNumber
        TempWP.WPToolType = AllWP(line).WPToolType
        TempWP.WPStatus = AllWP(line).WPStatus
    
    ''push the upper WP 1 step down.
        AllWP(line).LineNumber = AllWP(line + 1).LineNumber
        AllWP(line).NCProgram = AllWP(line + 1).NCProgram
        AllWP(line).ToolAmount = AllWP(line + 1).ToolAmount
        AllWP(line).ToolAmountLeft = AllWP(line + 1).ToolAmountLeft
        AllWP(line).ToolDiameter = AllWP(line + 1).ToolDiameter
        AllWP(line).WPNumber = AllWP(line + 1).WPNumber
        AllWP(line).WPToolType = AllWP(line + 1).WPToolType
        AllWP(line).WPStatus = AllWP(line + 1).WPStatus
        
    ''push the temp to the upper WP.
        AllWP(line + 1).LineNumber = TempWP.LineNumber
        AllWP(line + 1).NCProgram = TempWP.NCProgram
        AllWP(line + 1).ToolAmount = TempWP.ToolAmount
        AllWP(line + 1).ToolAmountLeft = TempWP.ToolAmountLeft
        AllWP(line + 1).ToolDiameter = TempWP.ToolDiameter
        AllWP(line + 1).WPNumber = TempWP.WPNumber
        AllWP(line + 1).WPToolType = TempWP.WPToolType
        AllWP(line + 1).WPStatus = TempWP.WPStatus
        
    End If
    
    If Index = 6 Then
    
        ''save the current workpiece into the TempWP.
        TempWP.LineNumber = AllWP(line).LineNumber
        TempWP.NCProgram = AllWP(line).NCProgram
        TempWP.ToolAmount = AllWP(line).ToolAmount
        TempWP.ToolAmountLeft = AllWP(line).ToolAmountLeft
        TempWP.ToolDiameter = AllWP(line).ToolDiameter
        TempWP.WPNumber = AllWP(line).WPNumber
        TempWP.WPToolType = AllWP(line).WPToolType
        TempWP.WPStatus = AllWP(line).WPStatus
    
    ''push the lower WP 1 step up.
        AllWP(line).LineNumber = AllWP(line - 1).LineNumber
        AllWP(line).NCProgram = AllWP(line - 1).NCProgram
        AllWP(line).ToolAmount = AllWP(line - 1).ToolAmount
        AllWP(line).ToolAmountLeft = AllWP(line - 1).ToolAmountLeft
        AllWP(line).ToolDiameter = AllWP(line - 1).ToolDiameter
        AllWP(line).WPNumber = AllWP(line - 1).WPNumber
        AllWP(line).WPToolType = AllWP(line - 1).WPToolType
        AllWP(line).WPStatus = AllWP(line - 1).WPStatus
        
        
        ''push the temp to the lower WP.
        AllWP(line - 1).LineNumber = TempWP.LineNumber
        AllWP(line - 1).NCProgram = TempWP.NCProgram
        AllWP(line - 1).ToolAmount = TempWP.ToolAmount
        AllWP(line - 1).ToolAmountLeft = TempWP.ToolAmountLeft
        AllWP(line - 1).ToolDiameter = TempWP.ToolDiameter
        AllWP(line - 1).WPNumber = TempWP.WPNumber
        AllWP(line - 1).WPToolType = TempWP.WPToolType
        AllWP(line - 1).WPStatus = TempWP.WPStatus
        
    End If
        
    Call ModCsvFile.SaveAllWP ''         save the data array into file
    Call SaveAutomation ''               save the array into file
    Call BtnDisplayWPTable_Click   '     refresh the work piece Table
    Call BtnDisplayPocketStatus_Click '' refresh the pocket status table
        
    Exit Sub
        
    
error:
  ret = MsgBox("The System was unable to change ptiority of a WorkPiece." & vbCrLf & _
  "Try to fix one of the options below :" & vbCrLf & _
  "1.Inputs can not be empty." & vbCrLf & _
  "2.Communication problem with the controller." & vbCrLf & _
  "(mdiMain.cmdMove_Click)", vbExclamation, "Change priority")

End Sub

Private Sub cmdOutOff_Click(Index As Integer)
'this function set the output Off
'the function get the index=Button Number

Dim ret As Integer
Dim bb As Boolean
Dim MyIndex As Integer
    
    MyIndex = Index + 20
    On Error GoTo out
    ArrayInteger(0) = CDbl(MyIndex)
    bb = WriteInteger(60, ArrayInteger())
    If bb = False Then
        GoTo out
    End If
    
    ''fMainForm.txtCycleMessage(1).Caption = "send :output " & CStr(Index) & " off"
    Exit Sub
    
out:
    Call FrmDialog.ShowDialogForm(57, 57, 57, "cmdOutOff_Click()", "mdimain", GeneralError)
  
End Sub

Private Sub cmdOutOn_Click(Index As Integer)

Dim ret As Integer
Dim bb As Boolean

    On Error GoTo out
   
    ArrayInteger(0) = CDbl(Index)
    bb = WriteInteger(60, ArrayInteger())
    If bb = False Then
        GoTo out
    End If
    ''fMainForm.txtCycleMessage(1).Caption = "send :output " & CStr(Index) & " off"
    
Exit Sub
out:
    Call FrmDialog.ShowDialogForm(57, 57, 57, "mdimain", "cmdOutOn_Click()", GeneralError)
  
End Sub


Private Sub CmdResetAllTests_Click()
''
''1.this function reset 6 jobs in the controller memory.

Dim ret As Integer

    If AppKeyState <> remote Then
        ret = MsgBox("can not reset tests cycles." & vbCrLf _
        & "the key is not in remote mode." _
        , vbExclamation, _
        "CmdResetAllTests_Click()")
        Exit Sub
    End If

    ProgressBarResetTests.Visible = True
    ProgressBarResetTests.Value = 1
    Call ModHandShake.ResetRobotJob("3_TEST_KIOSK_TO_POCKET.JBI")
    
    ProgressBarResetTests.Value = 2
    Call ModHandShake.ResetRobotJob("3_TEST_POCKET_TO_KIOSK.JBI")
    
    ProgressBarResetTests.Value = 3
    Call ModHandShake.ResetRobotJob("3_TEST_POCKET_TO_CHUCK.JBI")
    
    ProgressBarResetTests.Value = 4
    Call ModHandShake.ResetRobotJob("3_TEST_CHUCK_TO_POCKET.JBI")
    
    ProgressBarResetTests.Value = 5
    Call ModHandShake.ResetRobotJob("3_TEST_POCKET_TO_POCKET.JBI")
    
    ProgressBarResetTests.Value = 6
    Call ModHandShake.ResetRobotJob("3_TEST_ALL_POCKETS.JBI")
    
    ProgressBarResetTests.Value = 7
    CurrentTest = ""
    
    CheckLoopAllPockets.Value = vbUnchecked
    CheckLoopP2P.Value = vbUnchecked
    
    ProgressBarResetTests.Visible = False

End Sub

Private Sub CmdResetStatusTable_Click()

    Call ModAutomationStatus.AutomationStatusReset
    Call BtnDisplayPocketStatus_Click
    ModLogFile.LogAddLine (" Clear all pocket's status ")
    Call ModGUI.ListHMIUpdate(" Clear all pocket's status ")
    
End Sub

Private Sub CmdResetWPiece_Click()

Dim TempArray(8) As Double

    TempArray(0) = 0
    Call WriteInteger(11, TempArray())  '' tool counter
    Call WriteByte(104, TempArray()) ''    WPiece Counter
    Call WriteByte(23, TempArray()) ''     no more workpiece
    Call WriteByte(105, TempArray()) '''   the first part of the first WP.
    
    ChuckUnloadFirstTime = True
    Call ModHandShake.ResetAllRobotRequest
    
End Sub

Public Sub CmdRestoreDefult_Click()

    ''above pocket
    If textToolOffset(0).Enabled = True Then
        ModIni.ReadAbovePocket
        textToolOffset(0).Text = CStr(AbovePocket)
    End If
    
    ''above chuck
    If textToolOffset(1).Enabled = True Then
        ModIni.ReadAboveChuck
        textToolOffset(1).Text = CStr(AboveChuck)
    End If
    
    ''chuck stopper
    If textToolOffset(2).Enabled = True Then
        ModIni.ReadChuckStopper
        textToolOffset(2).Text = CStr(ChuckStopper)
    End If
    
    ''chuck depth
    If textToolOffset(3).Enabled = True Then
        ModIni.ReadChuckDepth
        textToolOffset(3).Text = CStr(ChuckDepth)
    End If
    
    ''pocket stoper offset
    If textToolOffset(4).Enabled = True Then
        ModIni.ReadPocketStopper
        textToolOffset(4).Text = CStr(PocketStopper)
    End If
    
    ''kiosk offset
    If textToolOffset(5).Enabled = True Then
        ModIni.ReadKioskStopper
        textToolOffset(5).Text = CStr(KioskStopper)
    End If

End Sub

Private Sub CmdSetOffset_Click()
''
''1.this function send data to robot to D11 and to D12.
''2.D11 : how high to grip the tool.taken fron the shelf surface.
''2.D12 : how high to grip the tool.taken fron the chuck surface.
Dim ret As Integer
On Error GoTo wronginput

    ChuckDepth = 274
    ChuckStopper = 75
    ShelfSafty = 10
    PocketDepth = 92

    If ((AppToolType = HSK) Or (AppToolType = other)) Then
        Exit Sub
    End If
    
    ''check the key state
    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(14, 12, 12, "MdiMain", "CmdSetOffset_Click", KeyPosition)
        Exit Sub
    End If
    
    If textToolOffset(0).Enabled = True Then
        If textToolOffset(0).Text = "" Then
            GoTo wronginput
        End If
        If IsNumeric(textToolOffset(0).Text) = False Then
            GoTo wronginput
        End If
        ArrayInteger(0) = CDbl(fMainForm.textToolOffset(0).Text) * 1000
        Call WriteDouble(11, ArrayInteger())      ''height above pocket
        AbovePocket = CDbl(fMainForm.textToolOffset(0).Text)
        ModIni.WriteAbovePocket
    End If
    
    
    If textToolOffset(1).Enabled = True Then
        If textToolOffset(1).Text = "" Then
            GoTo wronginput
        End If
        If IsNumeric(textToolOffset(1).Text) = False Then
            GoTo wronginput
        End If
        ArrayInteger(0) = CDbl(fMainForm.textToolOffset(1).Text) * 1000
        Call WriteDouble(12, ArrayInteger())       ''height above chuck
        AboveChuck = CDbl(fMainForm.textToolOffset(1).Text)
        ModIni.WriteAboveChuck
    End If
    
    
    If textToolOffset(2).Enabled = True Then
        If textToolOffset(2).Text = "" Then
            GoTo wronginput
        End If
        If IsNumeric(textToolOffset(2).Text) = False Then
            GoTo wronginput
        End If
        ArrayInteger(0) = CDbl(fMainForm.textToolOffset(2).Text) * 1000
        Call WriteDouble(68, ArrayInteger())      ''      Chuck Stopper
        ChuckStopper = CDbl(fMainForm.textToolOffset(2).Text)
        ModIni.WriteChuckStopper
    End If
    
    If textToolOffset(3).Enabled = True Then
        If textToolOffset(3).Text = "" Then
            GoTo wronginput
        End If
        If IsNumeric(textToolOffset(3).Text) = False Then
            GoTo wronginput
        End If
        ArrayInteger(0) = CDbl(fMainForm.textToolOffset(3).Text) * 1000
        Call WriteDouble(69, ArrayInteger())      ''        Chuck Depth
        ChuckDepth = CDbl(textToolOffset(3))
        ModIni.WriteChuckDepth
    End If
    
    If textToolOffset(4).Enabled = True Then
        If textToolOffset(4).Text = "" Then
            GoTo wronginput
        End If
        If IsNumeric(textToolOffset(4).Text) = False Then
            GoTo wronginput
        End If
        ArrayInteger(0) = CDbl(fMainForm.textToolOffset(4).Text) * 1000
        Call WriteDouble(67, ArrayInteger())      ''     pocket stopper
        PocketStopper = CDbl(fMainForm.textToolOffset(4).Text)
        ModIni.WritePocketStopper
        
    End If
    
    If textToolOffset(5).Enabled = True Then
        If textToolOffset(5).Text = "" Then
            GoTo wronginput
        End If
        If IsNumeric(textToolOffset(5).Text) = False Then
            GoTo wronginput
        End If
        ArrayInteger(0) = CDbl(fMainForm.textToolOffset(5).Text) * 1000
        Call WriteDouble(66, ArrayInteger())      ''     Kiosk   stopper
        KioskStopper = CDbl(fMainForm.textToolOffset(5).Text)
        ModIni.WriteKioskStopper
        
    End If
    
    ModLogFile.LogAddLine (" Send offsets to robot.")
    Call ModGUI.ListHMIUpdate(" Send offsets to robot.")
    
    OffsetSent = True
    ' elisha 15-01-2014
'''    ret = MsgBox(vbCrLf & "     Offsets have been Sent To Robot       " & vbCrLf _
'''        & "     Offsets have been saved on file" & vbCrLf _
'''        & "" & vbCrLf _
'''        , vbInformation, _
'''        "Send offset to Robot.")
    Exit Sub
    
    
wronginput:
    Call FrmDialog.ShowDialogForm(62, 62, 62, "MdiMain", "CmdSetOffset_Click()", InputIncomplete)
    textToolOffset(0).Text = ""
    textToolOffset(1).Text = ""
    OffsetSent = False
    
End Sub

Private Sub CmdSimChuckToPocket_Click()

    AppDiameter = AllWP(AppWPIndex).ToolDiameter
    HandShake.RequestPlaceFromChuck(0) = 1
    ModHandShake.WriteDataToRobot
    HandShake.RequestPlaceFromChuck(0) = 0
    
End Sub

Private Sub CmdSimKioskToPocket_Click()

    AppDiameter = AllWP(AppWPIndex).ToolDiameter
    HandShake.RequestPlaceFromKiosk(0) = 1
    ModHandShake.WriteDataToRobot
    HandShake.RequestPlaceFromKiosk(0) = 0

End Sub


Private Sub CmdSimPocketToChuck_Click()

    AppDiameter = AllWP(AppWPIndex).ToolDiameter
    HandShake.RequestTakeToChuck(0) = 1
    ModHandShake.WriteDataToRobot
    HandShake.RequestTakeToChuck(0) = 0
    
End Sub

Private Sub CmdSimPocketToKiosk_Click()

    AppDiameter = AllWP(AppWPIndex).ToolDiameter
    HandShake.RequestTakeToKiosk(0) = 1
    ModHandShake.WriteDataToRobot
    HandShake.RequestTakeToKiosk(0) = 0
    
End Sub



Private Sub CmdStartSawTest_Click()

Dim ret As Integer
Dim bb As Boolean


        
    If ((CurrentTest <> "3_TEST_STATION_TO_SPINDLE.JBI") _
        And (CurrentTest <> "3_TEST_SPINDLE_TO_STATION.JBI")) Then
        ret = MsgBox("the test's name is wrong." & vbCrLf _
            & "please selct correct test." _
            , vbExclamation, _
            "CmdStartSawTest_Click()")
    End If
        
    If AppKeyState <> remote Then
        ret = MsgBox("can not start saw cycle" & vbCrLf & "key is not in remote mode.", vbExclamation, "saw cycle")
        Exit Sub
    End If
    
    ret = SetServo(True)
    If ret = -1 Then
        ret = MsgBox("can not set servo on" & vbCrLf & "check door/window/E.Stop.", vbExclamation, "saw cycle")
        Exit Sub
    End If

    RunTestEndless = False
    
    bb = StartJob(CurrentTest)
    If bb = False Then
        ret = MsgBox("can not start saw cycle." & vbCrLf & "check communication or JOB in Controller", vbExclamation, "saw cycle")
        Exit Sub
    End If
    
End Sub

Private Sub CmdSysOperation_Click(Index As Integer)

Dim ret As Integer
Dim bb As Boolean


    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(44, 12, 12, "MdiMain", "CmdSysOperation_Click()", KeyPosition)
        Exit Sub
    End If
    
    ''****resume''****************************
    If Index = 0 Then
        ret = BscHoldOff(m_nCid)
        ret = BscContinueJob(m_nCid)
        
    ''****pause''*****************************
    ElseIf Index = 1 Then
        ret = BscHoldOn(m_nCid)
        
    ''****reset''*****************************
    ElseIf Index = 2 Then
    
        ret = MsgBox("You are about to reset all robot programs." & vbCrLf _
        & "Data will be initialize." & vbCrLf _
        & "Continue ? " _
        , vbQuestion + vbYesNo, _
            "Reset Robot Program.")
            If ret = 6 Then ''    6==YES
                ''
            ElseIf ret = 7 Then  ''7==no
                Exit Sub
            End If
            
        CurrentTest = ""

        CmdSysOperation(2).Enabled = False            ''  feedback the user
        ProgressBarReset.Visible = True
        ProgressBarReset.Value = 1
        
        Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''    pause the robot
        ProgressBarReset.Value = 1
        
        fMainForm.tmrRobotQuery.Enabled = False ''        pause the main timer
        ProgressBarReset.Value = 2
        
        fMainForm.tmrUpdateRobotStatus.Enabled = False '' pause the secondery timer
        ProgressBarReset.Value = 3
        
        Call ModHandShake.ResetAllJobsStatus
        ProgressBarReset.Value = 4
        
        Call ModHandShake.ResetAllProccess
        ProgressBarReset.Value = 5
        
        Call ResetRobotJob("2_AUTO_START_TEST.JBI")
        ProgressBarReset.Value = 6
        
        Call ResetRobotJob("3_CHUCK_TO_KIOSK.JBI")
        ProgressBarReset.Value = 7
        
        Call ResetRobotJob("3_CHUCK_TO_POCKET.JBI")
        ProgressBarReset.Value = 8
       
        Call ResetRobotJob("3_KIOSK_TO_CHUCK.JBI")
        ProgressBarReset.Value = 9
        
        Call ResetRobotJob("3_KIOSK_TO_POCKET.JBI")
        ProgressBarReset.Value = 10
        
        Call ResetRobotJob("3_POCKET_TO_CHUCK.JBI")
        ProgressBarReset.Value = 11
         
        Call ResetRobotJob("3_POCKET_TO_KIOSK.JBI")
        ProgressBarReset.Value = 12
        
        Call ResetRobotJob("4_TAKE_FROM_POCKET.JBI")
        ProgressBarReset.Value = 13
        
        Call ResetRobotJob("4_PLACE_ON_POCKET.JBI")
        ProgressBarReset.Value = 14
        
        Call ResetRobotJob("4_TAKE_FROM_CHUCK.JBI")
        ProgressBarReset.Value = 15
        
        Call ResetRobotJob("4_PLACE_ON_CHUCK.JBI")
        ProgressBarReset.Value = 16
        
        Call ResetRobotJob("4_PLACE_ON_SPINDLE.JBI")
        ProgressBarReset.Value = 17
        
        Call ResetRobotJob("4_TAKE_FROM_SPINDLE.JBI")
        ProgressBarReset.Value = 18
        
        Call ResetRobotJob("4_PLACE_ON_STATION.JBI")
        ProgressBarReset.Value = 19
        
        Call ResetRobotJob("4_TAKE_FROM_KIOSK.JBI")
        ProgressBarReset.Value = 20
        
        Call ResetRobotJob("4_TAKE_FROM_STATION.JBI")
        ProgressBarReset.Value = 21
        
        Call ResetRobotJob("4_PLACE_ON_KIOSK.JBI")
        ProgressBarReset.Value = 22
        
        Call ResetRobotJob("4_EXCHANGE_GRIPPER.JBI")
        ProgressBarReset.Value = 23
        
        Call ResetRobotJob("4_PARKING.JBI")
        ProgressBarReset.Value = 24
        
        Call ResetRobotJob("5_CALC_CHUCK.JBI")
        ProgressBarReset.Value = 25
        
        Call ResetRobotJob("5_CALC_KIOSK.JBI")
        ProgressBarReset.Value = 26
        
        Call ResetRobotJob("5_CALC_POCKET.JBI")
        ProgressBarReset.Value = 27
        
        Call ResetRobotJob("5_CURRENT_RETRUCT.JBI")
        ProgressBarReset.Value = 28
        
        Call ResetRobotJob("TEST_PLACE_ON_CHUCK.JBI")
        ProgressBarReset.Value = 29
        
        Call ResetRobotJob("TEST_TAKE_FROM_CHUCK.JBI")
        ProgressBarReset.Value = 30
        
        Call ResetRobotJob("3_TEST_KIOSK_TO_POCKET.JBI")
        ProgressBarReset.Value = 31
        
        Call ResetRobotJob("3_TEST_POCKET_TO_KIOSK.JBI")
        ProgressBarReset.Value = 32
        
        Call ResetRobotJob("3_TEST_POCKET_TO_CHUCK.JBI")
        ProgressBarReset.Value = 33
        
        Call ResetRobotJob("3_TEST_CHUCK_TO_POCKET.JBI")
        ProgressBarReset.Value = 34
        
        Call ResetRobotJob("3_TEST_POCKET_TO_POCKET.JBI")
        ProgressBarReset.Value = 35
        
        Call ResetRobotJob("3_TEST_ALL_POCKETS.JBI")
        ProgressBarReset.Value = 36
        ''Call FrmCommunication.ResetProfibus_Click
       
        Call ModGUI.ResetAllRobotStrings
        ProgressBarReset.Value = 37
        fMainForm.tmrRobotQuery.Enabled = True ''   resume the main timer
        ProgressBarReset.Value = 38
        ret = BscReset(m_nCid)     ''               Reset Robot alarm.
        ProgressBarReset.Value = 39
        Call ResetAllRobotRequest
        fMainForm.tmrUpdateRobotStatus.Enabled = True ''resume the secondery timer
        CmdSysOperation(2).Enabled = True  ''           feedback the user
        ProgressBarReset.Visible = False
        
    ''****start automat cycle '''***************
    ElseIf Index = 3 Then
    
        bb = modAllWorkPiece.WPExist ''check if there is atleast one WPIece.
        If bb = False Then
            Call FrmDialog.ShowDialogForm(44, 44, 44, "MdiMain", "CmdSysOperation_Click()", GeneralError)
            Exit Sub
        End If
    
        ''update the GUI
        pnlMode(0) = AllWP(1).NCProgram
        pnlMode(2) = AllWP(1).WPNumber
        
        ret = SetServo(True) ''give power to motors.
        If ret = -1 Then
            Call FrmDialog.ShowDialogForm(44, 11, 26, "mdiMain", "CmdSysOperation_Click()", EmergencyStop)
            Exit Sub
        End If
        
        ArrayInteger(0) = AllWP(AppWPIndex).ToolAmount ''   send tool amount
        Call WriteInteger(10, ArrayInteger())

        bb = ModHandShake.WriteDrillCode(AppDiameter)
        If bb = False Then
            Call FrmDialog.ShowDialogForm(44, 60, 60, "mdiMain", "CmdSysOperation_Click()", GeneralError)
            Exit Sub
        End If
            
        bb = ModHandShake.SendToolSensorState()
        If bb = False Then
            Call FrmDialog.ShowDialogForm(44, 61, 61, "mdiMain", "CmdSysOperation_Click()", GeneralError)
            Exit Sub
        End If
        
        If AppToolType <> HSK Then
            Call CmdSetOffset_Click ''send offset to controller
        End If
        
        bb = StartJob("2_AUTO_START_TEST.JBI")
        If bb = False Then
            Call FrmDialog.ShowDialogForm(44, 32, 32, "MdiMain", " CmdSysOperation_Click()", InputIncomplete)
            Exit Sub
        End If
       
      ''****Reset Profibus '''***************
    ElseIf Index = 4 Then
        
        Call FrmCommunication.ResetProfibus_Click
        
    End If
    
End Sub


Private Sub Exit_Click()

'cancel 2
'yes 1
Dim ret As Integer
Dim UserFdbk As fdbk

   
    UserFdbk = FrmDialog.ShowDialogForm(35, 35, 35, "mdimain", "Exit_Click()", ExitApp)
    If UserFdbk = UserYes Then
    
         Call ModCsvFile.SaveAutomation ''save the automation status into the HardDisk.
         Call ModCsvFile.SaveAllWP ''save the work piece table into the Hard disk.
         
        ''turn the timers off
        tmrUpdateRobotStatus.Enabled = False
        tmrRobotQuery.Enabled = False
        
        ArrayByte(0) = 1
        Call WriteByte(60, ArrayByte)
        Call WriteByte(61, ArrayByte)
    
        Call ModHandShake.ResetAllRobotRequest
        MotoComToolBox.CloseCommunication ''close communication with grace
        ModLogFile.LogAddLine ("End Application:" & CStr(Format(now, "hh:mm:ss")))
        ModLogFile.WriteLogFile

        Me.Hide ''   close the main Form
        Unload Me
        End ''       end application
    Else
            '''      do nothing
    End If
    
End Sub


Private Sub FrameLocation_DragDrop(Source As Control, X As Single, Y As Single)

End Sub








Private Sub IncRobotStep_Click(Index As Integer)
''  index 0:   0.1 mm
''  index 1:   1.0 mm
''  index 2:   2.0 mm
''  index 3:   5.0 mm

Dim StepValue As Double
    
    Select Case Index
    Case 0
        IncStep = 0.1
    Case 1
        IncStep = 1
    Case 2
        IncStep = 2
    Case 3
        IncStep = 5
    Case 4
        IncStep = 10
    End Select
    
End Sub








Private Sub MDIForm_Load()
''
Debug.Print "MDIForm_Load()"
''
Dim BoolShelvs(3) As Boolean
Dim jj As Integer

  
    Call Grid1Display ''                 display the workPiece Table.
    Call Grid2Display ''                 display the pocket status list.
    Call Grid3Display ''                 General Location Table
    Call DisplayTablePocketsLocations '' pocket locations.
    Call DisplayTableTeach  '            the first and last pocket.
    ' Call CenterForm 'elisha
    
    MotoComToolBox.SetCommunicationParameters
    MotoComToolBox.StartCommunication
    
    Call LoadComboAllStatus
    Call LoadComboSingleStatus
    Call LoadComboLocations
    '''Call LoadComboGeneralLocations ' <<<>>>
    
    Call fMainForm.BtnShowGeneralLocations_Click ''      display general locations on the screen.
    Call fMainForm.BtnViewPocketsLocations_Click ''      view pockets on the screen.
    Call fMainForm.BtnDisplayWPTable_Click   ''          display AllWorkpiece on the screen.
    Call fMainForm.BtnDisplayPocketStatus_Click   '''    status table.
    Call fMainForm.BtnShowFirstLastPockets_Click
    Call CmdRestoreDefult_Click
    
    'wake up as manual mode
    TopToolBar.Buttons.Item(5).Value = tbrPressed
    TopToolBar.Buttons.Item(4).Value = tbrUnpressed
    TopToolBar.Buttons.Item(3).Value = tbrUnpressed
    
    Update_ASM_GUI (1)
    
    ''fMainForm.LabelApp(1).Caption = "Diameter : " & CStr(AppDiameter)
    
    ModGUI.DisplayAppToolType
    
    Call fMainForm.BtnDisplayPocketStatus_Click '''display all pockets statuses.
    Call fMainForm.BtnDisplayWPTable_Click   ''    display AllWP() data on screen.
    
    ''start the timers of the application.
    fMainForm.tmrRobotQuery.Enabled = True
    fMainForm.tmrUpdateRobotStatus.Enabled = True
    
    If AppSimulation = True Then
        fMainForm.FrameSimulator.Visible = True
    ElseIf AppSimulation = False Then
        fMainForm.FrameSimulator.Visible = False
    End If
    
    
End Sub
 

Private Sub MDIForm_Unload(Cancel As Integer)

   

    ''Close #LogErrFileNumber

'    Do While Forms.Count > 1
'        Unload Forms(1)
'    Loop
  
End Sub

Private Sub CenterForm()
Dim Gleft As Integer
Dim Gtop As Integer
   ''gtop is the location of the topmost elemnt y position.
   
    ''center the elements horizontaly
    Gleft = (Screen.Width - SSTab2.Width) / 2
    SSTab2.Left = Gleft
    pnlGembal(0).Left = Gleft
    pnlMod.Left = Gleft + SSTab2.Width - pnlMod.Width
    
    fMainForm.LabelApp(1).Left = Gleft + 0.5 * (SSTab2.Width) - 0.5 * (fMainForm.LabelApp(1).Width) + 150
    LabelApp(1).Left = fMainForm.LabelApp(1).Left
    Picture1.Height = Screen.Height
    
    
    '' center the elements vertically
    Gtop = (Screen.Height - (pnlGembal(0).Height + 200 + SSTab2.Height)) / 2 - 400
    pnlGembal(0).Top = Gtop
    SSTab2.Top = Gtop + pnlGembal(0).Height + 200
    pnlMod.Top = Gtop
    LabelApp(0).Top = Gtop
    LabelApp(2).Top = LabelApp(0).Top + 600
    
    
End Sub


Private Sub MSFlexGrid1_Click()
'the table in the 'orderlist'
Dim i As Integer

    ''set background white
    MSFlexGrid1.BackColor = &HFFFFFF
    
    'get the selected line
    i = MSFlexGrid1.RowSel
    
    ''put the line number in the text box
    TextLineNumber.Text = i
    txtWorkPiece(6).Text = i
    
    ''unable/disable the Priority Buttons.
    If (CInt(TextLineNumber.Text = 1)) Then
        cmdMove(6).Enabled = False
    Else
        cmdMove(6).Enabled = True
    End If
    
    txtWorkPiece(2) = AllWP(i).WPNumber
    txtWorkPiece(3) = AllWP(i).NCProgram
    txtWorkPiece(4) = AllWP(i).ToolDiameter
    txtWorkPiece(5) = AllWP(i).ToolAmount
    txtWorkPiece(6) = AllWP(i).LineNumber
    
    
End Sub









Private Sub Option1_Click(Index As Integer)

Dim jj As Integer

    Select Case Index
    
    Case 0
        CurrentTest = "3_TEST_KIOSK_TO_POCKET.JBI"
        
    Case 1
        CurrentTest = "3_TEST_POCKET_TO_KIOSK.JBI"
        
    Case 2
        CurrentTest = "3_TEST_POCKET_TO_CHUCK.JBI"
        
    Case 3
        CurrentTest = "3_TEST_CHUCK_TO_POCKET.JBI"
        
    End Select
            
        RunTestEndless = False
        
    For jj = 0 To 3
        OptionLoop(jj).Value = vbUnchecked
    Next jj

End Sub






Private Sub OptionSaw_Click(Index As Integer)

    Select Case Index
    
    Case 0
        CurrentTest = "3_TEST_STATION_TO_SPINDLE.JBI"
        
    Case 1
        CurrentTest = "3_TEST_SPINDLE_TO_STATION.JBI"
        
    End Select
    
    RunTestEndless = False
    
End Sub

Private Sub optNightMode_Click(Index As Integer)
'0 on
'1 off
Dim ret As Integer
On Error GoTo label

    ''prevent the user to click the ON contradiction.
    If optOneTool(0) = True Then
        optNightMode(0) = False
        optNightMode(1) = True
        Exit Sub
    End If
    
    ''update the TextBox
    If optNightMode(0) = True Then
        LabelMode.Caption = "Night Mode"
        HermleWPMode = NightMode
        Call WriteOneCommByte(55, AppWPMode.NightMode) '''=3
    Else
       If optOneTool(0) = True Then
            LabelMode.Caption = "One Tool"
            HermleWPMode = OneTool
            Call WriteOneCommByte(55, AppWPMode.OneTool) ''=1
       Else
            LabelMode.Caption = "Work Piece"
            HermleWPMode = WorkPiece
            Call WriteOneCommByte(55, AppWPMode.WorkPiece) ''=2
       End If
    End If
        
Exit Sub
label:
    ret = MsgBox("The System was unable to change Night Mode." & vbCrLf & _
      "Try one of the options below :" & vbCrLf & _
      "1.Communication problem with the controller." & vbCrLf & _
      "2.wrong variable name." & vbCrLf & _
      "  (mdiMain.optNightMode_Click)", vbExclamation, "Error in Change NightMode")
End Sub

Private Sub optOneTool_Click(Index As Integer)
'0 on
'1 off
Dim ret As Integer
On Error GoTo label

    ''prevent the user to click the ON contradiction.
    If optNightMode(0) = True Then
        optOneTool(0) = False
        optOneTool(1) = True
        Exit Sub
    End If
    
     ''update the TextBox
    If optOneTool(0) = True Then
        LabelMode.Caption = "One Tool"
        HermleWPMode = OneTool
        Call WriteOneCommByte(55, AppWPMode.OneTool)
        Call ModGUI.ListHMIUpdate(" send AppWPMode : One Tool")
        
    Else
       If optNightMode(0) = True Then
            LabelMode.Caption = "Night Mode"
            HermleWPMode = NightMode
            Call WriteOneCommByte(55, AppWPMode.NightMode)
            Call ModGUI.ListHMIUpdate(" send AppWPMode : Night Mode")
       Else
            LabelMode.Caption = "Work Piece"
            HermleWPMode = WorkPiece
            Call WriteOneCommByte(55, AppWPMode.WorkPiece)
            Call ModGUI.ListHMIUpdate(" send AppWPMode : Work Piece")
       End If
    End If
        
Exit Sub
  
label:
    ret = MsgBox("Error while Change WorkPiece mode." & vbCrLf _
    & "the error is : " & Err.Description _
    , vbExclamation, _
    "optOneTool_Click")
      
End Sub

Private Sub REStop_Click()

Dim ret As Integer

    ''check the key state
    If AppKeyState <> remote Then
        ret = MsgBox("the key is not in REMOTE mode." _
            & vbCrLf & "Set Key to remote." _
            , vbExclamation, "REStop_Click()")
        Exit Sub
    End If
    
    ret = SetServo(True)
    
    If ret <> 0 Then
        Call FrmDialog.ShowDialogForm(68, 68, 68, "mdiMain", "REStop_Click()", EmergencyStop)
        Exit Sub
    End If
    
    Call MotoComToolBox.StartJob("SERVOFLOAT_OFF.JBI")
        
    
End Sub

Private Sub ScrollShelf_Change()
    
'    TextShelf.Text = ScrollShelf.Value
        
End Sub









Private Sub SliderAutoSpeed_Change()
''1.the automatic speed  and the manual speed are bound.
''2.the
Dim ret As Integer

On Error GoTo out

    LableAutoSpeed.Caption = SliderAutoSpeed.Value 'update the text
    AppSpeed = SliderAutoSpeed.Value 'get value from GUI.
   
    SliderManuSpeed.Value = AppSpeed 'update the Manual Speed slider
    
    ArrayInteger(0) = AppSpeed '''get ready to send.
    Call WriteInteger(12, ArrayInteger) 'send value to controller
    
 Exit Sub
 
out:
  ret = MsgBox("The System was unable to change the speed." & vbCrLf & _
  "Try to fix one of the options below :" & vbCrLf & _
  "1.Communication problem with the controller." & vbCrLf & _
  "2.type mismatch of vars." & vbCrLf & _
  "  (mdiMain.SliderAutoSpeed_Change)", vbExclamation, "Error in changing the Robot Speed")
  
End Sub

Private Sub SliderManuSpeed_Change()
''1.the automatic speed  and the manual speed are bound.
Dim ret As Integer
Dim bb As Boolean

On Error GoTo out

    LabelPercent.Caption = SliderManuSpeed.Value 'update the text
    AppSpeed = SliderManuSpeed.Value 'get value from GUI.
   
    SliderAutoSpeed.Value = AppSpeed 'update the Manual Speed slider
    
    ArrayInteger(0) = AppSpeed '''get ready to send.
    bb = WriteInteger(12, ArrayInteger) 'send value to controller
    If bb = False Then
        ''Call FrmDialog.ShowDialogForm(64, 64, 64, "MdiMain", "SliderManuSpeed_Change()", KeyPosition)
    Else
        Exit Sub
    End If
        
    
Exit Sub

out:
  ret = MsgBox("The System was unable to change the speed." & vbCrLf & _
  "Try to fix one of the options below :" & vbCrLf & _
  "1.Communication problem with the controller." & vbCrLf & _
  "2.type mismatch of vars." & vbCrLf & _
  "  (mdiMain.SliderManuSpeed_Change)", vbExclamation, "Error in changing the Robot Speed")
  
End Sub

Private Sub Grid1Display()
''1.the function display the workPiece Table.
Debug.Print "Grid1Display()"
Dim i%, W%

    For i = 0 To 4
        MSFlexGrid1.ColWidth(i) = 950
    Next i
    MSFlexGrid1.RowHeight(0) = 600
    MSFlexGrid1.ColWidth(2) = 1200

'    MSFlexGrid1.Height = MSFlexGrid1.RowHeightMin * (MSFlexGrid1.Rows - 1) + MSFlexGrid1.RowHeight(0) + 45
    MSFlexGrid1.Height = 4125
    For i = 0 To MSFlexGrid1.Cols - 1
        W = W + MSFlexGrid1.ColWidth(i)
        MSFlexGrid1.ColAlignment(i) = flexAlignCenterCenter
        MSFlexGrid1.Col = i
        MSFlexGrid1.Row = 0
        MSFlexGrid1.WordWrap = True
        'MSFlexGrid1.TextMatrix(i + 1, 0) = "Sub-" & i
    Next i
    
    MSFlexGrid1.Width = W ''+ 300
    MSFlexGrid1.TextStyleFixed = flexTextFlat
    MSFlexGrid1.ForeColor = vbNormal
    MSFlexGrid1.Appearance = flexFlat
    MSFlexGrid1.TextMatrix(0, 0) = "Number"
    MSFlexGrid1.TextMatrix(0, 1) = "Work Piece"
    MSFlexGrid1.TextMatrix(0, 2) = "NC Program"
    MSFlexGrid1.TextMatrix(0, 3) = "Diameter"
    MSFlexGrid1.TextMatrix(0, 4) = "Amount"
    MSFlexGrid1.TextMatrix(0, 5) = "Tool"


    MSFlexGrid1.Col = MSFlexGrid1.Cols - 1
  
    For i = 0 To MSFlexGrid1.Rows - 2
        MSFlexGrid1.TextMatrix(i + 1, 0) = i + 1
    Next i
   

End Sub

Private Sub Grid2Display()
''1.the function display the pocket status list.
Debug.Print "Grid2Display()"

Dim i%, W%

    ''set the width of column
    For i = 0 To 4
        MSFlexGrid2.ColWidth(i) = 1000
    Next i
    
    MSFlexGrid2.ColWidth(3) = 1250
    
    ''set the height of header.
    MSFlexGrid2.RowHeight(0) = 600
    
    ''set the height of the table
    MSFlexGrid2.Height = MSFlexGrid2.RowHeightMin * (MSFlexGrid2.Rows - 1) + MSFlexGrid2.RowHeight(0) + 450
    
    For i = 0 To MSFlexGrid2.Cols - 1
        W = W + MSFlexGrid2.ColWidth(i)
        MSFlexGrid2.ColAlignment(i) = flexAlignCenterCenter
        MSFlexGrid2.Col = i
        MSFlexGrid2.Row = 0
        MSFlexGrid2.WordWrap = True
    Next i
    
    'set the table width
    MSFlexGrid2.Width = W + 50
    
    MSFlexGrid2.TextStyleFixed = flexTextFlat
    MSFlexGrid2.ForeColor = vbNormal
    MSFlexGrid2.Appearance = flexFlat
    MSFlexGrid2.TextMatrix(0, 0) = "Pocket"
    MSFlexGrid2.TextMatrix(0, 1) = "Work Piece"
    MSFlexGrid2.TextMatrix(0, 2) = "Diameter"
    MSFlexGrid2.TextMatrix(0, 3) = "Status"
    MSFlexGrid2.TextMatrix(0, 4) = "Program"

    MSFlexGrid2.Col = MSFlexGrid2.Cols - 1
  


End Sub

Private Sub Grid3Display()
'1.this function set PreDisplay options like size,center,titls,,,
'   of the General Location Table.
Debug.Print "Grid3Display()"
Dim i%, W%


    MSFlexGrid3.ColWidth(0) = 900
    MSFlexGrid3.ColWidth(1) = 2000
    MSFlexGrid3.ColWidth(2) = 900
    MSFlexGrid3.ColWidth(3) = 900
    MSFlexGrid3.ColWidth(4) = 900

    MSFlexGrid3.RowHeight(0) = 600
    
    For i = 0 To MSFlexGrid3.Cols - 1
        W = W + MSFlexGrid3.ColWidth(i)
        MSFlexGrid3.ColAlignment(i) = flexAlignCenterCenter
        MSFlexGrid3.Col = i
        MSFlexGrid3.Row = 0
        MSFlexGrid3.WordWrap = True
    Next i
    
    MSFlexGrid3.Width = 5625
    MSFlexGrid3.Height = 2075
    
    MSFlexGrid3.TextStyleFixed = flexTextFlat
    MSFlexGrid3.ForeColor = vbNormal
    MSFlexGrid3.Appearance = flexFlat
    
    MSFlexGrid3.TextMatrix(0, 0) = "Number"
    MSFlexGrid3.TextMatrix(0, 1) = "Location"
    MSFlexGrid3.TextMatrix(0, 2) = "  X "
    MSFlexGrid3.TextMatrix(0, 3) = "  Y "
    MSFlexGrid3.TextMatrix(0, 4) = "  Z "

    MSFlexGrid3.Col = MSFlexGrid3.Cols - 1

End Sub


Private Sub cmdUnloadTool_Click()
''1.the function run RobotJob that take tool from the Correct pocket and put it in the kiosk.
''2.the function being called from the GUI.
''3.the function take no parameters.
''4.the f return no parameters.
Dim ret As Integer
Dim iWPiece As Integer
Dim iPocketNumber As Integer
Dim shelf_from As Integer
Dim column_from As Integer
Dim pocket_from As Integer
Dim MyPocket As String
Dim bb As Boolean

On Error GoTo labelUnload

''
    fMainForm.tmrRobotQuery.Enabled = True   ''resune the main timer
    
    ''check the key state
    If AppKeyState <> remote Then
        Call FrmDialog.ShowDialogForm(14, 12, 12, "MdiMain", "cmdUnloadTool_Click()", KeyPosition)
        Exit Sub
    End If
    
    If GripperStatus = 0 Then
        ret = MsgBox("The Gripper is Close." & vbCrLf _
        & "Pleae Open Gripper.", vbExclamation, "gripper state")
        Exit Sub
    End If
    
     ''check the servo state
    ret = SetServo(True)
    If ret = -1 Then
        Call FrmDialog.ShowDialogForm(14, 11, 26, "mdiMain", "cmdUnloadTool_Click()", EmergencyStop)
        Exit Sub
    End If
    
    ' check correct input from user:'fields should be integers
    If IsNumeric(txtLoadUnloadPocket.Text) = False Or txtLoadUnloadPocket.Text = "" Then
        ret = MsgBox("Wrong Pocket number ", vbInformation, "Error in Loading Tool To Pocket.")
        Exit Sub 'GoTo labelLoad
    End If
    ' check correct input from user:'fields should be integers
    If IsNumeric(txtWorkPiece(1).Text) = False Or txtWorkPiece(1).Text = "" Then
        ret = MsgBox("Wrong Workpiece number ", vbInformation, "Error in Loading Tool To Pocket.")
        Exit Sub 'GoTo labelLoad
    End If

    ''check if the shelf tooltype is the same as the app tool type.
    shelf_from = CInt(fMainForm.txtLoadUnloadShelf.Text)
    If AppShelvs(shelf_from).ShelfToolType <> AppToolType Then
        MsgBox "the shelf is incorrect.the shelf type is not the same as the gripper type", vbCritical
        Exit Sub
    End If
    
    'check if Work Piece is (1-100)
    If CInt(txtWorkPiece(1).Text) > 100 Or CInt(txtWorkPiece(1).Text) < 1 Then
        ret = MsgBox("Wrong Workpiece number ", vbInformation, "Error in Loading Tool To Pocket.")
        Exit Sub 'GoTo labelLoad
    End If
    
    ' check if pocket number is between the bounds.
    If AppToolType = HSK Then
        If (CInt((txtLoadUnloadPocket.Text)) >= 101 And CInt((txtLoadUnloadPocket.Text)) <= 110) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 201 And CInt((txtLoadUnloadPocket.Text)) <= 210) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 301 And CInt((txtLoadUnloadPocket.Text)) <= 310) _
            Then
        Else
            ret = MsgBox("Wrong Pocket number", vbInformation, "Error in Loading Tool To Pocket.")
            Exit Sub 'GoTo labelLoad
        End If

    ElseIf AppToolType = Drill Then
        If (CInt((txtLoadUnloadPocket.Text)) >= 101 And CInt((txtLoadUnloadPocket.Text)) <= 112) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 201 And CInt((txtLoadUnloadPocket.Text)) <= 212) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 301 And CInt((txtLoadUnloadPocket.Text)) <= 312) _
           Then
        Else
            ret = MsgBox("Wrong drill Pocket number", vbInformation, "Error in Loading Tool To Pocket.")
            Exit Sub 'GoTo labelLoad
        End If
        
    ElseIf AppToolType = Round Then
        If (CInt((txtLoadUnloadPocket.Text)) >= 101 And CInt((txtLoadUnloadPocket.Text)) <= 112) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 201 And CInt((txtLoadUnloadPocket.Text)) <= 212) Or _
           (CInt((txtLoadUnloadPocket.Text)) >= 301 And CInt((txtLoadUnloadPocket.Text)) <= 312) _
           Then
        Else
            ret = MsgBox("Wrong round Pocket number", vbInformation, "Error in Loading Tool To Pocket.")
            Exit Sub 'GoTo labelLoad
        End If
        
    End If

    Call SetServo(True)
        
    bb = ModHandShake.WriteDrillCode(AppDiameter)
    If bb = False Then
        Call FrmDialog.ShowDialogForm(14, 60, 60, "mdiMain", "cmdUnloadTool_Click()", GeneralError)
        Exit Sub
    End If
            
    bb = ModHandShake.SendToolSensorState()
    If bb = False Then
        Call FrmDialog.ShowDialogForm(14, 61, 61, "mdiMain", "cmdUnloadTool_Click()", GeneralError)
        Exit Sub
    End If
   
     
    If Jobs.AutoStart(0) = JobState.RUN Then
        SetOneCommByte (26)
    ElseIf Jobs.AutoStart(0) <> JobState.RUN Then
        bb = StartJob("3_POCKET_TO_KIOSK.JBI")
        If bb = False Then
            Call FrmDialog.ShowDialogForm(14, 32, 32, "mdimain", "cmdUnloadTool_Click()", InputIncomplete)
            Exit Sub
        End If
    End If
    
   Exit Sub
    
labelUnload:
   ret = MsgBox("The System was unable to UnLoad Tool To  Pocket," & vbCrLf _
        & "1.Check the Inputs.should not be empty." & vbCrLf _
        & "2.Check the Inputs.should be numbers only." & vbCrLf _
        & "3.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.cmdLoadTool_Click]", vbInformation, "Error in UnLoading Tool To Pocket.")


End Sub




Private Sub SSPanel_Click(Index As Integer)
'by pressing the Panel the textBox Erase.

    If Index = 1 Then
        txtWorkPiece(1).Text = ""
    End If
    
    If Index = 0 Then
        txtLoadUnloadPocket.Text = "101"
    End If
    
End Sub

Private Sub SSTab1_LostFocus()

Dim kk As Integer

    'once the user left the tab the image disapear.
    ImgFdBkWPiece.Visible = False
    
End Sub

Private Sub SSTab2_Click(PreviousTab As Integer)
''1.this function get the current tab number and lunch
''   a function according to the current tub.
''2.there is no use to the previous tab.
    
   
Dim CurrentTab As Integer

    ''get the current tub.
    ''add +1 to the tub number because its start from 0.
    CurrentTab = SSTab2.Tab + 1
    
    Select Case CurrentTab
       
        Case 1
            FlagInfo = False
            FlagDiagnostic = False
        
        Case 2
            FlagInfo = False
            FlagDiagnostic = False
                
            If AppSimulation = True Then
                fMainForm.FrameSimulator.Visible = True
            ElseIf AppSimulation = False Then
                fMainForm.FrameSimulator.Visible = False
            End If
    
           
        Case 3
            FlagInfo = False
            FlagDiagnostic = False
            Call BtnDisplayWPTable_Click
            Call CmdRestoreDefult_Click
            
        Case 4
            FlagInfo = False
            FlagDiagnostic = False
            IncRobotStep(2).Value = True
            
                        
        Case 5
            FlagInfo = False
            FlagDiagnostic = False
            Call BtnDisplayPocketStatus_Click
            
        Case 6
            FlagInfo = False
            FlagDiagnostic = False
            
            Call BtnShowGeneralLocations_Click
            Call BtnShowFirstLastPockets_Click
            Call BtnViewPocketsLocations_Click
                        
        Case 7
            FlagInfo = False
            FlagDiagnostic = True
        
        Case 8
            FlagInfo = True
            FlagDiagnostic = False
            
        Case 9 ''tests
            If ((UseSaw = True) And (HermleGear = Semi)) Then
                FrameSaw.Enabled = True
            Else
                FrameSaw.Enabled = False
            End If
    
        Case 10
            FlagInfo = True
            FlagDiagnostic = False
            
    End Select
    
End Sub


Private Sub SSTab3_Click(PreviousTab As Integer)
   
    If SSTab3.TabIndex = 0 Then
        TextToolsDiameter.Visible = False
        Call BtnShowGeneralLocations_Click
    End If
    
    If SSTab3.TabIndex = 1 Then
        TextToolsDiameter.Visible = False
        Call BtnShowFirstLastPockets_Click

    End If
    
    If SSTab3.TabIndex = 2 Then
        TextToolsDiameter.Visible = False
        Call BtnViewPocketsLocations_Click
    End If
    
End Sub









Private Sub textChangeWorkPiece_DblClick()
CurrentTBox = 2
    
    'the title to be shown in the Virtual KeyBoard.
    psTitle = "Change Status For All.Enter Work Piece Number."
    
    'the text in the calling textBox.
    psValue = textChangeWorkPiece.Text
    
    'display the virtual KB.
    frmKeypadAlph.Show
    
    'display default settings of Virtual KB.
    Call frmKeypadAlph.Form_Activate
    
End Sub



Private Sub TextPocketNumber_Change()
    TextPocketNumber = UpDnShelf.Value * 100 + UpDownPocket.Value
End Sub

Private Sub TextPocketNumber_DblClick()
CurrentTBox = 5
''the function display the Virtual keyboard.

    'the title to be shown in the Virtual KeyBoard.
    psTitle = "Teach Single Pocket.Pocket number."
    
    'the text in the calling textBox.
    psValue = TextPocketNumber.Text
    
    'display the virtual KB.
    frmKeypadAlph.Show
    
    'display default settings of Virtual KB.
    Call frmKeypadAlph.Form_Activate
    
   
End Sub


Private Sub TextToolsDiameter_DblClick()
CurrentTBox = 6
''the function display the Virtual keyboard.

    'the title to be shown in the Virtual KeyBoard.
    psTitle = "Teach Single Pocket.Tool Diameter."
    
    'the text in the calling textBox.
    psValue = TextToolsDiameter.Text
    
    'display the virtual KB.
    frmKeypadAlph.Show
    
    'display default settings of Virtual KB.
    Call frmKeypadAlph.Form_Activate
       
End Sub


Private Sub tmrQuesyStatus_Timer()
    If Me.tmrRobotQuery.Enabled = True And imgQuesryTimer.Visible = False Then
        imgQuesryTimer.Visible = True
    ElseIf Me.tmrRobotQuery.Enabled = False And imgQuesryTimer.Visible = True Then
        imgQuesryTimer.Visible = False: lblUpdateTime(2) = "---"
    End If
    'Refresh.me
End Sub

Private Sub tmrUpdateRobotStatus_Timer()
    
    If HermleCommState = OffLine Then
        Exit Sub
    End If
    
    fMainForm.TextListHmiCounter.Text = fMainForm.ListHMIInfo.ListCount
    If fMainForm.ListHMIInfo.ListCount > 32000 Then
        fMainForm.ListHMIInfo.Clear
    End If
        
    tmrUpdateRobotStatus.Interval = 3000
    imgPlcUpdate(1).Visible = True
    
    T21 = Timer ': T23 = (T21 - T22): lblUpdateTime(1) = Format(T23, "###.000")
    'T22 = T21
    'If tmrUpdateRobotStatus.Interval < T23 Then imgRedLampSource(1).ZOrder Else imgGrLampSource(1).ZOrder


    ' Read robot status take 30msec*??=??msec
    ' set timer > 2000 msec
    Tstatus(0) = Timer
    Tstatus(1) = Timer '- Tstatus(0)

    Call ModHandShake.ReadToolCounter
    Call DisplayTCPPosition '  30ms read
    Call DisplayJobName '      30ms read
    Call DisplayLineNumber '   30ms read
    Call DisplayServoStatus '  30ms read
    Call DisplayKeyState '     30ms read
    Call DisplayGripperState ' 30ms read
    Call DisplayWPProperties  '0 ms - GUI only
    Call DisplayAlarmStatus
    ''Call ModGUI.DisplayToolAmount

    Tstatus(2) = Timer '- Tstatus(1)
 
    ''taking care of the JOBS mechanism.
    Call ReadAllJobsStatus '            HandShake-read. 0.45 sec ms
    Call ModHandShake.ReadRobotStrings
    Call ModGUI.DisplayStringFromRobot
    Call UpdateGUIByJobsStatus '          0  ms - GUI only
    '''Call UpdatePocketsStatusByJobState '0  ms - Data Base only
    Call ResetJobStatus '                '30 ms - HandShake - Write
        
    Tstatus(3) = Timer '- Tstatus(2)
 
    ''take care of the TESTS mechanism =>480 ms
    If HermleGear = Semi Then
        Call ModHandShake.ReadAllTestsStatus '''                7x30=210 ms
        Call ModHandShake.RunTestsJobs '''                      8x30=240 ms
        Call ModHandShake.ResetSingleTestStatus(CurrentTest) '' 1x30=30  ms
    End If
  
    Tstatus(4) = Timer
  
    If SSTab2.Tab = 0 Then ' Automat

    ElseIf SSTab2.Tab = 1 Then ' 'tools' Semi automat
'        ReadString (14): PanelK2P(3).Caption = ArrayString
'        ReadString (15): PanelP2k(3).Caption = ArrayString
        
    ElseIf SSTab2.Tab = 2 Then '''WorkPiece
    
    ElseIf SSTab2.Tab = 3 Then '''manual
    
    ElseIf SSTab2.Tab = 4 Then '''pocket status
    
    ElseIf SSTab2.Tab = 5 Then '''teach
    
    ElseIf SSTab2.Tab = 6 Then '''diagnostic
        Call ModGUI.DisplayAllInputs
        
    ElseIf SSTab2.Tab = 7 Then '''operation
         
    ElseIf SSTab2.Tab = 8 Then ' Test programs
    
    ElseIf SSTab2.Tab = 9 And Me.SSTab5.Tab = 0 Then ' info
    End If
        
    Tstatus(5) = Timer
 
    T23 = (Timer - T21): lblUpdateTime(1) = Format(T23, "###.000")
    
    If tmrUpdateRobotStatus.Interval < T23 Then
        imgRedLampSource(1).ZOrder
    Else
        imgGrLampSource(1).ZOrder
    End If

    imgPlcUpdate(1).Visible = False

End Sub


Private Sub tmrRobotQuery_Timer()

    imgPlcUpdate(0).Visible = True
    T11 = Timer
  
    If HermleCommState = OffLine Then
        Exit Sub
    End If
    
    AppDiameter = AllWP(AppWPIndex).ToolDiameter
    If AppDiameter = 0 Then
            AppWPIndex = 1
    End If
        
    Call ModHandShake.ReadFirstLoadUnload ''        (kiosk-pocket)
    ' Read robot query take 30msec*6=180msec' set timer > 250 msec
    
    If (HermleGear = Automat) Then
        Call ReadAllRobotRequest   'HandShake-Read Requst from robot
        Call WriteDataToRobot      'HandShake-Write answer to robot.
    End If
    
    imgPlcUpdate(0).Visible = False
    T13 = (Timer - T11)
    lblUpdateTime(0) = Format(T13, "###.000")
    lblUpdateTime(2) = Format(T13, "###.000")
    
    If tmrRobotQuery.Interval < T13 Then
        imgRedLampSource(0).ZOrder
    Else
        imgGrLampSource(0).ZOrder
    End If

End Sub

Private Sub TopToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim ret As Integer

On Error GoTo LabelTBar

    Select Case Button
    
        Case "Shafir"
            Call AboutApplication_Click
            
        Case "Automat"
            HermleGear = Automat
            Call Update_ASM_GUI(HermleGear)
            Call WriteOneCommByte(51, 3)
            ArrayInteger(0) = CDbl(7) ''set value to be use in the System in the Robot
            Call WriteInteger(60, ArrayInteger()) ''write integer to robot
            
        Case "Semi"
            HermleGear = Semi
            Call Update_ASM_GUI(HermleGear)
            Call WriteOneCommByte(51, 2)
            ArrayInteger(0) = CDbl(7) ''set value to be use in the System in the Robot
            Call WriteInteger(60, ArrayInteger()) ''write integer to robot
            
        Case "Manual"
            HermleGear = Manual
            Call Update_ASM_GUI(HermleGear)
            Call WriteOneCommByte(51, 1)
            ArrayInteger(0) = CDbl(27) ''set value to be use in the System in the Robot
            Call WriteInteger(60, ArrayInteger()) ''write integer to robot
            'Call CmdSysOperation_Click(2)                ''reset application
            '''Call CmdSysOperation_Click(1) '            ''pause the robot
            '''FrmCommunication.ResetProfibus_Click       ''reset profibus
            
        Case "R.E.Stop"
            Call REStop_Click
                    
        Case "Options"
            frmOptions.Show
            
        Case "Communication"
            FrmCommunication.Show
           
        Case "Exit"
            Call Exit_Click
            
    End Select
    
    TopToolBar.Refresh
   
   Exit Sub
'''1 manual
'''2 semi
'''3 auto

    
LabelTBar:
   ret = MsgBox("The System was unable to Press the Button," & vbCrLf _
        & "Try to fix one of te Bellow :" & vbCrLf _
        & "1.Wrong Command." & vbCrLf _
        & "2.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.TopToolBar_ButtonClick]", vbInformation, "Error in Press the Top ToolBar.")

End Sub



Private Sub TxtSinglePocketNumber_DblClick()

CurrentTBox = 3

    'the title to be shown in the Virtual KeyBoard.
    psTitle = "Change Single status.Pocket Number."
    
    'the text in the calling textBox.
    psValue = textChangeWorkPiece.Text
    
    'display the virtual KB.
    frmKeypadAlph.Show
    
    'display default settings of Virtual KB.
    Call frmKeypadAlph.Form_Activate
    
End Sub


Private Sub TxtSingleToolDiameter_DblClick()

CurrentTBox = 4

    'the title to be shown in the Virtual KeyBoard.
    psTitle = "Change Single Status.Tool Diameter."
    
    'the text in the calling textBox.
    psValue = textChangeWorkPiece.Text
    
    'display the virtual KB.
    frmKeypadAlph.Show
    
    'display default settings of Virtual KB.
    Call frmKeypadAlph.Form_Activate
    
    
End Sub

Private Sub TxtViewPocket_Change()
    TextPocketNumber = UpDnShelf.Value * 100 + UpDownPocket.Value
End Sub

Private Sub txtWorkPiece_Change(Index As Integer)

Dim ret As Integer
Dim TempInt As Integer


    If Index = 1 Then
        AppToolsWPiece = CInt(fMainForm.txtWorkPiece(1).Text)
    End If
    
    If ((Index >= 2) And (Index <= 6)) Then
     
        ''check if the input is a number
        ret = IsNumeric(txtWorkPiece(Index).Text)
        If ret = 0 Then
            txtWorkPiece(Index).Text = ""
            Exit Sub
        End If
        
        ''check if the input is a number (between 0 to 9)
        If ((Index >= 2) And (Index <= 6)) Then
            TempInt = Asc(txtWorkPiece(Index).Text)
                If ((TempInt > 57) Or (TempInt < 48)) Then
                    txtWorkPiece(Index).Text = ""
                    Exit Sub
                End If
        End If
   
    End If
    
        ImgFdBkWPiece.Visible = False
    
End Sub

Private Sub txtWorkPiece_DblClick(Index As Integer)

Dim tempstring1 As String
Dim tempstring2 As String
CurrentTBox = 1


    tempstring1 = "" '"Add New WorkPiece. "
    
    Select Case Index
        Case 0
            tempstring2 = "Pocket"
        Case 1
            tempstring2 = "WorkPiece"
        Case 2
            tempstring2 = "WorkPiece"
        Case 3
            tempstring2 = "NC Program"
        Case 4
            tempstring2 = "Tool Diameter"
        Case 5
            tempstring2 = "Tool Amount"
        Case 6
            tempstring2 = "Line Number"
    End Select
    
    msValue = Index
    psTitle = tempstring1 & tempstring2 'the title of the Vir KB
    psValue = txtWorkPiece(Index).Text
    
    ''feedBack the User.disappear the green 'V' .
    ImgFdBkWPiece.Visible = False
    
    frmKeypadAlph.Show
    Call frmKeypadAlph.Form_Activate
         
End Sub

Private Sub LoadDefalt()
''the func display the default settings.
''the function is beeing called by :MDIForm_Load

    SliderAutoSpeed.Value = 50
    LableAutoSpeed.Caption = SliderAutoSpeed.Value
    
    
    SliderManuSpeed.Value = 50
    LabelPercent.Caption = SliderManuSpeed.Value
    
    TextShelf.Text = 1
    TextShelfNumber.Text = 1
    
    txtWorkPiece(1).Text = ""
    
    ''optMode(2).Value = True
    
    ImgFdBkWPiece.Visible = False
    
    TxtViewPocket.Text = "101"
    TextLineNumber.Text = "1"
    
    ComboLocations.Text = "Kiosk"
    

End Sub

Private Sub DisplayTablePocketsLocations()
''
Debug.Print "DisplayTablePocketsLocations()"
Dim i As Integer


    TablePocketsLocations.ColWidth(0) = 1585
    TablePocketsLocations.ColWidth(1) = 1000
    TablePocketsLocations.ColWidth(2) = 1000
    TablePocketsLocations.ColWidth(3) = 1000
    
    
    TablePocketsLocations.Width = 8060
     
    TablePocketsLocations.TextMatrix(0, 0) = "Pocket number"
    TablePocketsLocations.TextMatrix(0, 1) = "X"
    TablePocketsLocations.TextMatrix(0, 2) = "Y"
    TablePocketsLocations.TextMatrix(0, 3) = "Z"
    TablePocketsLocations.TextMatrix(0, 4) = "Rx"
    TablePocketsLocations.TextMatrix(0, 5) = "Ry"
    TablePocketsLocations.TextMatrix(0, 6) = "Rz"

    'set the width of the columns
    For i = 0 To TablePocketsLocations.Cols - 1
        TablePocketsLocations.ColAlignment(i) = flexAlignCenterCenter
    Next
    
    
    For i = 1 To 10
        TablePocketsLocations.TextMatrix(i, 0) = i + 100
    Next

End Sub

Private Sub DisplayTableTeach()
''1.the function set parameters of the table like "
''  width,height,title,,,
Debug.Print "DisplayTableTeach()"
    
    TableTeach.TextMatrix(0, 1) = "X"
    TableTeach.TextMatrix(0, 2) = "Y"
    TableTeach.TextMatrix(0, 3) = "Z"
    
    TableTeach.TextMatrix(1, 0) = "First point"
    TableTeach.TextMatrix(2, 0) = "Last Point"
    
    TableTeach.ColWidth(0) = 1200
    TableTeach.ColWidth(1) = 900
    TableTeach.ColWidth(2) = 900
    TableTeach.ColWidth(3) = 900
    
    TableTeach.Width = 3950
    TableTeach.Height = 900
    TextShelfNum.Text = "1"
    
    TableTeach.ColAlignment(0) = flexAlignCenterCenter
    TableTeach.ColAlignment(1) = flexAlignCenterCenter
    TableTeach.ColAlignment(2) = flexAlignCenterCenter
    TableTeach.ColAlignment(3) = flexAlignCenterCenter
      
End Sub


Private Sub UpDnDiameter_Change()

    TextToolsDiameter.Text = UpDnDiameter.Value
    
End Sub



Private Sub UpDnLine_DownClick()
'this function operate the :
'WorkPiece->TableOptions->LineNumber
Dim TempInt As Integer

    TempInt = CInt(TextLineNumber.Text)
    If TempInt > 1 Then
        TempInt = TempInt - 1
        TextLineNumber.Text = CStr(TempInt)
    End If
    
    
    If (UpDnLine.Value = 1) Then
        cmdMove(6).Enabled = False
    ElseIf (UpDnLine.Value <> 1) Then
         cmdMove(6).Enabled = True
    End If
        
End Sub

Private Sub UpDnLine_UpClick()
'this function operate the :
'WorkPiece->TableOptions->LineNumber
Dim TempInt As Integer

    TempInt = CInt(TextLineNumber.Text)
    TempInt = TempInt + 1
    TextLineNumber.Text = CStr(TempInt)
    
    If (UpDnLine.Value = 1) Then
        cmdMove(6).Enabled = False
    ElseIf (UpDnLine.Value <> 1) Then
         cmdMove(6).Enabled = True
    End If
    
End Sub

Private Sub UpDnPocketIndex_Change()

    TxtSinglePocketNumber = UpDnShelfList.Value * 100 + UpDnPocketIndex.Value
    
End Sub

Private Sub UpDnShelfList_Change()
    
    TextShelf.Text = UpDnShelfList.Value
    Call UpDnPocketIndex_Change
    fMainForm.LabelShelfType = CStr(AppShelvs(1).ShelfName)
    Call BtnDisplayPocketStatus_Click
    
End Sub

Private Sub UpDnShelf_Change()
''1.this function change the text in the ShlefTextBox
''2.the function refresh the First-And-Last Pocket Table.
  
Dim ret As Integer
Dim shelf As Integer
On Error GoTo label

    shelf = CInt(UpDnShelf.Value)
    LabelShelfType2.Caption = AppShelvs(shelf).ShelfName
    
    TextShelfNum.Text = CStr(shelf)
    
    Call UpDownPocket_Change
    
    If AppShelvs(shelf).ShelfToolType = AppToolType Then
        fMainForm.BtnTeachFirst.Enabled = True
        fMainForm.BtnTeachLast.Enabled = True
        fMainForm.BtnCalcPoints.Enabled = True
        fMainForm.BtnTeachSingle.Enabled = True
    ElseIf AppShelvs(shelf).ShelfToolType <> AppToolType Then
        fMainForm.BtnTeachFirst.Enabled = False
        fMainForm.BtnTeachLast.Enabled = False
        fMainForm.BtnCalcPoints.Enabled = False
        fMainForm.BtnTeachSingle.Enabled = False
    End If
    
    Call BtnShowFirstLastPockets_Click   ''refresh the table of first and last pocket
    
Exit Sub
label:
   ret = MsgBox("The System was unable to Change Shelf number," & vbCrLf _
        & "Try to fix one of te Bellow :" & vbCrLf _
        & "1.Wrong Command." & vbCrLf _
        & "2.Communication problem with the Controller." & vbCrLf _
        & "   [MdinMain.VScroll2_Change]", vbInformation, "Error in Change Shelf number.")

End Sub

Private Sub UpDnViewPocket_Change()

    TxtViewPocket = UpDownViewShelf.Value * 100 + UpDnViewPocket.Value

    Call BtnViewPocketsLocations_Click

End Sub

Private Sub UpDown1_Change()
    TxtSingleToolDiameter.Text = UpDown1.Value
End Sub

Private Sub UpDown2_Change()

    txtTestPocket(2) = UpDownTestPocket(2).Value * 100 + UpDown2.Value
            
End Sub



Private Sub UpDown3_Change()
    TextToolsDiameter2 = UpDown3.Value
    TextAllPocketsDiameter = UpDown3.Value
    TextDrillCode = UpDown3.Value
End Sub

Private Sub UpDownMyDiameter_Change()
    TextToolsDiameter2 = UpDownMyDiameter.Value
    TextAllPocketsDiameter = UpDownMyDiameter.Value
    TextDrillCode = UpDownMyDiameter.Value
End Sub

Private Sub UpDownStationNumber_Change()

    TextStationNumber.Text = CStr(UpDownStationNumber.Value)

End Sub

Private Sub UpDownTestShelf1_Change()
   Call UpDownTestPocket_Change(0)
End Sub

Private Sub UpDownTestShelf2_Change()
   Call UpDownTestPocket_Change(1)
End Sub

Private Sub UpDownViewShelf_Change()

'1.the function derived from Teach-ViewShelfs
'2.the function change the Number in the textBox
'3.the function change the indexs in the Table

Dim i As Integer
Dim ret As Integer

On Error GoTo error

    'change the number in the text box
    TextShelfNumber.Text = CStr(UpDownViewShelf.Value)
    Call UpDnViewPocket_Change
    
    'refresh the table
    Call BtnViewPocketsLocations_Click
    
    Exit Sub
    
error:
    ret = MsgBox("Error while Display locations" & vbCrLf _
    & "the error is : " & Err.Description, _
    vbExclamation, "UpDownViewShelf_Change()")
End Sub



Public Sub Delay(PauseTime As Long, Optional FlgDoeventEnable As Boolean = True)
Dim EndPeriod As Long
    EndPeriod = timeGetTime + PauseTime
    Do While timeGetTime() <= EndPeriod
        If FlgDoeventEnable Then DoEvents   ' Yield to other processes.
    Loop
End Sub






Private Sub ClearTablePocketsLocations()
'this function run on  all  the table :
'View Shelves Loc.
'and put "" in each cell

Dim i As Integer
Dim j As Integer
Dim Rows As Integer

Rows = TablePocketsLocations.Rows

    For i = 1 To Rows - 1
        For j = 0 To 3
            TablePocketsLocations.TextMatrix(i, j) = ""
        Next
    Next
    
End Sub

Private Sub ClearMSFlexGrid2()
'this function run on  all  the table :
'"Pocket List "
'and put "" in each cell

Dim i As Integer
Dim j As Integer
Dim Rows As Integer

    Rows = MSFlexGrid2.Rows

    For i = 1 To Rows - 1
        For j = 0 To 4
            MSFlexGrid2.TextMatrix(i, j) = ""
        Next
    Next
    
End Sub

Private Sub UpDownPocket_Change()
    TextPocketNumber = UpDownPocket.Value
End Sub

Private Sub UpDown5_Change()
    TextToolsDiameter2 = UpDown5.Value
    TextAllPocketsDiameter = UpDown5.Value
    TextDrillCode = UpDown5.Value
End Sub

Private Sub UpDownLoadUnloadShelf_Change()
   txtLoadUnloadShelf.Text = UpDownLoadUnloadShelf.Value
   Call UpDownLoadUnloadPocket_Change
End Sub

Private Sub UpDownLoadUnloadPocket_Change()
    txtLoadUnloadPocket = UpDownLoadUnloadShelf.Value * 100 + UpDownLoadUnloadPocket.Value
End Sub
Private Sub UpDownTestShelf_Change()
   TextShelfNumber2.Text = UpDownTestShelf.Value
End Sub

Private Sub UpDownTestPocket_Change(Index As Integer)

    If Index = 0 Then
        txtTestPocket(Index) = UpDownTestShelf1.Value * 100 + UpDownTestPocket(Index).Value
    Else
        txtTestPocket(Index) = UpDownTestShelf2.Value * 100 + UpDownTestPocket(Index).Value
    End If
    
    If Index = 2 Then
        txtTestPocket(2) = UpDownTestPocket(2).Value * 100 + UpDown2.Value
    End If
    
End Sub


Public Sub ClearIncrementalTarget()

Dim j As Integer

        For j = 0 To 5
            IncrementalTarget(j) = 0
        Next

End Sub


Private Sub CheckLoopAllPockets_Click()

Dim ret As Integer

    ''ret = MsgBox("the value is: " & CheckLoopAllPockets.Value _
        , vbOKOnly + vbInformation, "CheckLoopAllPockets_Click()")
    If CheckLoopAllPockets.Value = vbUnchecked Then
        Call fMainForm.BtnStopLoadUnload_Click
    End If
    

End Sub

Private Sub CheckLoopP2P_Click()

Dim ret As Integer

    ''ret = MsgBox("the value is: " & CheckLoopP2P.Value _
        , vbOKOnly + vbInformation, "CheckLoopP2P_Click()")
    If CheckLoopP2P.Value = vbUnchecked Then
        Call fMainForm.BtnStopLoadUnload_Click
    End If
    
End Sub






