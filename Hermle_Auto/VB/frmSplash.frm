VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7770
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   10500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7770
   ScaleWidth      =   10500
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BorderStyle     =   0  'None
      Height          =   7500
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   10245
      Begin VB.PictureBox picLogo 
         Height          =   2100
         Left            =   36
         ScaleHeight     =   2040
         ScaleWidth      =   2040
         TabIndex        =   1
         Top             =   180
         Width           =   2100
         Begin VB.Image Image1 
            Height          =   1995
            Left            =   0
            Picture         =   "frmSplash.frx":0000
            Stretch         =   -1  'True
            Top             =   0
            Width           =   1995
         End
      End
      Begin VB.Label lblVersion 
         AutoSize        =   -1  'True
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4185
         TabIndex        =   4
         Tag             =   "CompanyProduct"
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Image Image2 
         Height          =   4440
         Left            =   2232
         Picture         =   "frmSplash.frx":25A0
         Stretch         =   -1  'True
         Top             =   2916
         Width           =   5748
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Automation for Hermle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   432
         Left            =   3516
         TabIndex        =   3
         Tag             =   "CompanyProduct"
         Top             =   1116
         Width           =   3900
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         Caption         =   "Shafir Production System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   3195
         TabIndex        =   2
         Tag             =   "CompanyProduct"
         Top             =   405
         Width           =   4305
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    lblVersion.Caption = "Version :" & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Image1_Click()

    Me.Hide
    Unload (Me)
    
End Sub


