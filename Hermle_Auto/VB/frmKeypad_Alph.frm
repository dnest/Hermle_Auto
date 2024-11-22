VERSION 5.00
Begin VB.Form frmKeypadAlph 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6804
   ClientLeft      =   48
   ClientTop       =   48
   ClientWidth     =   10920
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   567
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   15
      TabIndex        =   52
      Top             =   -75
      Width           =   10890
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         Caption         =   "KeyPad"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   180
         TabIndex        =   53
         Top             =   180
         Width           =   10515
      End
   End
   Begin VB.Frame fraMain 
      Height          =   6195
      Left            =   15
      TabIndex        =   0
      Top             =   600
      Width           =   10890
      Begin VB.CheckBox chkNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   13
         Left            =   180
         Picture         =   "frmKeypad_Alph.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   5160
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   12
         Left            =   1140
         Picture         =   "frmKeypad_Alph.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   5160
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   61
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   49
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "_"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   42
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   5160
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   ")"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   44
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   5160
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "("
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   43
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   46
         Top             =   5160
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "Space"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   10
         Left            =   4800
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   5160
         Width           =   2475
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "Z"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   34
         Left            =   1740
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   35
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   36
         Left            =   3660
         Style           =   1  'Graphical
         TabIndex        =   42
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "V"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   37
         Left            =   4620
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   38
         Left            =   5580
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   39
         Left            =   6540
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "M"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   40
         Left            =   7500
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   4260
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "L"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   33
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "G"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   29
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   36
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "H"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   30
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "J"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   31
         Left            =   6900
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "K"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   32
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "O"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   22
         Left            =   8340
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   23
         Left            =   9300
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "A"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   25
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   30
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   26
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   29
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "D"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   27
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   28
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "Q"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   14
         Left            =   660
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "W"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   15
         Left            =   1620
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "E"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   16
         Left            =   2580
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "R"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   17
         Left            =   3540
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "T"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   18
         Left            =   4500
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "Y"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   19
         Left            =   5460
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "U"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   20
         Left            =   6420
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   21
         Left            =   7380
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2460
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   64
         Left            =   9420
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   5160
         Width           =   1275
      End
      Begin VB.CheckBox chkNo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.6
            Charset         =   177
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   62
         Left            =   9780
         Picture         =   "frmKeypad_Alph.frx":1194
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   41
         Left            =   8460
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   4260
         Width           =   915
      End
      Begin VB.TextBox txtEditor 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   36
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   945
         Left            =   2280
         TabIndex        =   15
         Top             =   360
         Width           =   7335
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   7
         Left            =   5940
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   8
         Left            =   6900
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   9
         Left            =   7860
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   11
         Left            =   9780
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   4
         Left            =   3060
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   5
         Left            =   4020
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   6
         Left            =   4980
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkShift 
         Caption         =   "Shift"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   4260
         Width           =   1335
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   3
         Left            =   2100
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   2
         Left            =   1140
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   1
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1560
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "CR"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   1755
         Index           =   63
         Left            =   9780
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3360
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   " Del All"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   60
         Left            =   180
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   420
         Width           =   915
      End
      Begin VB.CheckBox chkNo 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   855
         Index           =   0
         Left            =   8820
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1560
         Width           =   915
      End
   End
End
Attribute VB_Name = "frmKeypadAlph"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Public _
'psTitle As String, psValue As String
'
''Private _
''msValue As String

''psTitle:is the string to be shown in the Bunner in the head of the form.
''psValue:is the string to be shown in the main window of the form.

Private Sub Form_Load()
'   If psTitle <> vbNullString Then lblTitle.Caption = psTitle
'   txtEditor = psValue
End Sub


Public Sub Form_Activate()

   txtEditor.SetFocus
   txtEditor.Text = psValue
   lblTitle.Caption = psTitle
   txtEditor.SelStart = Len(txtEditor)
   txtEditor = psValue
   
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyEscape Then
      Me.Hide
   ElseIf KeyCode = vbKeyReturn Then
      Me.Hide
      psValue = txtEditor
   End If
End Sub


Private Sub chkNo_Click(Index As Integer)
   Dim _
   sTemp1 As String, sTemp2 As String, iNumCh As Integer, _
   sTemp3 As String
'dd
   If chkNo(Index) = vbChecked Then
     chkNo(Index) = vbUnchecked:     Exit Sub
   End If
   
   With txtEditor
   Select Case Index
   Case Is < 10
      sTemp1 = Left(.Text, .SelStart)
      sTemp2 = Right(.Text, Len(.Text) - .SelStart)
      iNumCh = .SelStart
      .Text = sTemp1 & chkNo(Index).Caption & sTemp2
      .SelStart = iNumCh + 1
   Case 10
      sTemp1 = Left(.Text, .SelStart)
      sTemp2 = Right(.Text, Len(.Text) - .SelStart)
      iNumCh = .SelStart
      .Text = sTemp1 & " " & sTemp2
      .SelStart = iNumCh + 1
   Case 13
      If .SelStart <> 0 Then _
         .SelStart = .SelStart - 1
   Case 13
      If .SelStart <> Len(.Text) Then _
         .SelStart = .SelStart + 1
   Case Is < 46
      sTemp3 = chkNo(Index).Caption
      If chkShift.Value = 0 Then
         sTemp3 = chkNo(Index).Caption
      Else
         sTemp3 = LCase(chkNo(Index).Caption)
      End If
      
      sTemp1 = Left(.Text, .SelStart)
      sTemp2 = Right(.Text, Len(.Text) - .SelStart)
      iNumCh = .SelStart
      .Text = sTemp1 & sTemp3 & sTemp2
      .SelStart = iNumCh + 1
   Case 60
       .Text = ""
   Case 61
      sTemp1 = Left(.Text, .SelStart)
      iNumCh = Len(.Text) - .SelStart - 1
      If iNumCh < 0 Then iNumCh = 0
      sTemp2 = Right(.Text, iNumCh)
      iNumCh = .SelStart
      If iNumCh < 0 Then iNumCh = 0
      .Text = sTemp1 & sTemp2
      .SelStart = iNumCh
   Case 62
      iNumCh = .SelStart - 1
      If iNumCh < 0 Then iNumCh = 0
      sTemp1 = Left(.Text, iNumCh)
      sTemp2 = Right(.Text, Len(.Text) - .SelStart)
      iNumCh = .SelStart - 1
      If iNumCh < 0 Then iNumCh = 0
      .Text = sTemp1 & sTemp2
      .SelStart = iNumCh
   End Select
   .SetFocus
   
   Select Case Index
   Case 63
   
      Me.Hide
      
      If CurrentTBox = 1 Then
        fMainForm.txtWorkPiece(msValue).Text = txtEditor.Text
        
      ElseIf CurrentTBox = 2 Then
        fMainForm.textChangeWorkPiece.Text = txtEditor.Text
        
      ElseIf CurrentTBox = 3 Then
        fMainForm.TxtSinglePocketNumber.Text = txtEditor.Text
        
      ElseIf CurrentTBox = 4 Then
         fMainForm.TxtSingleToolDiameter.Text = txtEditor.Text
         
      ElseIf CurrentTBox = 5 Then
        fMainForm.TextPocketNumber.Text = txtEditor.Text
      
      ElseIf CurrentTBox = 6 Then
        fMainForm.TextToolsDiameter.Text = txtEditor.Text
      
      End If
      
      
   Case 64
      Me.Hide
      
   End Select
   End With
End Sub

