VERSION 5.00
Begin VB.Form FormDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IS2903 hermle"
   ClientHeight    =   5925
   ClientLeft      =   8640
   ClientTop       =   5100
   ClientWidth     =   7080
   Icon            =   "Dialog.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5925
   ScaleWidth      =   7080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton YesButton 
      Caption         =   "YES"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   2340
      TabIndex        =   8
      Top             =   5200
      Width           =   1400
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "NO"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   3915
      TabIndex        =   1
      Top             =   5200
      Width           =   1400
   End
   Begin VB.CommandButton BtnExit 
      Caption         =   "CLOSE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   177
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   550
      Left            =   5535
      TabIndex        =   0
      Top             =   5200
      Width           =   1400
   End
   Begin VB.Label ExtraText 
      Caption         =   "ExtraText"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2340
      TabIndex        =   7
      Top             =   3960
      Width           =   4575
   End
   Begin VB.Label TextSolution 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TextSolution"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   2340
      TabIndex        =   6
      Top             =   2760
      Width           =   4575
   End
   Begin VB.Label LabelHeader 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabelHeader"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   645
      Left            =   720
      TabIndex        =   5
      Top             =   405
      Width           =   6225
   End
   Begin VB.Label LabelFunction 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabelFunction"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4095
      TabIndex        =   4
      Top             =   4560
      Width           =   2835
   End
   Begin VB.Label LabelModule 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabelModule"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2340
      TabIndex        =   3
      Top             =   4560
      Width           =   1635
   End
   Begin VB.Line Line1 
      X1              =   720
      X2              =   6960
      Y1              =   1260
      Y2              =   1260
   End
   Begin VB.Image Image1 
      Height          =   1170
      Left            =   585
      Picture         =   "Dialog.frx":030A
      Stretch         =   -1  'True
      Top             =   1890
      Width           =   1215
   End
   Begin VB.Label TextMain 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TextMain"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   2340
      TabIndex        =   2
      Top             =   1530
      Width           =   4575
   End
End
Attribute VB_Name = "FormDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim header(70) As String
Dim maintext(70) As String
Dim solution(70) As String
Dim ReturnValue As Integer



Private Sub BtnExit_Click()

    ReturnValue = Fdbk.UserExit
    Me.Hide
    Unload Me
    
End Sub

Private Sub CancelButton_Click()

    ReturnValue = Fdbk.UserNO
    Me.Hide
    Unload Me
    
End Sub



Private Sub Form_Load()

    Call DefineHeader
    Call DefineMainText
    Call DefineSolution

End Sub


Private Sub DefineHeader()

    header(11) = "Pocket to poket"
    header(12) = "all pockets"
    header(13) = "From Chuck to Pocket"
    header(14) = "Pocket To Kiosk"
    header(15) = "Pocket To Chuck"
    header(16) = "Chuck To Pocket"
    header(17) = "Kiosk To Chuck"
    header(18) = "Chuck To Kiosk"
    header(19) = "Test Cycle"
    header(21) = "Delete Line"
    header(22) = "AddWorkPiece"
    header(23) = "Change Priority"
    header(30) = "Start Application"
    header(31) = "Start Automat"
    header(32) = "Change Status for Multiply Pockets "
    header(33) = "Start robot job "
    header(34) = "Communication Test"
    header(35) = "Exit Application"
    header(36) = "change all status"
    header(37) = "calculate points"
    header(38) = "Place tool From Kiosk"
    header(39) = "Teach Position"
    header(40) = "tool counter in WorkPiece"
    header(41) = "find empty pocket"
    header(42) = "find a full pocket to unload in the first time"
    header(43) = "find an machined pocket to unload"
    header(44) = "Start automat cycle"
    header(45) = "Start Application"
    header(46) = "Change HSK Diameter"
    header(47) = "Change Drill Diameter"
    header(48) = "Look For pocket to Take to chuck."
    header(49) = "take tool to chuck"
    header(50) = "read data from robot"
    header(51) = "Reset Profibus communication"
    header(52) = "from pocket to kiosk"
    header(53) = "Wrong WorkPiece"
    header(54) = "Input/Output"
    header(55) = "Start Automatic Cycle"
    header(56) = "kiosk to pocket"
    header(57) = "Emergency Stop Release"
    header(58) = "Shelvs configuration"
    header(59) = "teach First pocket"
    header(60) = "Start Automatic cycle"
    header(61) = "Set Tool Sensor state"
    header(62) = "Offset Parameters"
    header(63) = "Adjustment Cycle"
    header(64) = "Change robot velocity"
    header(65) = "Tech Code"
    header(66) = "Change RoundTool Diameter"
    header(67) = "Unload From Pocket to kiosk."
    
    
End Sub

Private Sub DefineMainText()
    
    maintext(11) = "The System was not able to start the movement. "
    maintext(12) = "The System was not able to send data to Robot. "
    maintext(13) = "The System was not able to read data from robot. "
    maintext(14) = "Can not add WorkPiece."
    maintext(15) = "Can not Delete WorkPiece. "
    maintext(16) = "The System was not able to read data from robot. "
    maintext(17) = "can not find Empty pocket in this work piece."
    
    
    ''wrong input
    maintext(20) = "the input should not be empty. "
    maintext(21) = "the input should be a number. "
    
    maintext(22) = "The System was not able to read data from robot. "
    maintext(23) = "Can not add WorkPiece. "
    maintext(24) = "Can not Delete WorkPiece. "
    maintext(25) = "The System was not able to read data from robot. "
    maintext(26) = "The System was unable to change priority"
    maintext(30) = "can not connect to controller"
    maintext(31) = "can not change status.WorkPiece can not be empty"
    maintext(32) = "The system was not able to start the Robot Job"
    maintext(34) = "the system was unable to connect to the robot"
    maintext(35) = "you are about to exit the application"
    maintext(36) = "the system was not able to change all status"
    maintext(37) = "The System was unable to Calculate Points"
    maintext(38) = "the system can not send position to the controller."
    maintext(39) = "can not teach position.Input not correct."
    maintext(40) = "can not read tool counter from Controller."
    maintext(41) = "the pocket is not empty."
    maintext(42) = "the pocket is Wrong.the pocket should be machined or unmachined only."
    maintext(43) = "the pocket isn't correct.the pocket should be machined only."
    maintext(44) = "there is no workpiece.can not start cycle."
    maintext(45) = "The application already running."
    maintext(46) = "The Diameter is incorrect.Should be 100/200/300 For HSK"
    maintext(47) = "The Diameter is incorrect.Should be 1-7 for Drill "
    maintext(48) = "the system can not find a pocket to take from."
    maintext(49) = "the system can not send pocket number to the controller"
    maintext(50) = "data was read from controller to the computer"
    maintext(51) = "Can not reset profibus communication.communication might be lost or wrong file name"
    maintext(53) = "the workpiece number is incorrect"
    maintext(54) = "can not send command to controller.the key is not in Remote Mode"
    maintext(55) = "can not start cycle.gripper is not open."
    maintext(56) = "the pocket is wrong.pocket can not contain Zero value"
    maintext(58) = "the shelvs configuration is wrong."
    maintext(59) = "can not teach pocket.communication error or shelf not enable."
    maintext(60) = "can not write drill code to robot.check communication error."
    maintext(61) = "can not write tool sensor state.check communication error."
    maintext(62) = "can not send parameters to robot."
    maintext(63) = "make sure doors and chuck are open."
    maintext(64) = "can not change robot speed.might be Communication problem or key position."
    maintext(65) = "tech code is missing or not correct."
    maintext(66) = "the New Diameter is incorrect.should be 1 to 8."
    maintext(67) = "can not find machined pocket in this Work Piece."

End Sub

Private Sub DefineSolution()
    
    solution(11) = "there is no communication with Robot. "
    solution(12) = "the key is not in Remote Mode. "
    solution(13) = "teach position before calculation. "
    solution(14) = "not all fields are full. "
    solution(15) = "Insert a correct number. "
    solution(16) = "make sure all fields are full,or diameter is incorrect "
    solution(17) = "Input should be a number "
    solution(18) = "check pocket status,tool type or workpiece number."
    
    solution(20) = "please fill all inputs. "
    solution(21) = "please insert a only numbers. "
    
    solution(22) = "The System was not able to read data from robot. "
    solution(23) = "Can not add WorkPiece. "
    solution(24) = "Can not Delete WorkPiece. "
    solution(25) = "The System was not able to read data from robot. "
    solution(26) = "Emergemcy Stop might be pressed or door/window is open."
    solution(30) = "Check Controller On/Off and IP"
    solution(31) = "Make sure all inputs are full"
    solution(32) = "job name does not exist.Or Job Already running"
    solution(34) = "Check Motoman,Beckhoff or cable"
    solution(35) = "are you sure ?"
    solution(36) = "check inputs and robot communication"
    solution(37) = "Inputs should not be empty.or shelf is not enable. "
    solution(38) = "try to fix communication problem with robot.controller might be in fault or shut down "
    solution(39) = "please select position"
    solution(40) = "communication poblem or key not in REMOTE."
    solution(41) = "please select a different pocket."
    solution(42) = "make sure the pocket is full."
    solution(43) = "make sure the pocket is full and the tool is machined."
    solution(44) = "Insert a WorkPiece First."
    solution(45) = "close the current application."
    solution(46) = "Please enter correct diameter 100,200,300"
    solution(47) = "Please enter correct diameter (1-7)"
    solution(48) = "please fix the pocket status and press RESUME"
    solution(49) = "check communication with the robot.fault,alarm,power or IP."
    solution(50) = "the data was saved in the hard disk."
    solution(51) = "check controller for error,alarm or comm problem"
    solution(53) = "please select correct number"
    solution(54) = "please turn the key to Remote Mode"
    solution(55) = "Please Switch to Manual and Open Gripper."
    solution(56) = "check pocket status and workpiece."
    solution(58) = "Please select correct Tool Type."
    solution(59) = "check communication problem and shelf configuration."
    solution(60) = "communication problem with robot or Controller alarm."
    solution(61) = "communication error with robot or Controller alarm."
    solution(62) = "input can not be empty.input have to be numbers."
    solution(63) = "doors and chuck are open ?"
    solution(64) = "check for communication error with robot,Controller alarm,or key position."
    solution(65) = "please insert correct Tech Code."
    solution(66) = "Enter a correct value between 1 to 8."
    solution(66) = "check pocket status and shelf agian."
    
End Sub

Public Function ShowDialogForm _
        (ByVal Myheader As Integer, _
            ByVal MyTextMain, _
                ByVal MySolution As Integer, _
                    ByVal myModule As String, _
                        myCurrentFunction As String, _
                            myicon As DialogIcon) As Integer
    
    FormDialog.LabelHeader.Caption = header(Myheader) ''need once !  do not erase !
    FormDialog.LabelHeader.Caption = header(Myheader) ''need twice!  do not erase!
    FormDialog.TextMain.Caption = maintext(MyTextMain)
    TextSolution.Caption = solution(MySolution)
    ExtraText.Caption = ""
    LabelModule.Caption = myModule
    LabelFunction.Caption = myCurrentFunction
    
    If myicon = EmergencyStop Then
     Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\Estop_48.ico")
    
    ElseIf myicon = ServoPower Then
    Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\reset50.ico")
    
    ElseIf myicon = KeyPosition Then
    Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\key128.jpg")
    
    ElseIf myicon = InputIncomplete Then
    Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\X_48.ico")
    
    ElseIf myicon = GeneralError Then
    Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\X_48.ico")
    
    ElseIf myicon = ExitApp Then
    Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\ExitDoor.ico")
    
    ElseIf myicon = ExitApp Then
    Image1.Picture = LoadPicture(App.path & "\WorkDirectory\icons\V_48.ico")
    
    End If
    
    FormDialog.Show '''(vbModal)

    ShowDialogForm = ReturnValue

End Function


Private Sub YesButton_Click()

    ReturnValue = Fdbk.UserYes
    Me.Hide
    Unload Me

End Sub
