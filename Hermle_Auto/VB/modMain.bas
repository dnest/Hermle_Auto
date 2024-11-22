Attribute VB_Name = "modMain"
Option Explicit
Public Declare Function timeGetTime Lib "winmm.dll" () As Long
Public fMainForm As mdiMain
Public psTitle  As String
Public psValue  As String
Public msValue As Integer
Public PasswordOk As Boolean
Public DrillCode As Integer
Public LogErrFileNumber As Integer
Public LogErrFileName As String
Public WorkDir As String
Public Velocity As Integer
Public LoadingMessage As String
Public loadingStart As String
Public UnLoadingMessage As String
Public UnloadingStart As Boolean

'for AutoMode
Public iAutoTool As Boolean
Public iAutoNight As Boolean

'for virtual KeyBoard
Public CurrentTBox As Integer

Public ToolCounter  As Integer

''for the communication with the robot
Public m_nCid As Integer ''             number comm ID
Public open_mode As Integer ''          type of openning comm

''public arrays for data exchange with the robot---
Public AllJobFiles(50) As String
Public ArrayDouble(10) As Double
Public ArrayInteger(10) As Double
Public ArrayByte(10) As Double
Public ArrayPosition(10) As Double
Public ArrayJobs(50) As String
Public TempPosition(10) As Double
Public IncrementalTarget(5) As Double
Public ArrayString As String * 17
Public ArrayMessage(10) As Double

''---public Counters ---
Public DisplayTimerCounter As Integer
Public AppCurrentWPCounter As Integer
Public AppWPIndex As Double

''---public integers----
Public AppToolsWPiece As Integer ''for Load\Unload Only
Public GripperStatus As Integer

''---public strings---
Public CurrentTest As String
Public AppCurrentPocket As String
Public StringLog(100) As String
Public RobotAlarmString As String
Public RobotStrings(25) As String

''---public numeric consts---
Public Const TotalHSK = 10
Public Const TotalDRILL = 12
Public Const TotalROUND = 12
Public Const PI = 3.141592
Public Const AmountOfGeneralLocations = 127
Public ShelfOffset(3) As Double

''--public parameters------
Public AppDiameter As Integer
Public AppSpeed As Integer
Public AppGripperStyle As Integer

''--geometric const ------
Public PocketDepth As Double
Public ChuckDepth As Double
Public ChuckStopper As Double
Public PocketStopper As Double
Public ShelfSafty As Double
Public KioskStopper As Double
Public AboveChuck As Double
Public AbovePocket As Double


''--public boolian flags---
Public RunTestEndless As Boolean
Public LoadFirstTime As Boolean
Public UnloadFirstTime As Boolean
Public FlagInfo As Boolean
Public FlagDiagnostic As Boolean
Public AppSimulation As Boolean
Public UseToolSensor As Boolean
Public OffsetSent As Boolean
Public CycleDone As Boolean
Public UseExternalFile As Boolean
Public UseHMILogger As Boolean
Public UseHMIInfo As Boolean
Public ChuckUnloadFirstTime As Boolean
Public UseSaw As Boolean


''public enum ---
Public Enum AppGear
    Manual = 1
    Semi = 2
    Automat = 3
End Enum

Public Enum AppCommState
    OffLine = 0
    ReadOnly = 1
    OnLine = 2
End Enum

Public Enum fdbk
    UserExit = -1
    UserNO = 0
    UserYes = 1
End Enum


Public Enum IOinputsAdd

    kioskdooropen = 10
    KioskDoorClose = 11
    UserAck = 12
    ValvesClose = 13
    HolderDirectRobot = 14
    ToolFound = 15
    GripperIsOpen = 16
    Gripper2IsOpen = 20
    Spare_1 = 17
    Spare_2 = 20
    Estop_1 = 21
    Estop_2 = 22
    UserInterlock_1 = 23
    UserInterlock_2 = 24
    PCInterlock_1 = 25
    PCInterlock_2 = 26
    Spare_3 = 27
    
End Enum

Public Enum IOoutputsAdd
    KioskValveOpen = 10010
    KioskValveClose = 10011
    IndUserAck = 10012
    OpenGripper = 10013
    CloseGripper = 10014
    CabinetLight = 10015
    DoorInterLock = 10016
    spare = 10017
    IntelockhermleByPass = 10020
    Saftyrelay = 10021
End Enum


Public Enum AppWPMode
    WorkPiece = 2
    OneTool = 1
    NightMode = 3
End Enum

Public Enum tooltype
    other = 0
    HSK = 1
    Drill = 2
    Round = 3
End Enum

Public Enum DialogIcon
    EmergencyStop
    ServoPower
    KeyPosition
    InputIncomplete
    GeneralError
    ExitApp
    GreenV
End Enum

Public Enum DataType
    ByteType = 0
    IntegerType = 1
    DoubleType = 2
    Real = 3
    Position = 4
    BasePosition = 5
    StationPosition = 6
End Enum

Public Enum PocketStatus
    empty_ = 1     ''     no tool in pocket
    unmachined = 2 ''     tool after GOOD proccess in pocket
    machined = 3
    reserved = 4   ''     pocket not good.can not be used.
    Mask = 5       ''     the pocket's tooltype is different then the Apptooltype.
    occupied = 6 ''       pocket currently/temporarly can not be used.
    Broken = 7 ''         tool is being proccessed in the machined
    Disable = 8 ''        the pocket is not as the type of the gripper
End Enum

Public Enum JobState
    IDLE = 1
    RUN = 2
    Done = 3
End Enum

Public Enum KeyState
    error_
    teach
    Play
    remote
End Enum

''public types---
Public Type AllJobs
    AutoStart(10) As Double
    Parking(10) As Double
    Exchange(10) As Double
    KioskToPocket(10) As Double
    PocketToKiosk(10) As Double
    PocketToChuck(10) As Double
    ChuckToPocket(10) As Double
    dummy(10) As Double
    Loading(10) As Double
    Unloading(10) As Double
    TakeFromPocket(10) As Double
    PlaceOnPocket(10) As Double
    TakeFromKiosk(10) As Double
    PlaceOnKiosk(10) As Double
    TakeFromChuck(10) As Double
    PlaceOnChuck(10) As Double
    
End Type

Public Type AllTests
    PocketToPocket(10) As Double
    KioskToPocket(10) As Double
    PocketToKiosk(10) As Double
    PocketToChuck(10) As Double
    ChuckToPocket(10) As Double
    AllPockets(10) As Double
End Type

Public Type HandShakeByte
    RequestPlaceFromKiosk(10) As Double
    DonePlaceFromKiosk(10) As Double
    RequestTakeToKiosk(10) As Double
    DoneTakeToKiosk(10) As Double
    RequestPlaceFromChuck(10) As Double
    DonePlaceFromChuck(10) As Double
    RequestTakeToChuck(10) As Double
    DoneTakeToChuck(10) As Double
    RequestNewWorkPiece(10) As Double
    DoneNewWorkPiece(10) As Double
    RequestNCFinish(10) As Double
    MaskUser(10) As Double
    RequestChangePocketStatus(10) As Double
    RequestDeleteWPiece(10) As Double
    DoneDeleteWPiece(10) As Double
    
End Type

Public Type AutomationInfo
    LocationCountry As String
    LocationFactory As String
    AutomationNumber As String
    AutomationName As String
    HermleNumber As String
    HermleType As String
End Type

Public Type PocketProperties
    name As String
    shelf As Integer
    column As Integer
    pocket As Integer
    diameter As Integer
    CurrentTool As tooltype
    Status As PocketStatus
    WorkPiece As Integer
    ProgramNumber As Double
End Type

Public Type shelf
    ShelfNumber As Integer
    ShelfToolType As tooltype
    ShelfStatus As PocketStatus
    NumOfColumns As Integer
    ShelfEnable As Boolean
    BeenTaught As Boolean
    ShelfName As String
    NumOfPockets As Integer
    DefaultPocket As Integer
    DefaultDiameter As Integer
End Type

Public Type RobotPosition
    name As String
    X As Double
    Y As Double
    z As Double
    Rx As Double
    Ry As Double
    Rz As Double
    Dist As Double
    Alfa As Double
    org As Double
    xx As Double
    xy As Double
End Type

Public Type WorkPiece
    WPNumber As Integer
    NCProgram As Double
    ToolDiameter As Integer
    ToolAmount As Integer
    ToolAmountLeft As Integer
    LineNumber As Integer
    WPToolType As String
    WPStatus As PocketStatus
End Type

Public Type DrillPocket
    diameter(7) As RobotPosition
End Type

Public Type RoundPocket
    diameter(8) As RobotPosition
End Type

Public HermleAutomation As AutomationInfo
Public Jobs As AllJobs
Public Tests As AllTests
Public HandShake As HandShakeByte
Public AppToolType As tooltype
Public AppKeyState As KeyState
Public HermleGear As AppGear
Public AppWPiecePointer As PocketProperties
Public AppCurrentWorkPiece As WorkPiece
Public HermleCommState As AppCommState
Public HermleWPMode As AppWPMode
''
''public arrays---
Public AppShelvs(3) As shelf
Public HSKLocations(3, 10) As RobotPosition
Public DrillLocations(3, 12) As DrillPocket
Public RoundLocations(3, 12) As RoundPocket

Public AutomationStatus() As PocketProperties

Public AllWP(50) As WorkPiece
Public GeneralLocation(AmountOfGeneralLocations) As RobotPosition
Public JigHeight As Double

Public T11 As Double, T12 As Double, T13 As Double
Public T21 As Double, T22 As Double, T23 As Double, Tstatus(20) As Double

    


'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'***************************** M A I N   S U B *********************************************
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!

Sub Main()
''
Debug.Print "Sub Main()"
''
Dim bb As Boolean
Dim ret As Integer

    JigHeight = 30

    WorkDir = App.path & "\WorkDirectory"
    Set fMainForm = New mdiMain

    ''if the Allplication is already running
    If App.PrevInstance Then
        Call FrmDialog.ShowDialogForm(45, 45, 45, "ModMain", "Sub Main()", GeneralError)
        Exit Sub
    End If
     
    ''display the Splash form.
    frmSplash.Show: frmSplash.Refresh
    
    ModCsvFile.ReadAllWorkPiece ''        read all work piece from file into memory.
    ModIni.ReadShelvsConfig ''            read the shelvs configuration.
    ModIni.ReadAppToolType ''             read the application tooltype.
    ModCsvFile.ReadGeneralLocation ''     read the general positions locations
    
    If AppToolType = HSK Then
        ReDim AutomationStatus(3, 10)
    ElseIf AppToolType = Drill Then
        ReDim AutomationStatus(3, 12)
    ElseIf AppToolType = Round Then
        ReDim AutomationStatus(3, 12)
    End If
    
    ModCsvFile.ReadAutomationStatus '     read the status of the pockets.
    
    If AppToolType = HSK Then       '     load the pockets locations.
        ModCsvFile.ReadPocketsLocations ("HSKLocations")
    ElseIf AppToolType = Drill Then
        ModCsvFile.ReadPocketsLocations ("DrillLocations")
    ElseIf AppToolType = Round Then
        ModCsvFile.ReadPocketsLocations ("RoundLocations")
    End If
    
    Call uLogsInit
    bb = ModIni.ReadGripperStyle
    bb = ModIni.ReadSimulatorState
    Call ModIni.ReadUseExternalFile
    Call ModIni.ReadUseHMILogger
    Call ModIni.ReadUseHMIInfo
    Call ModIni.ReadSawState
    
    Load FrmCommunication
    Load fMainForm
    Load frmKeypadAlph
    Load FrmDialog
    Load frmOptions
    
    ModLogFile.LogAddLine ("start application:" & CStr(Format(now, "hh:mm:ss")))
    Call ModGUI.ListHMIUpdate(" start application:" & CStr(Format(now, "hh:mm:ss")))
    
    Call ModIni.ReadShelvsOffset
    Call ModIni.ReadToolSensorState
    Call ModHandShake.ResetAllCommBytes '' reset all the request from the robot.
    Call ModHandShake.ResetAllRobotRequest
    Call ModHandShake.ResetAllProccess ''   reset JOBS and Tests.
    
   'for the:"add workpiece"
    AppWPiecePointer.shelf = 1
    AppWPiecePointer.column = 1
  
    fMainForm.Show

    frmSplash.Hide
    Unload frmSplash
     
    AppWPIndex = 1
    OffsetSent = False

  Exit Sub
  
MainSubErr:
    ret = MsgBox("Error while loading application." & vbCrLf _
    & "the error is : " & Err.Description _
    , vbInformation _
    , App.Title)
End Sub
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
'*******************************************************************************************
'!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!


Private Sub uLogsInit()

''    'Call CreateDir(WorkDir & "LogFiles")
''    'Call checkSizeAndRemoveLogfiles
''
''    LogErrFileNumber = FreeFile
''    LogErrFileName = WorkDir & "LogFiles\Error_" & Day(now) & "_" & Month(now) & "_" & Year(now) & ".csv"
''    Open LogErrFileName For Append Access Write Lock Write As LogErrFileNumber
''    Write #LogErrFileNumber,
''    Write #LogErrFileNumber, Format(now, "dd/mm/yy hh:mm:ss"), "Start Program                    "

End Sub

Public Sub GlobalErr(str$)
On Error Resume Next
    If str$ <> "" Then
        Write #LogErrFileNumber, Format(now, "dd/mm/yy hh:mm:ss"), "Error = " & str$
    End If
'    With fMainForm
'        If .cmbWarnings.ListCount > 32766 Then .cmbWarnings.Clear
'        .cmbWarnings.AddItem str$ & " " & Time
'        .cmbWarnings.ListIndex = .cmbWarnings.ListCount - 1
'        .tmrResetshpClrWarnings.Enabled = True
'     End With
'     If onAutoCycle And Not FlagPauseAll Then flgRuntimeErr = True: BaryErrorCode = 4
End Sub

