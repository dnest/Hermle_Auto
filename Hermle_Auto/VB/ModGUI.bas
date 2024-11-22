Attribute VB_Name = "ModGUI"
Option Explicit
  
  
Public Sub DisplayTCPPosition()

Dim TCPLocation(15) As Double
Dim i As Integer

    Call ReadCurrentPosition(TCPLocation())

    For i = 0 To 5
       fMainForm.pnlPositionData(i) = TCPLocation(i)
    Next

End Sub
      
Public Sub DisplayServoStatus()

Dim servo As Boolean
On Error GoTo label

    servo = ReadServoStatus()
    
    If servo = True Then
        fMainForm.pnlServoStatus(0).ForeColor = RGB(0, 175, 0)
        fMainForm.pnlServoStatus(0).Caption = "ON"
    Else
        fMainForm.pnlServoStatus(0).ForeColor = RGB(255, 0, 0)
        fMainForm.pnlServoStatus(0).Caption = "OFF"
    End If
    
    Exit Sub
    
label:

    MsgBox "can not display servo state ; " & _
    Err.Description, vbCritical _
    , "DisplayServoStatus()"
    
End Sub

Public Sub DisplayLineNumber()

Dim MyLineNumber As Integer

    MyLineNumber = MotoComToolBox.ReadLineNumber()
    fMainForm.txtCycleStep.Caption = CStr(MyLineNumber)
 
End Sub

Public Sub DisplayJobName()

Dim MyJobName As String

    MyJobName = ReadJobName()
    fMainForm.txtCycleName.Caption = MyJobName

End Sub

Public Sub DisplayKeyState()

Dim KeyState As Integer

        
        Call ReadKeyState ''read the position of key in the TP.
        
        If AppKeyState = error_ Then
            fMainForm.pnlServoLoop(1).ForeColor = RGB(255, 0, 0)
            fMainForm.pnlServoLoop(1).Caption = "Error"
        ElseIf AppKeyState = teach Then
            fMainForm.pnlServoLoop(1).ForeColor = RGB(255, 0, 0)
            fMainForm.pnlServoLoop(1).Caption = "Teach"
        ElseIf AppKeyState = Play Then
            fMainForm.pnlServoLoop(1).ForeColor = RGB(255, 0, 0)
            fMainForm.pnlServoLoop(1).Caption = "Play"
        ElseIf AppKeyState = remote Then
            fMainForm.pnlServoLoop(1).ForeColor = RGB(0, 175, 0)
            fMainForm.pnlServoLoop(1).Caption = "Remote"
        End If
        
     
End Sub

Public Sub DisplayGripperState()

Dim ret As Integer
''Dim GripperStatus As Integer
Dim TempInteger As Integer


    GripperStatus = ReadIO(IOinputsAdd.GripperIsOpen)
     
    If GripperStatus <> 0 Then
        fMainForm.pnlServoLoop(2).Caption = "OPEN"
        
    ElseIf GripperStatus = 0 Then
        fMainForm.pnlServoLoop(2).Caption = "CLOSE"
        
    End If
    
''-1 : Header number error
''0:  Normal completion
''Others:  Error codes
    
End Sub



Public Sub DisplayRobotAlarm()
Dim ret As Integer

    ret = BscGetError2(m_nCid)
    If ret > 0 Then
        MsgBox CStr(ret)
    End If
    
'-1    : Acquisition Failure
' 0    :  No error
'Others:  Error codes
    
End Sub

Public Sub UpdateGUIByJobsStatus()
''1 .this function update the GUI according to the
''   statuses of all the jobs in the controller.
''   the function is being called after the  :
''   "ReadAllJobsStatus" function.
''2.the function is called from the main timer.
''3.the func get no parameters and return no parameters.


''1 idle
''2 run
''3 done

    If Jobs.KioskToPocket(0) = 1 Then
'        fMainForm.imgGrLampLoad.Visible = True
'        fMainForm.imgRedLampLoad.Visible = False
        fMainForm.LabelLoad.BackColor = vbRed
        fMainForm.LabelLoad.Caption = "Stop"
    ElseIf Jobs.KioskToPocket(0) = 2 Then
        fMainForm.LabelLoad.BackColor = vbGreen
        fMainForm.LabelLoad.Caption = "Run"
    ElseIf Jobs.KioskToPocket(0) = 3 Then
        fMainForm.LabelLoad.BackColor = vbRed
        fMainForm.LabelLoad.Caption = "Stop"
    Else
        fMainForm.LabelLoad.BackColor = vbRed
        fMainForm.LabelLoad.Caption = "Stop"
    End If
    
    If Jobs.PocketToKiosk(0) = 1 Then
'        fMainForm.imgGrLampUnload.Visible = True
'        fMainForm.imgRedLampUnload.Visible = False
        fMainForm.labelUnload.BackColor = vbRed
        fMainForm.labelUnload.Caption = "Stop"
    ElseIf Jobs.PocketToKiosk(0) = 2 Then
        fMainForm.labelUnload.BackColor = vbGreen
        fMainForm.labelUnload.Caption = "Run"
    ElseIf Jobs.PocketToKiosk(0) = 3 Then
        fMainForm.labelUnload.BackColor = vbRed
        fMainForm.labelUnload.Caption = "Stop"
    Else
        fMainForm.labelUnload.BackColor = vbRed
        fMainForm.labelUnload.Caption = "Stop"
    End If
    
    If Jobs.AutoStart(0) = 2 Then
        fMainForm.CmdSysOperation(3).Enabled = False
    ElseIf HermleGear = Automat Then ' elisha 15-01-2014
        fMainForm.CmdSysOperation(3).Enabled = True
    End If
    
    If Jobs.AutoStart(0) = 3 Then
        ChuckUnloadFirstTime = True
    End If
    
    ' Enable/Disable load user buttons request
    If (HandShake.MaskUser(0) = 0) And _
        (HermleGear = Automat) And _
            (Jobs.PocketToKiosk(0) = 0 Or Jobs.PocketToKiosk(0) = 1) And _
            (Jobs.KioskToPocket(0) = 0 Or Jobs.KioskToPocket(0) = 1) _
            Then
            fMainForm.cmdLoadTool.Enabled = True
   Else
        fMainForm.cmdLoadTool.Enabled = False
    End If

    ' Enable/Disable unload user buttons request
    If (HandShake.MaskUser(0) = 0) And _
        (HermleGear = Automat) And _
            (Jobs.PocketToKiosk(0) = 0 Or Jobs.PocketToKiosk(0) = 1) And _
            (Jobs.KioskToPocket(0) = 0 Or Jobs.KioskToPocket(0) = 1) _
            Then
            fMainForm.cmdUnloadTool.Enabled = True
  Else
        fMainForm.cmdUnloadTool.Enabled = False
    End If
    
End Sub


Public Sub LoadComboAllStatus()
''
Debug.Print "LoadComboAllStatus()"

    fMainForm.ComboAllStatus.AddItem "Empty     "
    fMainForm.ComboAllStatus.AddItem "Unmachined"
    fMainForm.ComboAllStatus.AddItem "Machined  "
    fMainForm.ComboAllStatus.AddItem "Reserved  "
    fMainForm.ComboAllStatus.AddItem "Mask      "
    fMainForm.ComboAllStatus.AddItem "Occupied  "
    fMainForm.ComboAllStatus.AddItem "Broken Tool"
    fMainForm.ComboAllStatus.AddItem "Disable   "
   
End Sub

Public Sub LoadComboSingleStatus()
''
''
Debug.Print "LoadComboSingleStatus()"

    fMainForm.ComboSingleStatus.AddItem "Empty     "
    fMainForm.ComboSingleStatus.AddItem "Unmachined"
    fMainForm.ComboSingleStatus.AddItem "Machined  "
    fMainForm.ComboSingleStatus.AddItem "Reserved  "
    fMainForm.ComboSingleStatus.AddItem "Mask      "
    fMainForm.ComboSingleStatus.AddItem "Occupied  "
    fMainForm.ComboSingleStatus.AddItem "Broken Tool"
    fMainForm.ComboSingleStatus.AddItem "Disable   "

End Sub

Public Sub LoadComboLocations()
''
''
Debug.Print "LoadComboLocations()"

    fMainForm.ComboLocations.AddItem "Kiosk"
    fMainForm.ComboLocations.AddItem "Chuck"
    
    If UseSaw = True Then
        fMainForm.ComboLocations.AddItem "Spindle"
        fMainForm.ComboLocations.AddItem "Station 1"
        fMainForm.ComboLocations.AddItem "Station 2"
    End If



End Sub

Public Sub DisplayStringFromRobot()

    With fMainForm.ListAutomatMessage
    
        If RobotStrings(11) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(11))
        End If
            
        If RobotStrings(14) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(14))
        End If
        
        If RobotStrings(18) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(18))
        End If
        
        If RobotStrings(19) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(19))
        End If
        
        If RobotStrings(17) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(17))
        End If
        
        If RobotStrings(20) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(20))
        End If
        
        If RobotStrings(21) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & "  " & RobotStrings(21))
        End If
        
        If .ListCount > 10 Then .TopIndex = .ListCount - 1
        
    End With
    
    fMainForm.PanelK2P(3).Caption = RobotStrings(14) & "..."
    fMainForm.PanelP2k(3).Caption = RobotStrings(15) & "..."
    
    With fMainForm.ListRobotInfo
        If RobotStrings(11) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 11) " & RobotStrings(11))
            Call WriteString(11, "")
        End If
        
        If RobotStrings(14) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 14) " & RobotStrings(14))
            Call WriteString(14, "")
        End If
        
        If RobotStrings(15) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 15) " & RobotStrings(15))
            Call WriteString(15, "")
        End If
        
        If RobotStrings(18) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 18) " & RobotStrings(18))
            Call WriteString(18, "")
        End If
        
        If RobotStrings(19) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 19) " & RobotStrings(19))
            Call WriteString(19, "")
        End If
        
        If RobotStrings(17) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 17) " & RobotStrings(17))
            Call WriteString(17, "")
        End If
        
        If RobotStrings(20) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 20) " & RobotStrings(20))
            Call WriteString(20, "")
        End If
        
        If RobotStrings(21) <> "" Then
            .AddItem (Format(now, "hh:mm:ss") & " (String 21) " & RobotStrings(21))
            Call WriteString(21, "")
        End If
        
        If .ListCount > 1 Then
            .TopIndex = .ListCount - 1
        End If
        
    End With

End Sub

Public Sub ResetAllRobotStrings()

Dim kk As Integer

    fMainForm.PanelK2P(3).Caption = ""
    fMainForm.PanelP2k(3).Caption = ""
    
    For kk = 1 To 25
        RobotStrings(kk) = ""
    Next
    
    
End Sub




Public Sub Update_ASM_GUI(ByVal myGear As AppGear)
'1.this function enable/disable some of the Buttons,
'   Frames,Fetures in the GUI ,according to the input state :
'2. Manual  = 1
'   Semi    = 2
'   Automat = 3
'3.the function get the Gear Of the Automation.
'4.the function return no parameters.

Dim j As Integer

    If myGear = Manual Then ' manual
    
        fMainForm.BtnTeachSingle.Enabled = True

        fMainForm.BtnTestPocketToPocket.Enabled = False
        fMainForm.BtnTestAllPockets.Enabled = False
        fMainForm.SSFrame1(11).Enabled = True
        fMainForm.SSFrame1(1200).Enabled = True
        fMainForm.FrameManualOperations.Enabled = True
        fMainForm.SSFrame1(8).Enabled = False: fMainForm.CmdSysOperation(3).Enabled = False
        fMainForm.FrameRunMode.Enabled = False
        
        For j = 1 To 10
            fMainForm.cmdOutOn(j).Enabled = True
            fMainForm.cmdOutOff(j).Enabled = True
        Next
        If AppGripperStyle = 1 Then
            fMainForm.cmdOutOn(10).Enabled = False
            fMainForm.cmdOutOff(10).Enabled = False
        ElseIf AppGripperStyle = 2 Then
            fMainForm.cmdOutOn(10).Enabled = True
            fMainForm.cmdOutOff(10).Enabled = True
        End If
        
        
        fMainForm.cmdGoParkingPos.Enabled = False
        fMainForm.cmdGoExchangePos.Enabled = False
        fMainForm.cmdGoZeroPos.Enabled = False
        
        For j = 0 To 3
            fMainForm.Option1(j).Enabled = False
        Next
        
        For j = 0 To 11
            fMainForm.BtnJog(j).Enabled = True
        Next
        fMainForm.SSTab3.Enabled = True
        fMainForm.TopToolBar.Buttons(3).Value = tbrUnpressed
        fMainForm.TopToolBar.Buttons(4).Value = tbrUnpressed
        fMainForm.TopToolBar.Buttons(5).Value = tbrPressed
        
        fMainForm.TopToolBar.Buttons(7).Enabled = True
        fMainForm.TopToolBar.Buttons(8).Enabled = True
        
        fMainForm.SSFrame1(2).Enabled = False
        fMainForm.SSFrame1(3).Enabled = False
        
        fMainForm.cmdLoadTool.Enabled = False
        fMainForm.cmdUnloadTool.Enabled = False
        
        fMainForm.BtnTestPocketToPocket.Enabled = False
        fMainForm.BtnTestAllPockets.Enabled = False
        fMainForm.BtnStartLoadUnload.Enabled = False
        fMainForm.BtnStopLoadUnload.Enabled = False
        
        fMainForm.UpDownTestShelf.Enabled = False
        fMainForm.UpDownTestPocket(0).Enabled = False
        fMainForm.UpDownTestPocket(1).Enabled = False
        fMainForm.UpDownTestShelf1.Enabled = False
        fMainForm.UpDownTestShelf2.Enabled = False
        fMainForm.UpDownTestShelf.Enabled = False
        fMainForm.UpDownMyDiameter.Enabled = False
        fMainForm.UpDown5.Enabled = False
        fMainForm.UpDown3.Enabled = False
        fMainForm.UpDown2.Enabled = False
        fMainForm.UpDownTestPocket(2).Enabled = False
        fMainForm.FrameLoadUnload.Enabled = False
        fMainForm.SSFrame8.Enabled = False
        fMainForm.CmdResetAllTests.Enabled = False
        fMainForm.LabelApp(2).Caption = "Manual"
        
        If UseSaw = True Then
            fMainForm.FrameSaw.Enabled = False
        ElseIf UseSaw = False Then
            fMainForm.FrameSaw.Enabled = False
        End If
        
             
    ElseIf myGear = Semi Then ' semi
    
        fMainForm.BtnTeachSingle.Enabled = False

        fMainForm.BtnTestPocketToPocket.Enabled = True
        fMainForm.BtnTestAllPockets.Enabled = True
        fMainForm.SSFrame1(11).Enabled = False
        fMainForm.SSFrame1(1200).Enabled = True
        fMainForm.FrameManualOperations.Enabled = True
        fMainForm.SSFrame1(8).Enabled = False: fMainForm.CmdSysOperation(3).Enabled = False
        fMainForm.FrameRunMode.Enabled = True
        For j = 1 To 10
            fMainForm.cmdOutOn(j).Enabled = False
            fMainForm.cmdOutOff(j).Enabled = False
        Next
        fMainForm.cmdGoParkingPos.Enabled = True
        fMainForm.cmdGoExchangePos.Enabled = True
        fMainForm.cmdGoZeroPos.Enabled = True
        
        For j = 0 To 3
            fMainForm.Option1(j).Enabled = True
        Next
        
        For j = 0 To 11
            fMainForm.BtnJog(j).Enabled = False
        Next
        fMainForm.SSTab3.Enabled = False
        
        fMainForm.TopToolBar.Buttons(3).Value = tbrUnpressed
        fMainForm.TopToolBar.Buttons(4).Value = tbrPressed
        fMainForm.TopToolBar.Buttons(5).Value = tbrUnpressed
        
        fMainForm.SSFrame1(2).Enabled = False
        fMainForm.SSFrame1(3).Enabled = False
        
        fMainForm.cmdLoadTool.Enabled = False
        fMainForm.cmdUnloadTool.Enabled = False
                
        fMainForm.BtnTestPocketToPocket.Enabled = True
        fMainForm.BtnTestAllPockets.Enabled = True
        fMainForm.BtnStartLoadUnload.Enabled = True
        fMainForm.BtnStopLoadUnload.Enabled = True
        
        fMainForm.UpDownTestShelf.Enabled = True
        fMainForm.UpDownTestPocket(0).Enabled = True
        fMainForm.UpDownTestPocket(1).Enabled = True
        fMainForm.UpDownTestShelf1.Enabled = True
        fMainForm.UpDownTestShelf2.Enabled = True
        fMainForm.UpDownTestShelf.Enabled = True
        fMainForm.UpDownMyDiameter.Enabled = True
        fMainForm.UpDown5.Enabled = True
        fMainForm.UpDown3.Enabled = True
        fMainForm.UpDown2.Enabled = True
        fMainForm.UpDownTestPocket(2).Enabled = True
        fMainForm.SSFrame8.Enabled = True
        fMainForm.FrameLoadUnload.Enabled = True
        fMainForm.LabelApp(2).Caption = "Semi"
        fMainForm.CmdResetAllTests.Enabled = True
        
        If UseSaw = True Then
            fMainForm.FrameSaw.Enabled = True
        ElseIf UseSaw = False Then
            fMainForm.FrameSaw.Enabled = False
        End If
            
    ElseIf myGear = Automat Then ' automat
    
        fMainForm.BtnTeachSingle.Enabled = False

        fMainForm.BtnTestPocketToPocket.Enabled = False
        fMainForm.BtnTestAllPockets.Enabled = False
      
        fMainForm.SSFrame1(11).Enabled = False
        fMainForm.SSFrame1(1200).Enabled = False
        fMainForm.FrameManualOperations.Enabled = False
        fMainForm.SSFrame1(8).Enabled = True: fMainForm.CmdSysOperation(3).Enabled = True
        fMainForm.FrameRunMode.Enabled = True
        For j = 1 To 10
            fMainForm.cmdOutOn(j).Enabled = False
            fMainForm.cmdOutOff(j).Enabled = False
        Next
        fMainForm.cmdGoParkingPos.Enabled = False
        fMainForm.cmdGoExchangePos.Enabled = False
        fMainForm.cmdGoZeroPos.Enabled = False
        
        For j = 0 To 3
            fMainForm.Option1(j).Enabled = False
        Next
        
        For j = 0 To 11
            fMainForm.BtnJog(j).Enabled = False
        Next
        fMainForm.SSTab3.Enabled = False
        
        fMainForm.TopToolBar.Buttons(3).Value = tbrPressed
        fMainForm.TopToolBar.Buttons(4).Value = tbrUnpressed
        fMainForm.TopToolBar.Buttons(5).Value = tbrUnpressed

        fMainForm.SSFrame1(2).Enabled = True
        fMainForm.SSFrame1(3).Enabled = True
        
        fMainForm.cmdLoadTool.Enabled = True
        fMainForm.BtnStartLoadUnload.Enabled = False
        fMainForm.BtnStopLoadUnload.Enabled = False
        
        fMainForm.BtnTestPocketToPocket.Enabled = False
        fMainForm.BtnTestAllPockets.Enabled = False
        fMainForm.BtnStartLoadUnload.Enabled = False
        fMainForm.UpDownTestShelf.Enabled = False
        fMainForm.UpDownTestPocket(0).Enabled = False
        fMainForm.UpDownTestPocket(1).Enabled = False
        fMainForm.UpDownTestShelf1.Enabled = False
        fMainForm.UpDownTestShelf2.Enabled = False
        fMainForm.UpDownTestShelf.Enabled = False
        fMainForm.UpDownMyDiameter.Enabled = False
        fMainForm.UpDown5.Enabled = False
        fMainForm.UpDown3.Enabled = False
        fMainForm.UpDown2.Enabled = False
        fMainForm.UpDownTestPocket(2).Enabled = False
        fMainForm.SSFrame8.Enabled = False
        fMainForm.FrameLoadUnload.Enabled = False
        fMainForm.LabelApp(2).Caption = "Automat"
        fMainForm.CmdResetAllTests.Enabled = False
        
        If UseSaw = True Then
            fMainForm.FrameSaw.Enabled = False
        ElseIf UseSaw = False Then
            fMainForm.FrameSaw.Enabled = False
        End If
                
    End If

End Sub



Public Sub DisplayWPProperties()

Dim TempAmountLeft As Integer

    TempAmountLeft = AllWP(AppWPIndex).ToolAmount - ToolCounter
    AllWP(AppWPIndex).ToolAmountLeft = TempAmountLeft
    
    ''display on the header
    fMainForm.pnlMode(0).Caption = AllWP(AppWPIndex).NCProgram
    fMainForm.pnlMode(1).Caption = AllWP(AppWPIndex).ToolAmountLeft
    fMainForm.pnlMode(2).Caption = AllWP(AppWPIndex).WPNumber
    fMainForm.LabelApp(1).Caption = "Diameter :" & AllWP(AppWPIndex).ToolDiameter
    
    If HermleWPMode = NightMode Then
        fMainForm.LabelApp(4).Caption = "Night Mode"
    ElseIf HermleWPMode = OneTool Then
        fMainForm.LabelApp(4).Caption = "One Tool"
    ElseIf HermleWPMode = WorkPiece Then
        fMainForm.LabelApp(4).Caption = "Work Piece"
    End If
    
    
    ''display in the : Automat->Tools.
    fMainForm.TextAmount = AllWP(AppWPIndex).ToolAmount
    fMainForm.TextAmountLeft = AllWP(AppWPIndex).ToolAmountLeft
        

End Sub

Public Sub DisplayAllInputs()

Dim InputState As Integer
    
    ''******
    ''kiosk
    ''******
    InputState = ReadIO(IOinputsAdd.kioskdooropen) ''read from controller
    Call DisplayOneInput(InputState, kioskdooropen) ''display on screen
            
    InputState = ReadIO(IOinputsAdd.KioskDoorClose)
    Call DisplayOneInput(InputState, KioskDoorClose)
    
    InputState = ReadIO(IOinputsAdd.UserAck)
    Call DisplayOneInput(InputState, UserAck)
    
    InputState = ReadIO(IOinputsAdd.ValvesClose)
    Call DisplayOneInput(InputState, ValvesClose)
    
    InputState = ReadIO(IOinputsAdd.HolderDirectRobot)
    Call DisplayOneInput(InputState, HolderDirectRobot)
    
    InputState = ReadIO(IOinputsAdd.ToolFound)
    Call DisplayOneInput(InputState, ToolFound)
    
    ''******
    '' robot
    ''******
    InputState = ReadIO(IOinputsAdd.GripperIsOpen)
    Call DisplayOneInput(InputState, GripperIsOpen)
    InputState = ReadIO(IOinputsAdd.Gripper2IsOpen)
    Call DisplayOneInput(InputState, Gripper2IsOpen)
    
    ''*********
    ''  cabinet
    ''*********
    InputState = ReadIO(IOinputsAdd.Estop_1)
    Call DisplayOneInput(InputState, Estop_1)
    
    InputState = ReadIO(IOinputsAdd.Estop_2)
    Call DisplayOneInput(InputState, Estop_2)
    
    InputState = ReadIO(IOinputsAdd.UserInterlock_1)
    Call DisplayOneInput(InputState, UserInterlock_1)
    
    InputState = ReadIO(IOinputsAdd.UserInterlock_2)
    Call DisplayOneInput(InputState, UserInterlock_2)
    
    InputState = ReadIO(IOinputsAdd.PCInterlock_1)
    Call DisplayOneInput(InputState, PCInterlock_1)
    
    InputState = ReadIO(IOinputsAdd.PCInterlock_2)
    Call DisplayOneInput(InputState, PCInterlock_2)
    

End Sub

Public Sub DisplayOneInput(ByVal InputState As Integer, ByVal Address As Integer)
''the address is the address relate to the input.
''you  can see the address number in the TP
''IN\OUT - > universal input

    If InputState = 1 Then
        fMainForm.LabelInput(Address).Caption = "On"
        fMainForm.LabelInput(Address).ForeColor = &H8000&
    ElseIf InputState = 0 Then
        fMainForm.LabelInput(Address).Caption = "Off"
        fMainForm.LabelInput(Address).ForeColor = vbRed
    End If

End Sub

Public Sub DisplayAppToolType()
    
    If AppToolType = HSK Then
        fMainForm.txtWorkPiece(7).Text = "HSK"
        
    ElseIf AppToolType = Drill Then
        fMainForm.txtWorkPiece(7).Text = "Drill"
    
    ElseIf AppToolType = Round Then
        fMainForm.txtWorkPiece(7).Text = "Round"
    
    End If
    
End Sub

''Public Sub DisplayToolAmount()
''
''Dim TempInt As Integer
''Dim bb As Boolean
''
''
''        ''display in the : Amount->Tools.
''        fMainForm.TextAmount = AllWP(AppWPIndex).ToolAmount
''        fMainForm.TextAmountLeft = AllWP(AppWPIndex).ToolAmountLeft
''
''End Sub

Public Sub DisplayGenLocNames()

    With frmOptions.AllLocationsTable
    
    .TextMatrix(8, 1) = "before spindle"
    .TextMatrix(10, 1) = "spindle"
    .TextMatrix(11, 1) = "kiosk"
    .TextMatrix(12, 1) = "chuck"
    .TextMatrix(13, 1) = "take from 16 "
    .TextMatrix(14, 1) = "parking"
    .TextMatrix(15, 1) = "exchange gripper "
    .TextMatrix(16, 1) = "take pocket"
    .TextMatrix(17, 1) = "place pocket"
    .TextMatrix(18, 1) = "station "
    .TextMatrix(19, 1) = "Go location"
    .TextMatrix(20, 1) = "place in 17"
    .TextMatrix(21, 1) = "pocket 101"
    .TextMatrix(22, 1) = "pocket 110"
    .TextMatrix(23, 1) = "pocket 201"
    .TextMatrix(24, 1) = "pocket 210"
    .TextMatrix(25, 1) = "pocket 301"
    .TextMatrix(26, 1) = "pocket 310"
    .TextMatrix(27, 1) = "mframe from p20"
    .TextMatrix(28, 1) = "  "
    .TextMatrix(29, 1) = "Middle retract above Chuck"
    .TextMatrix(30, 1) = "chuck calc"
    .TextMatrix(31, 1) = "retract chuck (indoors)"
    .TextMatrix(32, 1) = "retract above chuck"
    .TextMatrix(33, 1) = "above chuck"
    .TextMatrix(34, 1) = "Middle retract chuck"
    .TextMatrix(35, 1) = "  "
    .TextMatrix(36, 1) = "Middle retract above pocket"
    .TextMatrix(37, 1) = "pocket calc"
    .TextMatrix(38, 1) = "retract pocket"
    .TextMatrix(39, 1) = "retract above pocket"
    .TextMatrix(40, 1) = "above pocket"
    .TextMatrix(41, 1) = " "
    .TextMatrix(42, 1) = "Middle retract pocket"
    .TextMatrix(43, 1) = " "
    .TextMatrix(44, 1) = "Middle retract above Kiosk"
    .TextMatrix(45, 1) = "Kiosk calc"
    .TextMatrix(46, 1) = "retract Kiosk"
    .TextMatrix(47, 1) = "retract above Kiosk"
    .TextMatrix(48, 1) = "above Kiosk"
    .TextMatrix(49, 1) = "Middle retract Kiosk"
    .TextMatrix(50, 1) = "  "
    .TextMatrix(51, 1) = "Middle above chuck"
    .TextMatrix(52, 1) = "Middle above retract chuck"
    .TextMatrix(53, 1) = "Middle above pocket"
    .TextMatrix(54, 1) = "Middle above retract pocket"
    .TextMatrix(55, 1) = "Middle above kiosk"
    .TextMatrix(56, 1) = "Middle above retract kiosk"
    .TextMatrix(57, 1) = "  "
    .TextMatrix(58, 1) = "Middle retract above station"
    .TextMatrix(59, 1) = "station calc"
    .TextMatrix(60, 1) = "retract station"
    .TextMatrix(61, 1) = "retract above station"
    .TextMatrix(62, 1) = "above station"
    .TextMatrix(63, 1) = "Middle retract station"
    .TextMatrix(64, 1) = "Middle retract above spindle"
    .TextMatrix(65, 1) = "spindle calc"
    .TextMatrix(66, 1) = "retract spindle"
    .TextMatrix(67, 1) = "retract above spindle"
    .TextMatrix(68, 1) = "above spindle"
    .TextMatrix(69, 1) = "Middle retract spindle"
    .TextMatrix(70, 1) = "Middle above station"
    .TextMatrix(71, 1) = " "
    .TextMatrix(72, 1) = " "
    .TextMatrix(73, 1) = "Middle above retract chuck"
    .TextMatrix(74, 1) = " "
    .TextMatrix(75, 1) = " "
    .TextMatrix(76, 1) = " "
    .TextMatrix(77, 1) = " "
    .TextMatrix(79, 1) = " "
    .TextMatrix(80, 1) = "Landmark from 104 t 111"
    .TextMatrix(81, 1) = "Help mulmat retract calc"
    .TextMatrix(82, 1) = "Help mulmat retract calc"
    .TextMatrix(83, 1) = "Help mulmat middle retract"
    .TextMatrix(84, 1) = " "
    .TextMatrix(85, 1) = "Help mulmat middle retract"
    .TextMatrix(86, 1) = "Help mulmat retract calc tool 3"
    .TextMatrix(87, 1) = "Help delta high and retract"
    .TextMatrix(88, 1) = "Sensor Tool found retract"
    .TextMatrix(89, 1) = "Sensor Tool found pocket"
    .TextMatrix(90, 1) = "Middle above station"
    .TextMatrix(91, 1) = " "
    .TextMatrix(92, 1) = " "
    
    End With

End Sub


Public Sub DisplayAlarmStatus()

    
Dim ret As Integer
Dim Data As Integer
Dim Msg As String
Dim AlarmString As String * 36

    If HermleCommState = OffLine Then
        Exit Sub
    End If
    
    ret = ModHandShake.ReadRobotAlarm
    
    If ret = 0 Then
        fMainForm.txtCycleMessage(1).BackColor = vbGreen
        fMainForm.txtCycleMessage(1).Caption = " ROBOT RUN "
    ElseIf ret = 1 Then
        fMainForm.txtCycleMessage(1).BackColor = vbRed
        fMainForm.txtCycleMessage(1).Caption = CStr(Data) & " " & RobotAlarmString
        ModLogFile.LogAddLine (" Alarm:" & RobotAlarmString)
        Call ModGUI.ListHMIUpdate("Alarm:" & RobotAlarmString)
        
    End If
        
End Sub


Public Sub ListHMIUpdate(ByVal MyString As String)

    If UseHMIInfo = True Then
        If fMainForm.CheckListHMI.Value = vbChecked Then
            fMainForm.ListHMIInfo.AddItem (Format(now, "hh:mm:ss") & " " & MyString)
        End If
    End If

End Sub






