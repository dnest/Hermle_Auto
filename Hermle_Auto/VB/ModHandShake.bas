Attribute VB_Name = "ModHandShake"
Option Explicit

''Public AllPocketsStep As Integer
Public AllPocketsShelf As Integer
Public AllPocketsColumn As Integer
Public AllPocketsDiameter As Integer
Dim MyCurrentJob As String



Public Sub ReadAllTestsStatus()
''1.the function read from robot the state of all TESTS.
''2.the function put the value into Jobs.something in the memoory.
''3.LEGEND:
''  1= idle
''  2 =run
''  3 =done
Dim temp As Double

    Call ReadByte(39, Tests.PocketToPocket())
    Call ReadByte(42, Tests.KioskToPocket())
    Call ReadByte(43, Tests.PocketToKiosk())
    Call ReadByte(44, Tests.PocketToChuck())
    Call ReadByte(45, Tests.ChuckToPocket())
    Call ReadByte(48, Tests.AllPockets())
    
End Sub

Public Sub WriteAllJobsStatus(MyJobState As JobState)
'
'    idle = 1
'    run = 2
'    done = 3

Dim jj As Integer
  
    ArrayByte(0) = MyJobState

    ArrayByte(0) = 0
    For jj = 31 To 41
        Call WriteByte(jj, ArrayByte())
    Next
    
End Sub

Public Sub ResetAllJobsStatus()

Dim jj As Integer
Dim MyJobState As Integer

    If HermleCommState = OffLine Then
        Exit Sub
    End If

    ArrayByte(0) = IDLE
    For jj = 29 To 38
        Call WriteByte(jj, ArrayByte())
    Next
    
    Jobs.AutoStart(0) = 0
    Jobs.ChuckToPocket(0) = 0
    Jobs.Exchange(0) = 0
    Jobs.KioskToPocket(0) = 0
    Jobs.Parking(0) = 0
    Jobs.PocketToChuck(0) = 0
    Jobs.PocketToKiosk(0) = 0

End Sub

Public Function ReadOneJobStatus(ByVal jobname As Double) As Integer
''
''1.the function read one byte from robot.
''2.tyhe function get the JobName and read the appropieate number.
''3.Legend :
        ''1 idle
        ''2 run
        ''3 done
    
Dim jobnumber As Integer

    Select Case jobname

        Case Jobs.Parking(0)
            jobnumber = 33
        
        Case Jobs.Exchange(0)
            jobnumber = 34
        
        Case Jobs.KioskToPocket(0)
            jobnumber = 35
        
        Case Jobs.PocketToKiosk(0)
            jobnumber = 36
        
        Case Jobs.PocketToChuck(0)
            jobnumber = 37
        
        Case Jobs.ChuckToPocket(0)
            jobnumber = 38
           
    End Select
    
    Call ReadByte(jobnumber, ArrayByte())
    ReadOneJobStatus = ArrayByte(0)

End Function

Public Sub WriteOneJobStatus(ByVal jobname As Double, ByVal MyJobstatus As JobState)

Dim jobnumber As Double

    Select Case jobname

        Case Jobs.Parking(0)
            jobnumber = 33
        
        Case Jobs.Exchange(0)
            jobnumber = 34
        
        Case Jobs.KioskToPocket(0)
            jobnumber = 35
        
        Case Jobs.PocketToKiosk(0)
            jobnumber = 36
        
        Case Jobs.PocketToChuck(0)
            jobnumber = 37
        
        Case Jobs.ChuckToPocket(0)
            jobnumber = 38
        
    End Select
    
    ArrayByte(0) = MyJobstatus

    Call WriteByte(jobnumber, ArrayByte())
    

End Sub

Public Sub ReadAllRobotRequest()

Dim MyPocket As String

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim temp(120) As Double

    '7*30ms = 210 ms

    Call ReadByte(11, HandShake.RequestPlaceFromKiosk)
    temp(11) = HandShake.RequestPlaceFromKiosk(0)
    
    Call ReadByte(13, HandShake.RequestTakeToKiosk())
    temp(13) = HandShake.RequestTakeToKiosk(0)
     
    Call ReadByte(15, HandShake.RequestPlaceFromChuck)
    temp(15) = HandShake.RequestPlaceFromChuck(0)
     
    Call ReadByte(17, HandShake.RequestTakeToChuck)
    temp(17) = HandShake.RequestTakeToChuck(0)
    
    If HandShake.RequestTakeToChuck(0) = 1 Then
        ModGUI.ListHMIUpdate ("got request pocket to chuck arrived")
    End If
    
    Call ReadByte(21, HandShake.RequestNewWorkPiece)
    temp(21) = HandShake.RequestNewWorkPiece(0)
    
    If temp(21) = 1 Then
        temp(21) = temp(21)
    End If
     
    Call ReadByte(56, HandShake.RequestDeleteWPiece)
    temp(56) = HandShake.RequestDeleteWPiece(0)
     
    Call ReadByte(65, HandShake.MaskUser)
    temp(65) = HandShake.MaskUser(0)
     
    Call ReadByte(118, HandShake.RequestChangePocketStatus)
    temp(118) = HandShake.RequestChangePocketStatus(0)

''b11 Request Find pocket to place from Kiosk
''b12 Done Find pocket to place from Kiosk
''b13 Request Find pocket to take to Kiosk
''b14 Done Find pocket to  take to Kiosk
''b15 Request Find pocket to place from Chuck
''b16 Done Find pocket to place from Chuck
''b17 Request Find pocket to take to Chuck
''b18 Done Find pocket to  take to Chuck

End Sub

Public Sub WriteOneCommByte(ByVal MyCommByte As Integer, ByVal MyValue As Integer)
    
    ArrayByte(0) = MyValue
    Call WriteByte(MyCommByte, ArrayByte())
    Call ModGUI.ListHMIUpdate("Write byte. Byte (" & MyCommByte & ")=" & MyValue)
    
End Sub

Public Sub ResetOneCommByte(ByVal MyCommByte As Integer)

    ArrayByte(0) = 0
    Call WriteByte(MyCommByte, ArrayByte())
    Call ModGUI.ListHMIUpdate("Reset byte number : " & MyCommByte)

End Sub


Public Sub SetOneCommByte(ByVal MyCommByte As Integer)

    ArrayByte(0) = 1
    Call WriteByte(MyCommByte, ArrayByte())
    Call ModGUI.ListHMIUpdate("Set byte number. " & MyCommByte)

End Sub


Public Sub ResetAllCommBytes()
''1.this function reset all the request from the robot.
    
Dim kk As Integer

    ArrayByte(0) = 0
    For kk = 11 To 27
        Call WriteByte(kk, ArrayByte())
    Next

End Sub



Public Sub WriteDataToRobot()
''1.this function write the wanted pocket to the robot.
''2.the function is called from the timer.
''3.the function is being called after the :"ReadAllRobotRequest "

Dim MyPocket As String
Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim diameter As Integer
Dim kk As Integer
Dim bb As Boolean
Dim ret As Integer
Dim digit As String

Dim shelf_to As Integer
Dim column_to As Integer
Dim pocket_to As Integer

Dim shelf_from As Integer
Dim column_from As Integer
Dim pocket_from As Integer
Dim TempStatus As PocketStatus
Dim StringStatus As String

    ''************************************************************************
    '' kiosk to pocket
    ''************************************************************************
    If (HandShake.RequestPlaceFromKiosk(0) = 1) Then
        Call ModGUI.ListHMIUpdate("got request kiosk to pocket")
        If LoadFirstTime = True Then

            shelf_to = CInt(fMainForm.txtLoadUnloadShelf.Text) '''get data from GUI
            column_to = CInt(Right(fMainForm.txtLoadUnloadPocket.Text, 2))
            pocket_to = AppDiameter
            
            ''send sensor location
            Call ModHandShake.WriteSensorLocation(shelf_to, column_to, pocket_to, 20)
               
            TempStatus = GetPocketStatus(shelf_to, column_to) '''make sure the pocket is empty.
            If TempStatus <> empty_ Then
                ''**********
                MyPocket = FindEmptyPocket()
                If (MyPocket = "0") Then ''           If did not find pocket
                    Call ResetOneCommByte(11) ''      handShake  - reset the request
                    Call WriteOneCommByte(12, 2) ''   2=Not Found
                    Call FrmDialog.ShowDialogForm(70, 17, 18, "ModHandShake", "WriteDataToRobot()", GeneralError)
                    Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
                    fMainForm.tmrRobotQuery.Enabled = False ''    pause the main timer
                    Exit Sub
                End If
                If ((AppToolType = HSK) And (AppDiameter > 150)) Then          '''make sure the neighbors are empty.
                    bb = ModAutomationStatus.GetPocketNeighborsStatus(AppCurrentPocket, occupied)
                    If bb = False Then
                        Call PauseRobotJob("2_AUTO_START_TEST.JBI")
                        ret = MsgBox("Kiosk To Pocket : " & vbCrLf _
                            & "can not place tool in pocket." & vbCrLf & _
                            "Diameter is bigger then 150.Neighbors are not Occupied .", vbExclamation, _
                            "ModHandShake.WriteDataToRobot():kiosk to pocket.")
                        fMainForm.tmrRobotQuery.Enabled = False                ''pause the main timer
                    End If
                End If
                AppCurrentPocket = MyPocket
                fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
                ModLogFile.LogAddLine (" kiosk to Pocket: " & AppCurrentPocket)
                Call ModGUI.ListHMIUpdate("kiosk to Pocket: " & AppCurrentPocket)
                Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
                Call BuildArrayPosition(shelf, column, pocket, AppToolType)  ''build a position by its parameters.
                
                ''check if there is no 0 value in X,Y,Z
                For kk = 2 To 4
                    If (ArrayPosition(kk) = 0) Then
                        ret = MsgBox("ERROR ! the pocket :" & MyPocket & " contain ZERO value." _
                            , vbExclamation, "WriteDataToRobot()")
                        Exit Sub
                    End If
                Next
                
                ''send pocket to robot
                bb = MotoComToolBox.WritePosition(17, ArrayPosition())
                If bb = False Then
                    Call FrmDialog.ShowDialogForm(38, 38, 38, "modHandShake", "WriteDataToRobot()", GeneralError)
                    Exit Sub
                Else
                    
                End If
                
                Call ResetOneCommByte(11) ''  handShake  - reset the request
                Call SetOneCommByte(12) ''''  rise a flag up."position was written"
               
            ''***********
                Exit Sub
            End If

            
            If column_from < 10 Then
                digit = "0"
            Else
                digit = ""
            End If
            
            If ((AppToolType = HSK) And (AppDiameter > 150)) Then          '''make sure the neighbors are occupied.
                bb = ModAutomationStatus.GetPocketNeighborsStatus(AppCurrentPocket, occupied)
                If bb = False Then
                    Call PauseRobotJob("2_AUTO_START_TEST.JBI")
                    ret = MsgBox("can not place tool in pocket." & vbCrLf _
                        & "tool is big.Neighbors not occupied.", _
                            vbExclamation, "WriteDataToRobot():Kiosk To Pocket")
                    fMainForm.tmrRobotQuery.Enabled = False                ''pause the main timer
                End If
            End If

            MyPocket = CStr(shelf_to) & digit & CStr(column_to)
            AppCurrentPocket = MyPocket
            fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket ''display pocket on screen
            ModLogFile.LogAddLine (" kiosk to Pocket: " & AppCurrentPocket)  ''save pocket in logFile
            Call ModGUI.ListHMIUpdate("kiosk to Pocket: " & AppCurrentPocket)
            Call BuildArrayPosition(shelf_to, column_to, pocket_to, AppToolType) ''build a position by its parameters.
            ''send pocket to robot
            bb = MotoComToolBox.WritePosition(17, ArrayPosition())
            If bb = False Then
                Call FrmDialog.ShowDialogForm(38, 38, 38, "modHandShake", "WriteDataToRobot()", GeneralError)
                Exit Sub
            End If
        ElseIf LoadFirstTime = False Then
        
            MyPocket = FindEmptyPocket()
            
            ''write sensor location
            Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
            Call ModHandShake.WriteSensorLocation(shelf, column, pocket, 20)
            ''
            If (MyPocket = "0") Then ''           If did not find pocket
                Call ResetOneCommByte(11) ''      handShake  - reset the request
                Call WriteOneCommByte(12, 2) ''   2=Not Found
                Call FrmDialog.ShowDialogForm(70, 17, 18, "ModHandShake", "WriteDataToRobot()", GeneralError)
                Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
                fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
                Exit Sub
            End If

            AppCurrentPocket = MyPocket
            fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
            ModLogFile.LogAddLine (" kiosk to Pocket: " & AppCurrentPocket)
            Call ModGUI.ListHMIUpdate("kiosk to Pocket: " & AppCurrentPocket)
            Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
            Call BuildArrayPosition(shelf, column, pocket, AppToolType) ''              build a position by its parameters.
        End If
        
        ''check if there is no 0 value.
        bb = NoZeroValues(ArrayPosition)
        If bb = False Then
            ret = MsgBox("ERROR ! the pocket :" & MyPocket & " contain ZERO value.", _
                 vbExclamation, _
                "WriteDataToRobot():Kiosk To Pocket.")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI")
            Exit Sub
        End If

        ''send pocket to robot
        bb = MotoComToolBox.WritePosition(17, ArrayPosition())
        If bb = False Then
            Call FrmDialog.ShowDialogForm(38, 38, 38, "modHandShake", "WriteDataToRobot()", GeneralError)
            Exit Sub
        End If
        
        Call ResetOneCommByte(11) ''  handShake  - reset the request
        Call SetOneCommByte(12) ''''rise a flag up."position was written"
       
        
    ''************************************************************************
    ''   pocket to kiosk
    ''************************************************************************
    ElseIf (HandShake.RequestTakeToKiosk(0) = 1) Then
        Call ModGUI.ListHMIUpdate("got request pocket to kiosk ")
        
        If AppDiameter = 0 Then
            ret = MsgBox("Diameter can not be ZERO." & vbCrLf _
               & "make sure diameter is correct." & vbCrLf _
                , vbExclamation, _
                    "pocket to kiosk.")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
            fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
            Exit Sub
        End If
        
        If UnloadFirstTime = True Then
            shelf_from = CInt(fMainForm.txtLoadUnloadShelf.Text)
            column_from = CInt(Right(fMainForm.txtLoadUnloadPocket.Text, 2))
            pocket_from = AppDiameter
            
            ''write sensor location
            Call ModHandShake.WriteSensorLocation(shelf_from, column_from, pocket_from, 13)
            ''
            TempStatus = GetPocketStatus(shelf_from, column_from)
            If ((TempStatus <> machined) And (TempStatus <> unmachined)) Then
                ''****************************
                MyPocket = FindPocket(AppDiameter, AppToolType, machined, AppToolsWPiece, 0) '''the NC program is not importent.
                If (MyPocket = "0" Or MyPocket = "") Then ''          if did not find full pocket
                    Call ResetOneCommByte(13) ''     handshake.reset request.
                    Call WriteOneCommByte(14, 2) ''  2=Not Found
                    Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
                    fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
                    Call FrmDialog.ShowDialogForm(14, 67, 67, "ModHandShake", "WriteDataToRobot()", GeneralError)
                    Exit Sub
                End If
                AppCurrentPocket = MyPocket
                fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
                ModLogFile.LogAddLine ("Pocket to kiosk : " & AppCurrentPocket)   ''save pocket in LogFile
                Call ModGUI.ListHMIUpdate("Pocket to kiosk : " & AppCurrentPocket)
                Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
                Call BuildArrayPosition(shelf, column, pocket, AppToolType) ''              build a position by its parameters.
                TempStatus = GetPocketStatus(shelf, column)
                If TempStatus <> machined Then
                    Call FrmDialog.ShowDialogForm(14, 43, 43, "ModHandShake", "WriteDataToRobot()", GeneralError)
                    Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
                    fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
                    Exit Sub
                End If
                
                ''check if there is no 0 value.
                bb = NoZeroValues(ArrayPosition)
                If bb = False Then
                    ret = MsgBox("ERROR ! the pocket :" & MyPocket & " contain ZERO value." & vbCrLf _
                    & "make sure the diameter is corect" & vbCrLf, vbExclamation _
                         , "WriteDataToRobot():pocket to kiosk.")
                    Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
                    fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
                    Exit Sub
                End If
        
                bb = WritePosition(16, ArrayPosition()) ''send position to the robot.
                If bb = False Then
                    Call FrmDialog.ShowDialogForm(14, 40, 40, "modHandShake", "WriteDataToRobot()", GeneralError)
                    Exit Sub
                ElseIf bb = True Then
                   
                End If
                Call ResetOneCommByte(13)
                Call SetOneCommByte(14) ''hand shake :position was written.
                Exit Sub
                ''****************************
            End If
            
            If column_from < 10 Then
                digit = "0"
            Else
                digit = ""
            End If
            
            MyPocket = CStr(shelf_from) & digit & CStr(column_from)
            AppCurrentPocket = MyPocket
            fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
            ModLogFile.LogAddLine ("Pocket to kiosk : " & AppCurrentPocket)
            Call ModGUI.ListHMIUpdate("Pocket to kiosk : " & AppCurrentPocket)
            Call BuildArrayPosition(shelf_from, column_from, pocket_from, AppToolType) ''build a position by its parameters.
      
        ElseIf UnloadFirstTime = False Then
        
            MyPocket = FindPocket(AppDiameter, HSK, machined, AppToolsWPiece, 0) '''the NC program is not importent.

            If (MyPocket = "0" Or MyPocket = "") Then ''       if did not find full pocket
                Call ResetOneCommByte(13) ''                   handshake.reset request.
                Call WriteOneCommByte(14, 2) ''                2=Not Found
                Call PauseRobotJob("2_AUTO_START_TEST.JBI") '' pause the robot
                fMainForm.tmrRobotQuery.Enabled = False ''     pause the main timer
                Call FrmDialog.ShowDialogForm(14, 67, 67, "ModHandShake", "WriteDataToRobot()", GeneralError)
                Exit Sub
            End If
                        
            ''write sensor location
            Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
            Call ModHandShake.WriteSensorLocation(shelf, column, pocket, 13)
            ''
            AppCurrentPocket = MyPocket
            fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
            ModLogFile.LogAddLine (" Pocket to kiosk : " & AppCurrentPocket)   ''        save pocket in LogFile
            Call ModGUI.ListHMIUpdate("Pocket to kiosk : " & AppCurrentPocket)
            
            Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''              disassemble string to values
            Call BuildArrayPosition(shelf, column, pocket, AppToolType)  ''              build a position by its parameters.
            TempStatus = GetPocketStatus(shelf, column)
            If TempStatus <> machined Then
                Call FrmDialog.ShowDialogForm(14, 43, 43, "ModHandShake", "WriteDataToRobot()", GeneralError)
                Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
                fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
                Exit Sub
            End If
            
        End If
        
        ''check if there is no 0 value.
        bb = NoZeroValues(ArrayPosition)
        If bb = False Then
            ret = MsgBox("ERROR ! the pocket :" & MyPocket & " contain ZERO value.", _
                 vbExclamation, _
                "WriteDataToRobot():pocket to kiosk.")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI")
            Exit Sub
        End If
        
        bb = WritePosition(16, ArrayPosition()) ''send position to the robot.
        If bb = False Then
            Call FrmDialog.ShowDialogForm(14, 40, 40, "modHandShake", "WriteDataToRobot()", GeneralError)
            Exit Sub
        ElseIf bb = True Then
           
        End If
        Call ResetOneCommByte(13)
        Call SetOneCommByte(14) ''hand shake :position was written.
        
    ''************************************************************************
    ''   chuck to pocket
    ''************************************************************************
    ElseIf (HandShake.RequestPlaceFromChuck(0) = 1) Then
    
        If AppDiameter = 0 Then
            ret = MsgBox("Diameter can not be ZERO." & vbCrLf _
               & "Please Fill WorkPiece data." & vbCrLf _
                , vbExclamation, _
                    "Chuck to pocket.")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
            fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
            Exit Sub
        End If
        
        Call ModGUI.ListHMIUpdate("got request chuck to pocket")
        
        If ChuckUnloadFirstTime = True Then
            MyPocket = ModAutomationStatus.FindEmptyPocket
            ChuckUnloadFirstTime = False
            If MyPocket = "0" Then
                ret = MsgBox("can not find EMPTY pocket.", vbExclamation, "from chuck to pocket")
            End If
    
        ElseIf ChuckUnloadFirstTime = False Then
            MyPocket = ModAutomationStatus.FindPocket _
                (AppDiameter, HSK, reserved, _
                    AllWP(AppWPIndex).WPNumber, _
                        AllWP(AppWPIndex).NCProgram)
            If MyPocket = "0" Then
                ret = MsgBox("can not find RESERVED pocket.", vbExclamation, "from chuck to pocket")
            End If
        End If
        
        If (MyPocket = "0" Or MyPocket = "") Then ''if did not find pocket
            Call ResetOneCommByte(15)
            Call WriteOneCommByte(16, 2) ''2=Not Found
            ret = MsgBox("can not find RESERVED/EMPTY pocket.", vbExclamation, "from chuck to pocket")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
            fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
            Exit Sub
        End If
                   
        ''write sensor location
        Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
        Call ModHandShake.WriteSensorLocation(shelf, column, pocket, 20)
        
        
        AppCurrentPocket = MyPocket
        fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
        ModLogFile.LogAddLine (" chuck to Pocket: " & AppCurrentPocket)  ''save pocket in LogFile
        Call ModGUI.ListHMIUpdate("chuck to Pocket: " & AppCurrentPocket)
        
        Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
        Call BuildArrayPosition(shelf, column, pocket, AppToolType) ''              build a position {ArrayPosition} by its parameters.
        
        ''check if there is no 0 value.
        bb = NoZeroValues(ArrayPosition)
        If bb = False Then
            ret = MsgBox("ERROR ! the pocket :" & MyPocket & " contain ZERO value.", _
                 vbExclamation, _
                "WriteDataToRobot():chuck to pocket.")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI")
            Exit Sub
        End If

        bb = WritePosition(17, ArrayPosition()) ''send position to controller
        If bb = False Then
            ret = MsgBox("error write position - find pocket to place from chuck", vbExclamation, "Communication problem")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
            fMainForm.tmrRobotQuery.Enabled = False ''pause the main timer
            Exit Sub
        ElseIf bb = True Then
        ''
        End If
       
        Call ResetOneCommByte(15)
        Call SetOneCommByte(16)
    
       
    ''************************************************************************
    '' pocket to chuck
    ''************************************************************************
    ElseIf (HandShake.RequestTakeToChuck(0) = 1) Then
    
        Call ModGUI.ListHMIUpdate("got request pocket to chuck")
        AppDiameter = AllWP(AppWPIndex).ToolDiameter
        MyPocket = FindPocket(AppDiameter, HSK, unmachined, AllWP(AppWPIndex).WPNumber, AllWP(AppWPIndex).NCProgram)
        If (MyPocket = 0) Then
            Call ResetOneCommByte(17)
            Call WriteOneCommByte(18, 2)
            ''fMainForm.txtCycleMessage(1).Caption = "reset byte 17,byte 18=2"
            Call PauseRobotJob("2_AUTO_START_TEST.JBI") ''pause the robot
            fMainForm.tmrRobotQuery.Enabled = False ''    pause the main timer
            Call FrmDialog.ShowDialogForm(48, 48, 48, "modHandShake", "WriteDataToRobot()", InputIncomplete)
            Exit Sub
        End If
        
                                
        ''write sensor location
        Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''disassemble string to values
        Call ModHandShake.WriteSensorLocation(shelf, column, pocket, 13)
        ''
        AppCurrentPocket = MyPocket
        fMainForm.LabelApp(3).Caption = "Pocket: " & AppCurrentPocket
        ModLogFile.LogAddLine ("Pocket to chuck : " & AppCurrentPocket)   ''save pocket in LogFile
        Call ModGUI.ListHMIUpdate("Pocket to chuck : " & AppCurrentPocket)
        Call DecodePocketFromString(MyPocket, shelf, column, pocket) ''   disassemble string to values
        Call BuildArrayPosition(shelf, column, pocket, AppToolType)  ''   build a position {ArrayPosition} by its parameters.
        
        ''check if there is no 0 value.
        bb = NoZeroValues(ArrayPosition)
        If bb = False Then
            ret = MsgBox("ERROR ! the pocket :" & MyPocket & " contain ZERO value." & vbCrLf _
                 & "Make sure the Diameter is correct", vbExclamation, _
                "WriteDataToRobot():pocket to chuck.")
            Call PauseRobotJob("2_AUTO_START_TEST.JBI")
            Exit Sub
        End If
        
        bb = WritePosition(16, ArrayPosition())
        If bb = False Then
            Call FrmDialog.ShowDialogForm(48, 49, 49, "ModHandShake", "WriteDataToRobot()", GeneralError)
            Exit Sub
        ElseIf bb = True Then
            ''
        End If
        
        Call ResetOneCommByte(17)
        Call SetOneCommByte(18)

        
    ''************************************************************************
    '' Send New Work piece
    ''************************************************************************
    ElseIf (HandShake.RequestNewWorkPiece(0) = 1) Then
    
        If AllWP(AppWPIndex).ToolAmountLeft = 0 Then
            AppWPIndex = AppWPIndex ''''+ 1
        End If
        
        ModLogFile.LogAddLine ("Got request for New Work piece ")
        Call ModGUI.ListHMIUpdate("Got request for New Work piece ")
        
        ArrayDouble(0) = AllWP(AppWPIndex).NCProgram
        ArrayInteger(0) = AllWP(AppWPIndex).ToolAmount
        
        Call WriteDouble(13, ArrayDouble())   ''   send NC program number
        Call WriteInteger(10, ArrayInteger()) ''   send tool amount
        
        Call ResetOneCommByte(21) '''  handshake.reset the request.
        Call SetOneCommByte(22)   '''  handshake-WPiece was sent.

        ''byte 23 => no more Work Pieces
        
        ModLogFile.LogAddLine ("write new workpiece " & CStr(AppWPIndex))
        Call ModGUI.ListHMIUpdate("write new workpiece " & CStr(AppWPIndex))
        
        
    ''************************************************************************
    ''    change pocket status acoording to robot request
    ''************************************************************************
    ElseIf (HandShake.RequestChangePocketStatus(0) = 1) Then        ''Byte 118
    
        ModLogFile.LogAddLine (" get request for change pocket status ")
        Call ModGUI.ListHMIUpdate("get request for change pocket status ")
        Call ReadByte(117, ArrayByte())
        
        If ArrayByte(0) = 1 Then
            TempStatus = PocketStatus.empty_
            StringStatus = "Empty"
            
        ElseIf ArrayByte(0) = 2 Then
            TempStatus = PocketStatus.unmachined
            StringStatus = "unmachined"
            
        ElseIf ArrayByte(0) = 3 Then
            TempStatus = PocketStatus.machined
            StringStatus = "machined"
            
        ElseIf ArrayByte(0) = 4 Then
            TempStatus = PocketStatus.reserved
            StringStatus = "reserved"
            
        ElseIf ArrayByte(0) = 5 Then
            TempStatus = PocketStatus.Mask
            StringStatus = "Mask"
            
        ElseIf ArrayByte(0) = 6 Then
            TempStatus = PocketStatus.occupied
            StringStatus = "occupied"
            
        ElseIf ArrayByte(0) = 7 Then
            TempStatus = PocketStatus.Broken
            StringStatus = "Broken"
            
        ElseIf ArrayByte(0) = 8 Then
            TempStatus = PocketStatus.Disable
            StringStatus = "Disable"
            
        End If
        
        Call ModAutomationStatus.SetPocketStatus(TempStatus, AppCurrentPocket)
        ''for the unloading proccess...
        Call ModAutomationStatus.SetPocketWrokPiece(AllWP(AppWPIndex).WPNumber, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketDiameter(AllWP(AppWPIndex).ToolDiameter, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketProgramNumber(AllWP(AppWPIndex).NCProgram, AppCurrentPocket)
        ''
        
        Call fMainForm.BtnDisplayPocketStatus_Click
        
        CycleDone = True
        ResetOneCommByte (118) ''   reset the request
        SetOneCommByte (119)   ''   the status change was done
        
        ModLogFile.LogAddLine (" status changed by robot : " & AppCurrentPocket & " : " & StringStatus)
        Call ModGUI.ListHMIUpdate(" status changed by robot : " & AppCurrentPocket & " : " & StringStatus)
        
        ''************************************************************************
        ''    Delete Work Piece by robot request
        ''************************************************************************
    ElseIf (HandShake.RequestDeleteWPiece(0) = 1) Then
        
        ModLogFile.LogAddLine ("got request to delete workpiece number 1.")
        Call ModGUI.ListHMIUpdate("got request to delete workpiece number 1.")
        Call modAllWorkPiece.WorkPieceReset(1)
        Call modAllWorkPiece.ReOrderAllWPiece
        Call ModCsvFile.SaveAllWP ''         save the data array into file
        Call fMainForm.BtnDisplayWPTable_Click   '     refresh the work piece Table
        ModLogFile.LogAddLine ("Delete work piece number 1 by robot request.")
        Call ModGUI.ListHMIUpdate("Delete work piece number 1 by robot request.")
        ModHandShake.ResetOneCommByte (56)
        ModHandShake.SetOneCommByte (57)
        
        If HermleWPMode = NightMode Then
            If AllWP(1).WPNumber = 0 Then     '''there is no  WorkPiece
                ModHandShake.SetOneCommByte (23)
            ElseIf AllWP(1).WPNumber <> 0 Then '''there is WorkPiece
                ModHandShake.ResetOneCommByte (23)
            End If
        End If
        
    End If
End Sub


Public Sub RunTestsJobs()
''1.the function manage the test cycles.
''2.this function decode the state of the current test.
''  if there is a request for loop Test the function start the test again.
''3.the f is called from the :tmrUpdateRobotStatus_Timer()
''4.legend:
''      1 =idle
''      2 =run
''      3 =done

Dim ret As Double
Dim AllPocketsStep As Integer

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer

Dim shelf_from As Integer
Dim column_from As Integer
Dim pocket_from As Integer

Dim shelf_to As Integer
Dim column_to As Integer
Dim pocket_to As Integer
Dim LineNumber As Integer
    
    If (HermleGear <> Semi) Then
        Exit Sub
    End If
    
    Select Case CurrentTest
    
    Case "3_TEST_KIOSK_TO_POCKET.JBI"
        If ((Tests.KioskToPocket(0) = JobState.IDLE) And (fMainForm.OptionLoop(0).Value = True)) Then
            shelf = CInt(Left(fMainForm.txtTestPocket(2).Text, 1)) ''get info from GUI
            column = CInt(Right(fMainForm.txtTestPocket(2).Text, 2))
            pocket = CInt(fMainForm.TextToolsDiameter2.Text)
            Call BuildArrayPosition(shelf, column, pocket, AppToolType)
            Call MotoComToolBox.WritePosition(16, ArrayPosition) ''send info to controller
            Call SetServo(True)
            Call StartJob(CurrentTest)
        End If
        ''
    Case "3_TEST_POCKET_TO_KIOSK.JBI"
        If ((Tests.PocketToKiosk(0) = JobState.IDLE) And (fMainForm.OptionLoop(1).Value = True)) Then
            shelf = CInt(Left(fMainForm.txtTestPocket(2).Text, 1))  ''get info from GUI
            column = CInt(Right(fMainForm.txtTestPocket(2).Text, 2))
            pocket = CInt(fMainForm.TextToolsDiameter2.Text)
            Call BuildArrayPosition(shelf, column, pocket, AppToolType)
            Call MotoComToolBox.WritePosition(17, ArrayPosition) ''send info to controller
            Call SetServo(True)
            Call StartJob(CurrentTest)
            
        End If
        
    Case "3_TEST_POCKET_TO_CHUCK.JBI"
        If ((Tests.PocketToChuck(0) = JobState.IDLE) And (fMainForm.OptionLoop(2).Value = True)) Then
                shelf = CInt(Left(fMainForm.txtTestPocket(2).Text, 1)) ''get info from GUI
                column = CInt(Right(fMainForm.txtTestPocket(2).Text, 2))
                pocket = CInt(fMainForm.TextToolsDiameter2.Text)
                
                Call BuildArrayPosition(shelf, column, pocket, AppToolType)
                Call MotoComToolBox.WritePosition(16, ArrayPosition) ''send info to controller
                Call SetServo(True)
                Call StartJob(CurrentTest)
                
        End If
        
    Case "3_TEST_CHUCK_TO_POCKET.JBI"
        If ((Tests.ChuckToPocket(0) = JobState.IDLE) And (fMainForm.OptionLoop(3).Value = True)) Then
            shelf = CInt(Left(fMainForm.txtTestPocket(2).Text, 1)) ''get info from GUI
            column = CInt(Right(fMainForm.txtTestPocket(2).Text, 2))
            pocket = CInt(fMainForm.TextToolsDiameter2.Text)
            
            Call BuildArrayPosition(shelf, column, pocket, AppToolType)
            Call MotoComToolBox.WritePosition(17, ArrayPosition) ''send info to controller
            Call SetServo(True)
            Call StartJob(CurrentTest)
            
        End If
        
    Case "3_TEST_POCKET_TO_POCKET.JBI"
        If ((Tests.PocketToPocket(0) = JobState.IDLE) And (fMainForm.CheckLoopP2P.Value = vbChecked)) Then
            shelf_from = CInt(Left(fMainForm.txtTestPocket(0).Text, 1))
            column_from = CInt(Right(fMainForm.txtTestPocket(0).Text, 2))
            pocket_from = CInt(fMainForm.TextToolsDiameter2.Text)
            
            shelf_to = CInt(Left(fMainForm.txtTestPocket(1).Text, 1))
            column_to = CInt(Right(fMainForm.txtTestPocket(1).Text, 2))
            pocket_to = CInt(fMainForm.TextToolsDiameter2.Text)
            
            Call BuildArrayPosition(shelf_from, column_from, pocket_from, AppToolType)
            Call MotoComToolBox.WritePosition(16, ArrayPosition)
            Call ModHandShake.WriteSensorLocation(shelf_from, column_from, pocket_from, 13)
        
            Call BuildArrayPosition(shelf_to, column_to, pocket_to, AppToolType)
            Call MotoComToolBox.WritePosition(17, ArrayPosition)
            Call ModHandShake.WriteSensorLocation(shelf_to, column_to, pocket_to, 20)
            
            Call SetServo(True)
            Call StartJob(CurrentTest)
            
        End If
        
        Case "3_TEST_ALL_POCKETS.JBI"
                If ((Tests.AllPockets(0) = JobState.IDLE) And (fMainForm.CheckLoopAllPockets.Value = vbChecked)) Then
                    AllPocketsColumn = AllPocketsColumn + 1
                    
                    If AppToolType = HSK Then
                        If AllPocketsColumn = TotalHSK Then
                            Call WriteOneJobStatus(Tests.AllPockets(0), Done)
                            CurrentTest = ""
                            RunTestEndless = False
                            Call fMainForm.BtnStopLoadUnload_Click
                            Exit Sub
                        End If
                        
                    ElseIf AppToolType = Drill Then
                        If AllPocketsColumn = TotalDRILL Then
                            Call WriteOneJobStatus(Tests.AllPockets(0), Done)
                            CurrentTest = ""
                            RunTestEndless = False
                            Call fMainForm.BtnStopLoadUnload_Click
                            Exit Sub
                        End If
                        
                    ElseIf AppToolType = Round Then
                        If AllPocketsColumn = TotalROUND Then
                            Call WriteOneJobStatus(Tests.AllPockets(0), Done)
                            CurrentTest = ""
                            RunTestEndless = False
                            Call fMainForm.BtnStopLoadUnload_Click
                            Exit Sub
                        End If
                    End If
                    
                    shelf_from = CInt(Left(fMainForm.TextShelfNumber2.Text, 1))
                    column_from = AllPocketsColumn
                    pocket_from = CInt(fMainForm.TextAllPocketsDiameter.Text)
                    
                    shelf_to = CInt(Left(fMainForm.TextShelfNumber2.Text, 1))
                    column_to = AllPocketsColumn + 1
                    pocket_to = CInt(fMainForm.TextAllPocketsDiameter.Text)
            
                    Call BuildArrayPosition(AllPocketsShelf, AllPocketsColumn, AllPocketsDiameter, AppToolType)
                    Call MotoComToolBox.WritePosition(16, ArrayPosition)
                    Call ModHandShake.WriteSensorLocation(shelf_from, column_from, pocket_from, 13)
                    
                    Call BuildArrayPosition(AllPocketsShelf, AllPocketsColumn + 1, AllPocketsDiameter, AppToolType)
                    Call MotoComToolBox.WritePosition(17, ArrayPosition)
                    Call ModHandShake.WriteSensorLocation(shelf_to, column_to, pocket_to, 20)
                        
                    Call SetServo(True)
                    Call StartJob(CurrentTest)
                
                End If
                
        End Select
End Sub


                    
Public Sub ResetSingleTestStatus(ByVal MyJob As String)
''1.the function set value=IDLE to the test state if the test is done.
''2.the function called from the timer :"tmrUpdateRobotStatus_Timer"
''3.the function get one parameter,JobName, as string and reset that test.
''4.:LEGEND :
        ''1 idle
        ''2 run
        ''3 done

Dim ret As Integer

    ArrayDouble(0) = 1
    Select Case MyJob

 
    Case "3_TEST_KIOSK_TO_POCKET.JBI"
    If Tests.KioskToPocket(0) = JobState.Done Then
        Call WriteByte(42, ArrayDouble())
        Tests.KioskToPocket(0) = JobState.IDLE
    End If


    Case "3_TEST_POCKET_TO_KIOSK.JBI"
    If Tests.PocketToKiosk(0) = JobState.Done Then
        Call WriteByte(43, ArrayDouble())
        Tests.PocketToKiosk(0) = JobState.IDLE
    End If

    Case "3_TEST_POCKET_TO_CHUCK.JBI"
    If Tests.PocketToChuck(0) = JobState.Done Then
        Call WriteByte(44, ArrayDouble())
        Tests.PocketToChuck(0) = JobState.IDLE
    End If

    Case "3_TEST_CHUCK_TO_POCKET.JBI"
     If Tests.ChuckToPocket(0) = JobState.Done Then
        Call WriteByte(45, ArrayDouble())
        Tests.ChuckToPocket(0) = JobState.IDLE
    End If

    Case "3_TEST_POCKET_TO_POCKET.JBI"
     If Tests.PocketToPocket(0) = JobState.Done Then
        Call WriteByte(39, ArrayDouble())
        Tests.PocketToPocket(0) = JobState.IDLE
    End If

    Case "3_TEST_ALL_POCKETS.JBI"
    If Tests.AllPockets(0) = JobState.Done Then
        Call WriteByte(48, ArrayDouble())
        Tests.AllPockets(0) = JobState.IDLE
    End If
        
    End Select
    
''1 idle
''2 run
''3 done
        
End Sub




Public Sub UpdatePocketsStatusByJobState()
''1.this function update the {AutomationStatus(3, 12)} according
''  to the state of the current RobotJob.
''2.the  function also update the CSV File :"AutomationStatus.csv" and update the GUI.
''3.the function called from the timer.
''4.the fun take no parameter
''5.the f return no parameter.



    If Jobs.KioskToPocket(0) = JobState.Done Then
        Call ModAutomationStatus.SetPocketStatus(unmachined, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketWrokPiece(AllWP(AppWPIndex).WPNumber, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketDiameter(AllWP(AppWPIndex).ToolDiameter, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketProgramNumber(AllWP(AppWPIndex).NCProgram, AppCurrentPocket)
        
        
        ''update the neighbors status.
        If (AppToolType = HSK) And (AppDiameter > 150) Then
            Call ModAutomationStatus.SetPocketNeighborsStatus(AppCurrentPocket, occupied)
        End If
    End If
    
    If Jobs.PocketToKiosk(0) = JobState.Done Then
        Call ModAutomationStatus.SetPocketStatus(empty_, AppCurrentPocket)
    End If
    
    If Jobs.PocketToChuck(0) = JobState.Done Then
         Call ModAutomationStatus.SetPocketStatus(reserved, AppCurrentPocket)
         Call ModAutomationStatus.SetPocketWrokPiece(AllWP(AppWPIndex).WPNumber, AppCurrentPocket)
         Call ModAutomationStatus.SetPocketDiameter(AllWP(AppWPIndex).ToolDiameter, AppCurrentPocket)
         Call ModAutomationStatus.SetPocketProgramNumber(AllWP(AppWPIndex).NCProgram, AppCurrentPocket)
    End If
    
    If Jobs.ChuckToPocket(0) = JobState.Done Then
    
        Call ModAutomationStatus.SetPocketWrokPiece(AllWP(AppWPIndex).WPNumber, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketDiameter(AllWP(AppWPIndex).ToolDiameter, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketStatus(machined, AppCurrentPocket)
        Call ModAutomationStatus.SetPocketProgramNumber(AllWP(AppWPIndex).NCProgram, AppCurrentPocket)
    
        ''update the neighbors status.
        If (AppToolType = HSK) And (AppDiameter > 150) Then
            Call ModAutomationStatus.SetPocketNeighborsStatus(AppCurrentPocket, occupied)
        End If
    End If

    Call fMainForm.BtnDisplayPocketStatus_Click
    
''1 idle
''2 run
''3 done
    
    
End Sub



Public Function DisplayCurrentWPCounter() As Boolean

''''1.this function read integer 11 from the robot controller.
''''  the int represent the tool counter in the current WorkPiece.
''''2.the function update the global integer:"AppCurrentWPCounter"
''''3. the function called from the timer.
''''4.the function beeing called from the timer:tmrUpdateRobotStatus_Timer().
''
''
''Dim bb As Boolean
''
''    bb = ReadInteger(11, ArrayInteger())
''    If bb = True Then
''        AppCurrentWPCounter = ArrayInteger(0)
''        ''         tool left             =      all tools in Wp              - Tools been proccessed
''        ''fMainForm.TextAmountLeft.Caption = CStr(AllWP(AppWPIndex).ToolAmount - AppCurrentWPCounter)
''        DisplayCurrentWPCounter = True
''    Else
''        DisplayCurrentWPCounter = False
''        ''Call FrmDialog.ShowDialogForm(40, 40, 40, "ModHandShake", "DisplayCurrentWPCounter()", GeneralError)
''        Exit Function
''    End If
''
End Function


Public Sub ReadAllJobsStatus()
''1.the function read from robot the state of all JOBS.
''2.the function put the value into Jobs.something in the memory.
''15 X 30 [ms] = 0.45 sec

''1 idle
''2 run
''3 done
Dim jj As Integer

    Call ReadByte(29, Jobs.AutoStart())
    Call ReadByte(33, Jobs.Parking())
    Call ReadByte(34, Jobs.Exchange())
'    Call ReadByte(35, Jobs.KioskToPocket())
'    Call ReadByte(36, Jobs.PocketToKiosk())
    Call ReadByte(37, Jobs.PocketToChuck())
    Call ReadByte(38, Jobs.ChuckToPocket())
    Call ReadByte(46, Jobs.KioskToPocket())
    Call ReadByte(47, Jobs.PocketToKiosk())
    Call ReadByte(137, Jobs.TakeFromPocket())
    Call ReadByte(138, Jobs.PlaceOnPocket())
    Call ReadByte(139, Jobs.TakeFromKiosk())
    Call ReadByte(140, Jobs.PlaceOnKiosk())
    Call ReadByte(141, Jobs.TakeFromChuck())
    Call ReadByte(142, Jobs.PlaceOnChuck())
    
End Sub


Public Sub ResetJobStatus()
''1.the function reset one job
''2.the f reset the job that its state is 'Done'
''3.the f called from:"tmrUpdateRobotStatus_Timer()"
''4.the f take no parameter .
''5.the function return no parameters.
    
    If Jobs.AutoStart(0) = 3 Then
        Call WriteOneCommByte(29, JobState.IDLE)
        CycleDone = False
        
    ElseIf Jobs.Parking(0) = 3 Then
        Call WriteOneCommByte(33, JobState.IDLE)
        CycleDone = False
        
    ElseIf Jobs.Exchange(0) = 3 Then
        Call WriteOneCommByte(34, JobState.IDLE)
        CycleDone = False
        
    ElseIf Jobs.KioskToPocket(0) = 3 Then
        Call WriteOneCommByte(46, JobState.IDLE)
        CycleDone = False
        
    ElseIf Jobs.PocketToKiosk(0) = 3 Then
        Call WriteOneCommByte(47, JobState.IDLE)
        CycleDone = False
        
    ElseIf Jobs.PocketToChuck(0) = 3 Then
        Call WriteOneCommByte(37, JobState.IDLE)
        CycleDone = False
        
    ElseIf Jobs.ChuckToPocket(0) = 3 Then
        Call WriteOneCommByte(38, JobState.IDLE)
        CycleDone = False
        
    ''fMainForm.txtCycleMessage(1).Caption = "Reset one Job Status()"
        
    
    
    End If
    
''1 IDLE
''2 RUN
''3 DONE

End Sub
 Public Sub ResetAllProccess()
''1.the function reset the state  of all JOBS and TESTS.
''2.[29-38] jobs
''3.[39-45] Tests

 Dim jj As Integer
 
     If HermleCommState = OffLine Then
        Exit Sub
    End If
 
    For jj = 29 To 47
        ResetOneCommByte (jj) ''all the proccess
    Next
    
    ResetOneCommByte (65) ''the MaskUser
        
    Jobs.AutoStart(0) = 0
    Jobs.ChuckToPocket(0) = 0
    Jobs.Exchange(0) = 0
    Jobs.KioskToPocket(0) = 0
    Jobs.Parking(0) = 0
    Jobs.PocketToChuck(0) = 0
    Jobs.PocketToKiosk(0) = 0
    
    Tests.AllPockets(0) = 0
    Tests.ChuckToPocket(0) = 0
    Tests.KioskToPocket(0) = 0
    Tests.PocketToChuck(0) = 0
    Tests.PocketToKiosk(0) = 0
    Tests.PocketToPocket(0) = 0
    
 End Sub

Public Sub BuildArrayPosition _
   (shelf As Integer, _
        column As Integer, _
            pocket As Integer, _
                Mytooltype As tooltype)
''1.the function build a position to be send to the robot.
''2.the function get 3 arguments.
''3.the functioin build (return) the array - a global variable.

Dim digit As String

    If Mytooltype = Drill Then
        If (column > 12) Then
            Exit Sub
        End If
        ArrayPosition(0) = 1 ''1=XYZ
        ArrayPosition(1) = 1 ''1=robot cordinate system
        ArrayPosition(2) = DrillLocations(shelf, column).diameter(pocket).X
        ArrayPosition(3) = DrillLocations(shelf, column).diameter(pocket).Y
        ArrayPosition(4) = DrillLocations(shelf, column).diameter(pocket).z
        ArrayPosition(5) = DrillLocations(shelf, column).diameter(pocket).Rx
        ArrayPosition(6) = DrillLocations(shelf, column).diameter(pocket).Ry
        ArrayPosition(7) = DrillLocations(shelf, column).diameter(pocket).Rz

    ElseIf Mytooltype = HSK Then
        If (column > 10) Then
            Exit Sub
        End If
        ArrayPosition(0) = 1 ''1=XYZ
        ArrayPosition(1) = 1 ''1=robot cordinate system
        ArrayPosition(2) = HSKLocations(shelf, column).X
        ArrayPosition(3) = HSKLocations(shelf, column).Y
        ArrayPosition(4) = HSKLocations(shelf, column).z
        ArrayPosition(5) = HSKLocations(shelf, column).Rx
        ArrayPosition(6) = HSKLocations(shelf, column).Ry
        ArrayPosition(7) = HSKLocations(shelf, column).Rz
        
    ElseIf Mytooltype = Round Then
        If (column > 12) Then
            Exit Sub
        End If
        ArrayPosition(0) = 1 ''1=XYZ
        ArrayPosition(1) = 1 ''1=robot coordinate system
        ArrayPosition(2) = RoundLocations(shelf, column).diameter(pocket).X
        ArrayPosition(3) = RoundLocations(shelf, column).diameter(pocket).Y
        ArrayPosition(4) = RoundLocations(shelf, column).diameter(pocket).z
        ArrayPosition(5) = RoundLocations(shelf, column).diameter(pocket).Rx
        ArrayPosition(6) = RoundLocations(shelf, column).diameter(pocket).Ry
        ArrayPosition(7) = RoundLocations(shelf, column).diameter(pocket).Rz
        
        ''correct the Z parameter.
        If pocket <= 4 Then
            PocketDepth = 86
        ElseIf pocket >= 5 Then
            PocketDepth = 92
        Else
            PocketDepth = 92
        End If
        
        ArrayPosition(4) = RoundLocations(shelf, column).diameter(pocket).z - _
            PocketDepth + ChuckDepth - ChuckStopper + PocketStopper + ShelfSafty

    End If
    
    If column < 10 Then
        digit = "0"
    End If
    
    ModLogFile.LogAddLine ("Build pocket number : " & shelf & digit & column)
    Call ModGUI.ListHMIUpdate("Build pocket number : " & shelf & digit & column)
    
End Sub

Public Sub PauseRobotJob(ByVal jobname As String)

Dim ret As Integer

    If HermleCommState = OffLine Then
        Exit Sub
    End If

    ret = BscSelectJob(m_nCid, jobname)
    ret = BscHoldOn(m_nCid)
    ret = BscHoldOff(m_nCid)
    
    ModLogFile.LogAddLine (" Pause robot job : " & jobname)
    Call ModGUI.ListHMIUpdate(" Pause robot job : " & jobname)
    
End Sub



Public Sub ResetProfibus()

Dim bb As Boolean

    bb = StartJob("RESET_PROFIBUS.JBI")
    If bb = False Then
        Call FrmDialog.ShowDialogForm(51, 51, 51, "form communication", " ResetProfibus_Click()", GeneralError)
        Exit Sub
    End If
    
    ModLogFile.LogAddLine (" Reset profibus")
    Call ModGUI.ListHMIUpdate("Reset profibus")

End Sub

Public Function WriteDrillCode(MyDrillCode As Integer) As Boolean

Dim bb As Boolean

    ArrayInteger(0) = CDbl(MyDrillCode)
    bb = WriteInteger(41, ArrayInteger())
    
    If bb = False Then
        WriteDrillCode = False
    ElseIf bb = True Then
        WriteDrillCode = True
    End If
    
    ModLogFile.LogAddLine (" Write DrillCode to Integer 41. DrillCode=" & MyDrillCode)
    Call ModGUI.ListHMIUpdate("Write DrillCode to Integer 41. DrillCode=" & MyDrillCode)
        
End Function
Public Function SendToolSensorState() As Boolean
''1.this function send byte to robot
''2.the f set if Do/Don't use the ToolPresence Sensor.
''3.the function write into byte number 116 in the controller.
''4.the f is beeing called 6 times.

Dim bb As Boolean
    
    If UseToolSensor = True Then
        ArrayByte(0) = 1
    ElseIf UseToolSensor = False Then
        ArrayByte(0) = 0
    End If
        
    bb = WriteByte(116, ArrayByte())
    
    If bb = False Then
        SendToolSensorState = False
    ElseIf bb = True Then
        SendToolSensorState = True
    End If
    
    ModLogFile.LogAddLine ("send ToolSensor state to byte 116. ToolSensor=" & CStr(ArrayByte(0)))
    Call ModGUI.ListHMIUpdate("send ToolSensor state to byte 116. ToolSensor=" & CStr(ArrayByte(0)))

End Function
Public Sub ResetRobotJob(ByVal jobname As String)

Dim ret As Integer

    ret = BscSelectJob(m_nCid, jobname)
    ''0=found job    ''other number=error
    If ret <> 0 Then
        Exit Sub
    End If
    
    ret = BscHoldOn(m_nCid)
    ret = BscHoldOff(m_nCid)
    
    ret = BscSetLineNumber(m_nCid, 0)
    ''0:      Normal completion
    ''Others:  Error codes
    If (ret <> 0) Then
        Exit Sub
    End If
    
End Sub


Public Sub ReadFirstLoadUnload()
    
''1.this function read the state of the LOAD Cycle and the UNLOAD Cycle
'   from the Kiosk to pocket and back.
''2.the function Determine if it is the first LOAD or not.
''3.the function Determine if it is the first Unload or not.


Dim TempArray(8) As Double

    If HermleCommState = OffLine Then
        Exit Sub
    End If

    Call MotoComToolBox.ReadByte(60, TempArray())
    If TempArray(0) = 1 Then
        LoadFirstTime = True
    ElseIf TempArray(0) = 0 Then
        LoadFirstTime = False
    End If
    
    Call MotoComToolBox.ReadByte(61, TempArray())
    If TempArray(0) = 1 Then
        UnloadFirstTime = True
    ElseIf TempArray(0) = 0 Then
        UnloadFirstTime = False
    End If

End Sub



Public Function WriteSensorLocation _
       (ByVal MyShelf As Integer, _
        ByVal MyColumn As Integer, _
        ByVal MyDiameter As Integer, _
        ByVal PositionNumber As Integer _
        ) As Boolean
''1.this function write position to robot.
''2.the position where the robot should stand,so the Sensor 'see'
''  if there is or there is not a tool in the current pocket.
Dim fdbk As Integer
On Error GoTo label

Dim SentIndex As Integer

    If UseToolSensor = False Then
        WriteSensorLocation = True
        ModLogFile.LogAddLine ("did not send sensor location to robot:UseToolSensor = False")
        Call ModGUI.ListHMIUpdate(" did not send sensor location to robot:UseToolSensor = False")
        Exit Function
    End If
    
    If AppToolType = HSK Then
        ''SentIndex = 1
        Exit Function
        
    ElseIf AppToolType = Drill Then
        If (MyDiameter >= 1) And (MyDiameter <= 4) Then
            SentIndex = 4
        ElseIf MyDiameter >= 5 And MyDiameter <= 7 Then
            SentIndex = 7
        End If
            
    ElseIf AppToolType = Round Then
        If (MyDiameter <= 1) And (MyDiameter <= 4) Then
            SentIndex = 4
        ElseIf MyDiameter >= 5 And MyDiameter <= 8 Then
            SentIndex = 8
        End If
    End If
    
    Call ModHandShake.BuildArrayPosition(MyShelf, MyColumn, SentIndex, AppToolType)
    Call MotoComToolBox.WritePosition(PositionNumber, ArrayPosition())
    
    WriteSensorLocation = True
    ModLogFile.LogAddLine ("send sensor location to robot with drill code : " & SentIndex)
    Call ModGUI.ListHMIUpdate(" send sensor location to robot with drill code : " & SentIndex)
    
Exit Function
label:

    fdbk = MsgBox("can not Send SensorPosition to controller " & vbCrLf & "the error is :" & Err.Description, _
        vbCritical, "WriteSensorLocation()")
End Function


Public Function ReadToolCounter() As Boolean
''1.the sub read tool counter from robot.
''2.the sub read integer number 11.
''3.the function update the AmountLeft Counter.

On Error GoTo label

    ReadToolCounter = False
    
    If HermleCommState = OffLine Then
        ReadToolCounter = True
        Exit Function
    End If
    
    Call ReadInteger(11, ArrayInteger())
    ToolCounter = ArrayInteger(0)
   
    ReadToolCounter = True
    Exit Function
    
label:
    ReadToolCounter = False
    
End Function


Public Function ReadRobotAlarm() As Integer

Dim ret As Integer
Dim Data As Integer
Dim Msg As String
Dim AlarmString As String * 36

    If HermleCommState = OffLine Then
        Exit Function
    End If
    
    ret = BscIsAlarm(m_nCid)
   
    If ret = 1 Then ''1 there is alarm
    
        ret = BscReadAlarmS(m_nCid, Data, AlarmString)
        ret = BscGetFirstAlarmS(m_nCid, Data, AlarmString)
        RobotAlarmString = AlarmString
        ReadRobotAlarm = 1
        Exit Function
    
    ElseIf ret = 0 Then  ''0 there is no alarm
    
        RobotAlarmString = "Robot Run.No Alarm."
        ReadRobotAlarm = 0
        Exit Function
        
    End If
    
End Function

Public Function NoZeroValues(ArrayPosition) As Boolean
 ''check if there is no 0 value in X,Y,Z
 
Dim kk As Integer
    NoZeroValues = False

     For kk = 2 To 4
         If (ArrayPosition(kk) = 0) Then
            NoZeroValues = False
            Exit Function
         End If
     Next
     
     NoZeroValues = True
     
End Function

Public Function ResetAllRobotRequest() As Boolean

On Error GoTo error
    ResetAllRobotRequest = False
    
    Call ModHandShake.ResetOneCommByte(11)
    Call ModHandShake.ResetOneCommByte(13)
    Call ModHandShake.ResetOneCommByte(15)
    Call ModHandShake.ResetOneCommByte(17)
    Call ModHandShake.ResetOneCommByte(21)
    Call ModHandShake.ResetOneCommByte(56)
    Call ModHandShake.ResetOneCommByte(65)
    Call ModHandShake.ResetOneCommByte(118)
    

    HandShake.RequestPlaceFromKiosk(0) = 0
    HandShake.RequestTakeToKiosk(0) = 0
    HandShake.RequestPlaceFromChuck(0) = 0
    HandShake.RequestTakeToChuck(0) = 0
 
    
    ResetAllRobotRequest = True
    Exit Function
    
error:
    ResetAllRobotRequest = False
End Function


Public Sub ReadRobotStrings()
    
    If Jobs.AutoStart(0) = JobState.RUN Then
        ReadString (11)
        RobotStrings(11) = ArrayString
    End If
    
    If Jobs.KioskToPocket(0) = JobState.RUN Then
        ReadString (14)
        RobotStrings(14) = ArrayString
    End If
    
    If Jobs.PocketToKiosk(0) = JobState.RUN Then
        ReadString (15)
        RobotStrings(15) = ArrayString
    End If
    
    If Jobs.TakeFromPocket(0) = JobState.RUN Then
        ReadString (18)
        RobotStrings(18) = ArrayString
    End If
    
    If Jobs.PlaceOnPocket(0) = JobState.RUN Then
        ReadString (19)
        RobotStrings(19) = ArrayString
    End If
    
    If Jobs.TakeFromKiosk(0) = JobState.RUN Then
        ReadString (16)
        RobotStrings(16) = ArrayString
    End If
    
    If Jobs.PlaceOnKiosk(0) = JobState.RUN Then
        ReadString (17)
        RobotStrings(17) = ArrayString
    End If
    
    If Jobs.TakeFromChuck(0) = JobState.RUN Then
        ReadString (20)
        RobotStrings(20) = ArrayString
    End If
    
    If Jobs.PlaceOnChuck(0) = JobState.RUN Then
        ReadString (21)
        RobotStrings(21) = ArrayString
    End If
    
End Sub


