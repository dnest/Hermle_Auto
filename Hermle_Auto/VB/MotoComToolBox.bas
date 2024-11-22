Attribute VB_Name = "MotoComToolBox"
Option Explicit


Dim m_nCid As Integer
Dim open_mode As Integer
Dim strIPaddr As String
Dim ether_mode As Integer
Dim hWnd As Long
Dim ret As Long
    
Public Function WriteByte(ByVal ByteNumber As Integer, ByRef ArrayData() As Double) As Boolean

    If HermleCommState = OffLine Then
        WriteByte = True
        Exit Function
    End If

    ret = BscPutVarData(m_nCid, 0, ByteNumber, ArrayData(0))

    If UseHMILogger = True Then
        With fMainForm.lstTransmitToRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Byte - " & ByteNumber & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If

    If ret = 0 Then
        WriteByte = True
    Else
        WriteByte = False: fMainForm.lblDllWriteError = Val(fMainForm.lblDllWriteError) + 1
    End If

'0     :  Normal completion
'Others:  Error codes
End Function

Public Function WriteInteger(ByVal IntegerNumber As Integer, ByRef ArrayData() As Double) As Boolean

Dim vardata(0 To 10) As Double
Dim strVal As String

    If HermleCommState = OffLine Then
        WriteInteger = True
        Exit Function
    End If
    
    ret = BscPutVarData(m_nCid, 1, IntegerNumber, ArrayData(0))

    If UseHMILogger = True Then
        With fMainForm.lstTransmitToRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Integer - " & IntegerNumber & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If
   
    If ret = 0 Then
        WriteInteger = True
    Else
        WriteInteger = False: fMainForm.lblDllWriteError = Val(fMainForm.lblDllWriteError) + 1
    End If

'0     :  Normal completion
'Others:  Error codes
End Function

Public Function WriteDouble(ByVal DoubleIndex As Integer, ByRef ArrayData() As Double) As Boolean

    If HermleCommState = OffLine Then
        WriteDouble = True
        Exit Function
    End If
    
    ret = BscPutVarData(m_nCid, 2, DoubleIndex, ArrayData(0))

    If UseHMILogger = True Then
        With fMainForm.lstTransmitToRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Double - " & DoubleIndex & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If
   
    If ret = 0 Then
        WriteDouble = True
    Else
        WriteDouble = False: fMainForm.lblDllWriteError = Val(fMainForm.lblDllWriteError) + 1
    End If

'0:  Normal completion
'Others:  Error codes
End Function


Public Function ReadByte(ByVal ByteNumber As Double, ByRef ArrayData() As Double) As Boolean

    If HermleCommState = OffLine Then
        ReadByte = True
        Exit Function
    End If
    
    ret = BscGetVarData(m_nCid, 0, ByteNumber, ArrayData(0))

    If UseHMILogger = True Then
        With fMainForm.lstReceivedFromRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Byte - " & ByteNumber & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If
    
   If ByteNumber = 17 And ArrayData(0) = 1 Then
        ByteNumber = ByteNumber
    End If
    If ret = 0 Then
       ReadByte = True
    Else
        ReadByte = False: fMainForm.lblDllWriteError = Val(fMainForm.lblDllWriteError) + 1
    End If

''0:  Normal completion
''Others:  Error codes
End Function

Public Sub SetCommunicationParameters()
    
        'ethernet communication
        ''open_mode = 16     'ethernet BSC mode
        open_mode = 256   'ethernet E-Server mode
        strIPaddr = "192.168.100.10" '' the IP of the Controller
        ether_mode = 0    'for host function client-mode is neccessary
        hWnd = fMainForm.hWnd         'handle of dialog window

End Sub


Public Function ReadInteger(ByVal IntegerNumber As Double, ByRef ArrayData() As Double) As Boolean
      
    If HermleCommState = OffLine Then
        ReadInteger = True
        Exit Function
    End If
    
    ret = BscGetVarData(m_nCid, 1, IntegerNumber, ArrayData(0))

    If UseHMILogger = True Then
        With fMainForm.lstReceivedFromRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Integer - " & IntegerNumber & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If
                
    If ret = 0 Then
        ReadInteger = True
    Else
       ReadInteger = False: fMainForm.lblDllReadError = Val(fMainForm.lblDllReadError) + 1
    End If
           
End Function


Public Function ReadPosition(ByVal PositionNumber As Double, ByRef ArrayData() As Double) As Boolean
''1.the function read position from controller.
''  NOT the current location !
''2.the function get the position number to be read
''3.the function return True or False if the function success or fail.

    If HermleCommState = OffLine Then
        ReadPosition = True
        Exit Function
    End If
    
    ret = BscGetVarData(m_nCid, 4, PositionNumber, ArrayData(0))
    
    If UseHMILogger = True Then
        With fMainForm.lstReceivedFromRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Position - " & PositionNumber & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If
       
    If ret = 0 Then
        ReadPosition = True
    Else
        ReadPosition = False: fMainForm.lblDllReadError = Val(fMainForm.lblDllReadError) + 1
    End If
           
End Function

Public Function WritePosition(ByVal PositionNumber As Double, ByRef ArrayData() As Double) As Boolean
''1.the function write one position to the controller.
''2.the function get the position number to be written.
''3.the function return True or False if the function success or fail.

    If HermleCommState = OffLine Then
        WritePosition = True
        Exit Function
    End If

    If AppToolType = HSK Then
        ArrayData(9) = tooltype.HSK
    ElseIf AppToolType = Drill Then
        ArrayData(9) = tooltype.Drill
    ElseIf AppToolType = Round Then
        ArrayData(9) = tooltype.Round
    End If
    
    ArrayData(9) = 0
    
    ret = BscPutVarData(m_nCid, 4, PositionNumber, ArrayData(0))
    
    If UseHMILogger = True Then
        With fMainForm.lstTransmitToRobot
            .AddItem Format(now, "hh:mm:ss") & " Bad - " & ret & " Position - " & PositionNumber & " Value - " & ArrayData(0)
            If .TopIndex > 32000 Then
                .Clear
            Else
                .TopIndex = .ListCount - 1
            End If
        End With
    End If
   
    If ret = 0 Then
        WritePosition = True
    Else
        WritePosition = False: fMainForm.lblDllWriteError = Val(fMainForm.lblDllWriteError) + 1
    End If
        
    Call ModLogFile.LogAddLine(" send position : " & CStr(PositionNumber))
    Call ModGUI.ListHMIUpdate("send position : " & CStr(PositionNumber))
         
End Function
Public Function ReadFile(ByVal FileName As String) As Boolean

    If HermleCommState = OffLine Then
        ReadFile = True
        Exit Function
    End If
    
    ret = BscUpLoad(m_nCid, UCase$(CStr(FileName)))
    
    If ret = 0 Then
        ReadFile = True
    Else
        ReadFile = False
    End If
    
''0:      Normal completion
''Others:  Receiving error
           
End Function

Public Function WriteFile(ByVal FileName As String) As Boolean

    If HermleCommState = OffLine Then
        WriteFile = True
        Exit Function
    End If
    
    ret = BscDownLoad(m_nCid, FileName)
    
    If ret = 0 Then
        WriteFile = True
    Else
        WriteFile = False
    End If
           

End Function
    

Public Sub CloseCommunication()

        ret = BscDisConnect(m_nCid)
        ret = BscClose(m_nCid)

End Sub
Public Function ReadAllFiles(ByRef ArrayData() As String) As Boolean

Dim temp As String
Dim i As Integer

    i = 0

    If HermleCommState = OffLine Then
        ReadAllFiles = True
        Exit Function
    End If
    
    ret = BscFindFirst(m_nCid, temp, 50)
    ret = BscFindFirst(m_nCid, ArrayData(i), 50)
    If ret = 0 Then
        Do While ret <> -1
            i = i + 1
            ret = BscFindNext(m_nCid, ArrayData(i), 100)
            If ret <> 0 Then
                Exit Do
            End If
        Loop
    Else
    Call CloseCommunication
    End If

End Function
Public Function ReadCurrentPosition(ByRef ArrayData() As Double) As Boolean
''1.the function read the current position into 'ArrayData()'

Dim strVal As String
Dim frame As String
Dim External As Integer
Dim RobotConfig As Integer
Dim ToolNumber As Integer
Dim ret As Integer

    frame = "ROBOT"
    External = 0
    
    If HermleCommState = OffLine Then
        ReadCurrentPosition = True
        Exit Function
    End If
    
    ret = BscIsRobotPos(m_nCid, frame, External, RobotConfig, ToolNumber, ArrayData(0))
    
    If ret = 0 Then
        ReadCurrentPosition = True
    Else
        ReadCurrentPosition = False
    End If
           
''-1 : Acquisition Failure
'' 0 : Normal completion

End Function


Public Function ReadServoStatus() As Boolean

Dim ret As Integer

    If HermleCommState = OffLine Then
        ReadServoStatus = True
        Exit Function
    End If
    
    ReadServoStatus = False
         
    ret = BscIsServo(m_nCid)
    
    If ret = 0 Then
        ReadServoStatus = False
    Else
        ReadServoStatus = True
    End If
           
End Function

Public Function StartJob(ByVal jobname As String) As Boolean

Dim TempInt As Integer

    If HermleCommState = OffLine Then
        StartJob = True
        Exit Function
    End If
    

    ret = BscSelectJob(m_nCid, jobname)
    ''0=found job
    ''other=error
    If ret <> 0 Then
        StartJob = False
        Exit Function
    End If
    
    TempInt = BscHoldOn(m_nCid)
    TempInt = BscHoldOff(m_nCid)
    
    ret = BscSetLineNumber(m_nCid, 0)
    ''0:      Normal completion
    ''Others:  Error codes
    If (ret <> 0) Then
        StartJob = False
        Exit Function
    End If

    ret = BscStartJob(m_nCid)
        ''0      : Normal completion
        ''1      : Current job name not specified
        ''Others :  Error codes
    If ret = 0 Then
        StartJob = True
        Exit Function
    End If
    If ret <> 0 Then
       StartJob = False
       Exit Function
    End If
    
    Call ModLogFile.LogAddLine("Start Job : " & jobname)
    Call ModGUI.ListHMIUpdate(" Start Job : " & jobname)
    
End Function

Public Function SetServo(ByVal ServoState As Boolean) As Integer

Dim ret As Integer

    If HermleCommState = OffLine Then
        Exit Function
    End If
    
    If ServoState = True Then
        ret = BscServoOn(m_nCid)
        SetServo = ret
    Else
        ret = BscServoOff(m_nCid)
        SetServo = ret
    End If

End Function


Public Function IncrementalMove(ByVal Velocity As Double, ByVal Index As Integer)

Dim ret As Integer

    ret = BscImov(m_nCid, "V", Velocity, "ROBOT", 1, IncrementalTarget(0))
   ' re = BscImov(Commun,Cnst, Speed   , Coord  ,Tool,Position)
      
End Function




Public Function StartCommunication()

'step 1: get a hardware key handle


    If HermleCommState = OffLine Then
        Exit Function
    End If
    
    m_nCid = BscOpen(App.path, open_mode)

    If m_nCid >= 0 Then
        
        ret = BscSetEServer(m_nCid, strIPaddr)  ''  Eserver mode

        If ret = 1 Then
            'step 3: Establish a connection
            ret = BscConnect(m_nCid)
            If ret = 1 Then
                '...
            Else
                ret = BscDisConnect(m_nCid)
                ret = BscClose(m_nCid)
                m_nCid = -1
                MsgBox ("Error establish connection !")
            End If
        Else
                ret = BscDisConnect(m_nCid)
                ret = BscClose(m_nCid)
            m_nCid = -1
            MsgBox ("Error setting up ethernet !")
        End If
    Else
        MsgBox ("Hardware Key Error !")
    End If
    
End Function


Public Function ReadKeyState() As Integer

Dim ret As Integer

    If HermleCommState = OffLine Then
        ReadKeyState = KeyState.remote
        Exit Function
    End If
    

    ret = BscIsTeachMode(m_nCid)
    
        If (ret) = -1 Then
            ReadKeyState = KeyState.error_
            AppKeyState = error_
            
        ElseIf (ret) = 0 Then
            ReadKeyState = KeyState.teach
            AppKeyState = teach
            
        ElseIf (ret) = 1 Then
            ReadKeyState = KeyState.Play
            AppKeyState = Play
            
        End If
        
        ret = BscIsRemoteMode(m_nCid)
        If ret = 1 Then
            ReadKeyState = KeyState.remote
            AppKeyState = remote
            
        ElseIf (ret = -1) Then
            
        End If
        
End Function




Public Function ReadLineNumber() As Integer

Dim ret As Integer

    If HermleCommState = OffLine Then
        ReadLineNumber = 0
        Exit Function
    End If
    

    ret = BscIsJobLine(m_nCid)
    ReadLineNumber = ret
    
End Function




Public Function ReadJobName() As String

Dim mySize As Integer
Dim MyJobName As String
Dim ret As Integer

    If HermleCommState = OffLine Then
        ReadJobName = " Off Line "
        Exit Function
    End If
    

    MyJobName = "CurrentRobotJobName"
    ret = BscIsJobName(m_nCid, MyJobName, 100)
    
    If ret = 0 Then
        ReadJobName = MyJobName
    End If
    
    If ret = -1 Then
        ReadJobName = "Error Name"
    End If
    
    
    
''-1 : Acquisition Failure
'' 0 :  Normal completion
    
End Function




Public Function WriteAllGeneralLocationToRobot()
''1.this function being called from:
''  ButtonRestoreGeneralLocations_Click() ONLY !!

Dim kk As Double

    If HermleCommState = OffLine Then
        Exit Function
    End If
    
        kk = 10
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
        
        kk = 11
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
        
              
        kk = 12
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
                     
        kk = 21
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
        
        kk = 22
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
               
        kk = 23
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
               
        kk = 24
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
               
        kk = 25
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
               
        kk = 26
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
               
        kk = 120
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
               
        kk = 121
        ArrayPosition(0) = 1
        ArrayPosition(1) = 1
        ArrayPosition(2) = GeneralLocation(kk).X
        ArrayPosition(3) = GeneralLocation(kk).Y
        ArrayPosition(4) = GeneralLocation(kk).z
        ArrayPosition(5) = GeneralLocation(kk).Rx
        ArrayPosition(6) = GeneralLocation(kk).Ry
        ArrayPosition(7) = GeneralLocation(kk).Rz
        Call WritePosition(kk, ArrayPosition())
       
End Function


Public Function TestCommunication() As Boolean

Dim ret As Integer

    If HermleCommState = OffLine Then
        Exit Function
    End If
    

    ret = BscGetVarData(m_nCid, DataType.ByteType, 1, ArrayByte(0))

    If ret = 0 Then
        TestCommunication = True
    Else
        TestCommunication = False
    End If
    
''0:      Normal completion
''Others:  Error codes
    
End Function


Public Function ReadString(ByVal StringIndex As Integer) As Boolean

Dim ret As Integer

    If HermleCommState = OffLine Then
        ReadString = True
        Exit Function
    End If
    

    ret = BscHostGetVarData(m_nCid, 7, StringIndex, ArrayDouble(0), ArrayString)
    
''the ArrayDouble is not in use in this case

    
    If ret = 0 Then
        ReadString = True
    Else
        ReadString = False
    End If
    
''0:      Normal completion
''Others:  Error codes
     
End Function


Public Function WriteString(ByVal StringIndex As Integer, ByVal MyString As String) As Boolean

Dim ret As Integer


    If HermleCommState = OffLine Then
        WriteString = True
        Exit Function
    End If
    

    ret = BscHostPutVarData(m_nCid, 7, StringIndex, ArrayDouble(0), MyString)
    
    If ret = 0 Then
        WriteString = True
    Else
        WriteString = False
    End If
    
''0:  Normal completion
''Others:  Error codes
     
End Function


Public Function ReadIO(ByVal IOIndex As Integer) As Integer

Dim ret As Integer
Dim IOStatus As Integer


    If HermleCommState = OffLine Then
        Exit Function
    End If
    
     ret = BscReadIO2(m_nCid, IOIndex, 1, IOStatus)
     
     ''the system was able to read the status.
     If ret = 0 Then
        If IOStatus = 0 Then
            ReadIO = 0
        Else
            ReadIO = 1
        End If
     End If
     
     
     ''the system was NOT able to read the status
     If ret <> 0 Then
        ReadIO = -1
     End If
     
     
''-1 : Header number error
''0:  Normal completion
''Others:  Error codes

 '''(ByVal nCid%, ByVal startadd&, ByVal ionum%, stat%) As Integer
End Function

Public Function ReadStepNumber() As Integer

Dim ret As Integer

    If HermleCommState = OffLine Then
        ReadStepNumber = 0
        Exit Function
    End If
    

    ret = BscIsJobStep(m_nCid)
    ReadStepNumber = ret
    
End Function

Public Function WriteIO(ByVal OutputState As Integer, OutputAddress As Double)

Dim ret As Integer

     ''ret = BscWriteIO2(m_nCid, OutputAddress, 1, OutputState)

End Function


''***--***--**************************************************************
''    ByteType = 0
''    DoubleType = 2
''    Real = 3
''    Position = 4
''    Base Position = 5




''                vardata(0) = CDbl(1)                                            'data type (XYZ)
''                vardata(1) = CDbl(1)                                            'Coordinate Type
''                vardata(2) = CDbl(frmParameters.txtRobPtsPars(0).Text)          'X
''                vardata(3) = CDbl(frmParameters.txtRobPtsPars(1).Text)          'Y
''                vardata(4) = CDbl(frmParameters.txtRobPtsPars(2).Text)          'Z
''                vardata(5) = CDbl(frmParameters.txtRobPtsPars(3).Text)          'Rx
''                vardata(6) = CDbl(frmParameters.txtRobPtsPars(4).Text)          'Ry
''                vardata(7) = CDbl(frmParameters.txtRobPtsPars(5).Text)          'Rz
















