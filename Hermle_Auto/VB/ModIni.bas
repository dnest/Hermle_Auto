Attribute VB_Name = "ModIni"
Option Explicit

Public Sub ReadAppToolType()
''1.the function read the Application tooltype from ini file.
''2.the function assign the return value into the global var:AppToolType
Debug.Print "ReadAppToolType()"
''
Dim TempString As String * 2
Dim LongRet As Long
Dim ret As Integer

    On Error GoTo error
    LongRet = GetPrivateProfileString _
    ("application", "AppToolType", "HSK", TempString, 2, App.path & "\WorkDirectory\data\IS2904.ini")
    
    ''MsgBox "TempString=" & TempString
    TempString = Left(TempString, LongRet)
    
    If TempString = "1 " Then
        AppToolType = HSK
    ElseIf TempString = "2 " Then
        AppToolType = Drill
    ElseIf TempString = "3 " Then
        AppToolType = Round
    Else
        GoTo error
    End If

    ''MsgBox "apptooltype=" & AppToolType
         
''other = 0
''HSK = 1
''Drill = 2
''Round = 3
    Exit Sub
error:
    ret = MsgBox("Error while reading tooltype." & vbCrLf _
        & "the error is : " & "tool not defined or file not found." _
            , vbExclamation, "ReadAppToolType()")
  
End Sub


Public Sub WriteAppToolType(ByVal TempToolType As String)
''1.the function write the application tooltype into ini file.
''2.the function get one parameter:string
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString _
        ("application", "AppToolType", TempToolType, App.path & "\WorkDirectory\data\IS2904.ini")
           
''other = 0
''HSK = 1
''Drill = 2
''Round = 3
    
  
End Sub


Public Sub ReadShelvsConfig()
''1.this function read shelvs configuration from ini file.
''2.the function Assign The data into array :AppShelvs(jj)
''3. the f beeing called from :
''          CmdDisplayShelvs_Click()
''          frmoptions.Form_Load()
Debug.Print "ReadShelvsConfig()"
''
Dim TempInt As Integer
Dim TempFirst As String * 2
Dim TempSecond As String * 2
Dim TempThird As String * 2
Dim LongRet As Long

    LongRet = GetPrivateProfileString _
        ("shelvs", "first", "Round", TempFirst, 2, App.path & "\WorkDirectory\data\IS2904.ini")
        
    LongRet = GetPrivateProfileString _
        ("shelvs", "second", "Round", TempSecond, 2, App.path & "\WorkDirectory\data\IS2904.ini")
        
    LongRet = GetPrivateProfileString _
        ("shelvs", "third", "HSK", TempThird, 2, App.path & "\WorkDirectory\data\IS2904.ini")
    
    TempFirst = Left(TempFirst, LongRet)
    TempSecond = Left(TempSecond, LongRet)
    TempThird = Left(TempThird, LongRet)
    
    If TempFirst = "3 " Then
        AppShelvs(1).ShelfToolType = Round
        AppShelvs(1).ShelfName = "Round"
    ElseIf TempFirst = "2 " Then
        AppShelvs(1).ShelfToolType = Drill
        AppShelvs(1).ShelfName = "Drill"
    ElseIf TempFirst = "1 " Then
        AppShelvs(1).ShelfToolType = HSK
        AppShelvs(1).ShelfName = "HSK"
    ElseIf TempFirst = "0 " Then
        AppShelvs(1).ShelfToolType = other
        AppShelvs(1).ShelfName = "Not In Use"
    End If
    
    If TempSecond = "3 " Then
        AppShelvs(2).ShelfToolType = Round
        AppShelvs(2).ShelfName = "Round"
    ElseIf TempSecond = "2 " Then
        AppShelvs(2).ShelfToolType = Drill
        AppShelvs(2).ShelfName = "Drill"
    ElseIf TempSecond = "1 " Then
        AppShelvs(2).ShelfToolType = HSK
        AppShelvs(2).ShelfName = "HSK"
    ElseIf TempSecond = "0 " Then
        AppShelvs(2).ShelfToolType = other
        AppShelvs(2).ShelfName = "Not In Use"
    End If
    
    If TempThird = "3 " Then
        AppShelvs(3).ShelfToolType = Round
         AppShelvs(3).ShelfName = "Round"
    ElseIf TempThird = "2 " Then
        AppShelvs(3).ShelfToolType = Drill
        AppShelvs(3).ShelfName = "Drill"
    ElseIf TempThird = "1 " Then
        AppShelvs(3).ShelfToolType = HSK
        AppShelvs(3).ShelfName = "HSK"
    ElseIf TempThird = "0 " Then
        AppShelvs(3).ShelfToolType = other
        AppShelvs(3).ShelfName = "Not In Use"
    End If
    
End Sub

Public Function SaveMachineParameters() As Boolean

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString _
        ("information", "Country", HermleAutomation.LocationCountry, App.path & "\WorkDirectory\data\IS2904.ini")
    
    BoolRet = WritePrivateProfileString _
        ("information", "factory", HermleAutomation.LocationFactory, App.path & "\WorkDirectory\data\IS2904.ini")
           
    BoolRet = WritePrivateProfileString _
        ("information", "AutoName", HermleAutomation.AutomationName, App.path & "\WorkDirectory\data\IS2904.ini")
           
    BoolRet = WritePrivateProfileString _
        ("information", "AutoNumber", HermleAutomation.AutomationNumber, App.path & "\WorkDirectory\data\IS2904.ini")
           
    BoolRet = WritePrivateProfileString _
        ("information", "HermleNumber", HermleAutomation.HermleNumber, App.path & "\WorkDirectory\data\IS2904.ini")
           
    BoolRet = WritePrivateProfileString _
        ("information", "HermleType", HermleAutomation.HermleType, App.path & "\WorkDirectory\data\IS2904.ini")
           
    SaveMachineParameters = BoolRet
End Function


Public Function ReadShelvsOffset() As Boolean
''
''1.this function read the Offsets of the shelvs from the first shelf.
''2.the function assign the values to the global vars :"ShelfOffset()"

Dim TempString As String * 4
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("ShelvsOffset", "First", "0", TempString, 4, FilePath)
    TempString = Left(TempString, LongRet)
    ShelfOffset(1) = CInt(TempString)
    
    LongRet = GetPrivateProfileString _
            ("ShelvsOffset", "Second", "50", TempString, 4, FilePath)
    TempString = Left(TempString, LongRet)
    ShelfOffset(2) = CInt(TempString)
    
    LongRet = GetPrivateProfileString _
            ("ShelvsOffset", "Thierd", "100", TempString, 4, FilePath)
    TempString = Left(TempString, LongRet)
    ShelfOffset(3) = CInt(TempString)
    
    

End Function


Public Function ReadToolSensorState() As Boolean

Dim TempString As String * 2
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("application", "ToolSensor", "1", TempString, 2, FilePath)
    TempString = Left(TempString, LongRet)
    
    If TempString = "1 " Then
        UseToolSensor = True
    ElseIf TempString = "0 " Then
        UseToolSensor = False
    End If
    
End Function


Public Function ReadAbovePocket() As Boolean

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("offsets", "AbovePocket", "20", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    AbovePocket = CDbl(TempString)
    
'
'Public AboveChuck As Double
'Public AbovePocket As Double
'

End Function
Public Function ReadAboveChuck() As Boolean

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("offsets", "AboveChuck", "20", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    AboveChuck = CDbl(TempString)
    
End Function

Public Sub WriteAbovePocket()
''1.the function write the OffSet Above Pocket
''2.the function get no parameter
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString("offsets", "AbovePocket" _
                , AbovePocket, App.path & "\WorkDirectory\data\IS2904.ini")
End Sub

Public Sub WriteAboveChuck()
''1.the function write the OffSet Above chuck
''2.the function get no parameter
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString("offsets", "AboveChuck" _
                , AboveChuck, App.path & "\WorkDirectory\data\IS2904.ini")
           
End Sub

Public Function ReadGripperStyle() As Boolean

Dim TempString As String * 2
Dim LongRet As Long
Dim FilePath As String

On Error GoTo error
    ReadGripperStyle = False
    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    LongRet = GetPrivateProfileString _
            ("gripper", "style", "1", TempString, 2, FilePath)
    TempString = Left(TempString, LongRet)
    AppGripperStyle = CInt(TempString)
    ReadGripperStyle = True
    Exit Function
error:
 ReadGripperStyle = False
End Function

Public Function ReadSimulatorState() As Boolean

Dim TempString As String * 2
Dim LongRet As Long
Dim FilePath As String
Dim TempInt As Integer

On Error GoTo error
    ReadSimulatorState = False
    
    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    LongRet = GetPrivateProfileString _
            ("application", "simulator", "1", TempString, 2, FilePath)
            
    TempString = Left(TempString, LongRet)
    TempInt = CInt(TempString)
    
    If TempInt = 1 Then
        AppSimulation = True
    ElseIf TempInt = 0 Then
        AppSimulation = False
    Else
        AppSimulation = False
    End If
    
    ReadSimulatorState = True
    Exit Function
error:
 ReadSimulatorState = False
End Function

Public Sub ReadChuckStopper()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("offsets", "ChuckStopper", "20", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    ChuckStopper = CDbl(TempString)

End Sub

Public Sub ReadChuckDepth()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("offsets", "ChuckDepth", "20", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    ChuckDepth = CDbl(TempString)

End Sub

Public Sub ReadPocketStopper()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("offsets", "PocketStopper", "20", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    PocketStopper = CDbl(TempString)

End Sub
Public Sub ReadKioskStopper()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("offsets", "KioskStopper", "20", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    KioskStopper = CDbl(TempString)

End Sub


Public Sub WriteChuckStopper()
''1.the function write the OffSet Above chuck
''2.the function get no parameter
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString("offsets", "ChuckStopper" _
                , ChuckStopper, App.path & "\WorkDirectory\data\IS2904.ini")
           
End Sub

Public Sub WriteKioskStopper()
''1.the function write the OffSet Above chuck
''2.the function get no parameter
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString("offsets", "KioskStopper" _
                , KioskStopper, App.path & "\WorkDirectory\data\IS2904.ini")
           
End Sub

Public Sub WritePocketStopper()
''1.the function write the OffSet Above chuck
''2.the function get no parameter
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString("offsets", "PocketStopper" _
                , PocketStopper, App.path & "\WorkDirectory\data\IS2904.ini")
           
End Sub

Public Sub WriteChuckDepth()
''1.the function write the OffSet Above chuck
''2.the function get no parameter
''3.the f return no parameter.

Dim BoolRet As Boolean

    BoolRet = WritePrivateProfileString("offsets", "ChuckDepth" _
                , ChuckDepth, App.path & "\WorkDirectory\data\IS2904.ini")
           
End Sub


Public Sub ReadUseExternalFile()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String
Dim TempDouble As Double

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("Documentation", "UseExternalFile", "0", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    TempDouble = CDbl(TempString)
    
    If TempDouble = 1 Then
        UseExternalFile = True
    ElseIf TempDouble = 0 Then
        UseExternalFile = False
    Else
        TempDouble = False
    End If
    
End Sub

Public Sub ReadUseHMILogger()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String
Dim TempDouble As Double

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("Documentation", "UseHMILogger", "0", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    TempDouble = CDbl(TempString)
    
    If TempDouble = 1 Then
        UseHMILogger = True
    ElseIf TempDouble = 0 Then
        UseHMILogger = False
    Else
        UseHMILogger = False
    End If

End Sub



Public Sub ReadUseHMIInfo()

Dim TempString As String * 3
Dim LongRet As Long
Dim FilePath As String
Dim TempDouble As Double

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("Documentation", "UseHMIInfo", "0", TempString, 3, FilePath)
            
    TempString = Left(TempString, LongRet)
    TempDouble = CDbl(TempString)
    
    If TempDouble = 1 Then
        UseHMIInfo = True
    ElseIf TempDouble = 0 Then
        UseHMIInfo = False
    Else
        UseHMIInfo = False
    End If
    
End Sub


Public Sub ReadSawState()

Dim TempString As String * 2
Dim LongRet As Long
Dim FilePath As String

    FilePath = App.path & "\WorkDirectory\data\IS2904.ini"
    
    LongRet = GetPrivateProfileString _
            ("application", "saw", "0", TempString, 2, FilePath)
    TempString = Left(TempString, LongRet)
    
    If TempString = "1 " Then
        UseSaw = True
    ElseIf TempString = "0 " Then
        UseSaw = False
    End If
    
End Sub








