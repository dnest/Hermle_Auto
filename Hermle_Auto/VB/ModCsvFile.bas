Attribute VB_Name = "ModCsvFile"
Option Explicit

Public Sub SaveArray(ByVal ArrayName As String)

Dim CsvPath As String
Dim filenumber As Integer
Dim Today As Date
Dim now As Date
Dim country As String
Dim Factory As String
Dim MachineNumber As Integer
Dim i As Integer
Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim ret As Integer
    
On Error GoTo error

        CsvPath = App.path & "\WorkDirectory\Data\"
        CsvPath = CsvPath & ArrayName & ".csv"
        filenumber = FreeFile
        Open CsvPath For Output As #filenumber
        ''
        Print #filenumber, "Pocket Name", ",", "x", ",", "y", ",", "z", ",", "Rx", ",", "Ry", ",", "Rz", ",", "Distance", ",", "Alfa"
        
        If ArrayName = "HSKLocations" Then
            For shelf = 1 To 3
                For column = 1 To 10
                    Print #filenumber, _
                         HSKLocations(shelf, column).name, _
                    ",", HSKLocations(shelf, column).X, _
                    ",", HSKLocations(shelf, column).Y, _
                    ",", HSKLocations(shelf, column).z, _
                    ",", HSKLocations(shelf, column).Rx, _
                    ",", HSKLocations(shelf, column).Ry, _
                    ",", HSKLocations(shelf, column).Rz, _
                    ",", HSKLocations(shelf, column).Dist, _
                    ",", HSKLocations(shelf, column).Alfa
                Next
            Next
        End If
        
        If ArrayName = "DrillLocations" Then
                    
            For shelf = 1 To 3
                For column = 1 To 12
                    For pocket = 1 To 7
                        Print #filenumber, _
                             DrillLocations(shelf, column).diameter(pocket).name, _
                        ",", DrillLocations(shelf, column).diameter(pocket).X, _
                        ",", DrillLocations(shelf, column).diameter(pocket).Y, _
                        ",", DrillLocations(shelf, column).diameter(pocket).z, _
                        ",", DrillLocations(shelf, column).diameter(pocket).Rx, _
                        ",", DrillLocations(shelf, column).diameter(pocket).Ry, _
                        ",", DrillLocations(shelf, column).diameter(pocket).Rz, _
                        ",", DrillLocations(shelf, column).diameter(pocket).Dist, _
                        ",", DrillLocations(shelf, column).diameter(pocket).Alfa
                    Next
                Next
            Next
            
            
        End If
        
        If ArrayName = "RoundLocations" Then
                    
            For shelf = 1 To 3
                For column = 1 To 12
                    For pocket = 1 To 8
                        Print #filenumber, _
                             RoundLocations(shelf, column).diameter(pocket).name, _
                        ",", RoundLocations(shelf, column).diameter(pocket).X, _
                        ",", RoundLocations(shelf, column).diameter(pocket).Y, _
                        ",", RoundLocations(shelf, column).diameter(pocket).z, _
                        ",", RoundLocations(shelf, column).diameter(pocket).Rx, _
                        ",", RoundLocations(shelf, column).diameter(pocket).Ry, _
                        ",", RoundLocations(shelf, column).diameter(pocket).Rz, _
                        ",", RoundLocations(shelf, column).diameter(pocket).Dist, _
                        ",", RoundLocations(shelf, column).diameter(pocket).Alfa
                    Next
                Next
            Next
            
            
        End If
        
        Close filenumber
        Exit Sub
error:
       ret = MsgBox("Error while saving locations" & vbCrLf & _
        "Type : " & ArrayName & vbCrLf _
            & " the error is  : " & Err.Description _
            , vbExclamation, "Modcsvfile.SaveArray()")
End Sub

Public Sub SaveAutomation()
''1.the function save the AutomationStatus into the HardDisk.
''2.the function save the status of every pocket in the automation.
 Debug.Print "SaveAutomation"
''
Dim AutomationPath As String
Dim filenumber As Integer
Dim Today As Date
Dim now As Date
Dim country As String
Dim Factory As String
Dim MachineNumber As Integer
Dim i As Integer
Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim ret As Integer

On Error GoTo error

    If AppToolType = HSK Then
        AutomationPath = App.path & "\WorkDirectory\Data\"
        AutomationPath = AutomationPath & "HSKStatus.csv"
        filenumber = FreeFile
        Open AutomationPath For Output As #filenumber
        ''
        Print #filenumber, "Pocket Name", ",", "shelf", ",", "column", ",", "pocket", ",", "diameter", ",", "CurrentTool", ",", "Status", ",", "WorkPiece", ",", "programNumber"
        
        For shelf = 1 To 3
            For column = 1 To 10
                Print #filenumber _
                , AutomationStatus(shelf, column).name, _
                ",", AutomationStatus(shelf, column).shelf, _
                ",", AutomationStatus(shelf, column).column, _
                ",", AutomationStatus(shelf, column).pocket, _
                ",", AutomationStatus(shelf, column).diameter, _
                ",", AutomationStatus(shelf, column).CurrentTool, _
                ",", AutomationStatus(shelf, column).Status, _
                ",", AutomationStatus(shelf, column).WorkPiece _
                ; ",", AutomationStatus(shelf, column).ProgramNumber
            Next
        Next
        Close filenumber
    End If
    
        If AppToolType = Drill Then
        AutomationPath = App.path & "\WorkDirectory\Data\"
        AutomationPath = AutomationPath & "DrillStatus.csv"
        filenumber = FreeFile
        Open AutomationPath For Output As #filenumber
        ''
        Print #filenumber, "Pocket Name", ",", "shelf", ",", "column", ",", "pocket", ",", "diameter", ",", "CurrentTool", ",", "Status", ",", "WorkPiece", ",", "programNumber"
        
        For shelf = 1 To 3
            For column = 1 To 12
                Print #filenumber _
                , AutomationStatus(shelf, column).name, _
                ",", AutomationStatus(shelf, column).shelf, _
                ",", AutomationStatus(shelf, column).column, _
                ",", AutomationStatus(shelf, column).pocket, _
                ",", AutomationStatus(shelf, column).diameter, _
                ",", AutomationStatus(shelf, column).CurrentTool, _
                ",", AutomationStatus(shelf, column).Status, _
                ",", AutomationStatus(shelf, column).WorkPiece _
                ; ",", AutomationStatus(shelf, column).ProgramNumber
            Next
        Next
        Close filenumber
    End If
    
    If AppToolType = Round Then
        AutomationPath = App.path & "\WorkDirectory\Data\"
        AutomationPath = AutomationPath & "RoundStatus.csv"
        filenumber = FreeFile
        Open AutomationPath For Output As #filenumber
        ''
        Print #filenumber, "Pocket Name", ",", "shelf", ",", "column", ",", "pocket", ",", "diameter", ",", "CurrentTool", ",", "Status", ",", "WorkPiece", ",", "programNumber"
        
        For shelf = 1 To 3
            For column = 1 To 12
                Print #filenumber _
                , AutomationStatus(shelf, column).name, _
                ",", AutomationStatus(shelf, column).shelf, _
                ",", AutomationStatus(shelf, column).column, _
                ",", AutomationStatus(shelf, column).pocket, _
                ",", AutomationStatus(shelf, column).diameter, _
                ",", AutomationStatus(shelf, column).CurrentTool, _
                ",", AutomationStatus(shelf, column).Status, _
                ",", AutomationStatus(shelf, column).WorkPiece _
                ; ",", AutomationStatus(shelf, column).ProgramNumber
            Next
        Next
        Close filenumber
    End If
    
    Exit Sub
error:
ret = MsgBox("error while saving status automation." _
        & vbCrLf & "tool type  is :" & AppToolType & vbCrLf _
            & "the error is : " & Err.Description _
            , vbExclamation, "Modcsvfile.SaveAutomation()")
    
  
End Sub


Public Sub SaveAllWP()

Dim WorkPiecePath As String
Dim filenumber As Integer
Dim Today As Date
Dim now As Date
Dim country As String
Dim Factory As String
Dim MachineNumber As Integer
Dim i As Integer
Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim ret As Integer

On Error GoTo error:

        WorkPiecePath = App.path & "\WorkDirectory\Data\"
        WorkPiecePath = WorkPiecePath & "ALLWorkPiece.csv"
        filenumber = FreeFile
        Open WorkPiecePath For Output As #filenumber
        Print #filenumber, "Line Number", ",", " WorkPieceNumber", ",", "NCProgram", ",", "Tool Amount", ",", "ToolAmountLeft", ",", "ToolDiameter", ",", "WP ToolType"
        
        For i = 1 To 50
            Print #filenumber, _
            AllWP(i).LineNumber _
            ; ",", AllWP(i).WPNumber _
            ; ",", AllWP(i).NCProgram _
            ; ",", AllWP(i).ToolAmount _
            ; ",", AllWP(i).ToolAmountLeft _
            ; ",", AllWP(i).ToolDiameter _
            ; ",", AllWP(i).WPToolType

        Next
        Close filenumber
    
Exit Sub
error:
    ret = MsgBox("Error while saving AllWork Pieces." & vbCrLf _
            & "The error is : " & Err.Description _
                , vbExclamation, " ModcsvFile.SaveAllWP()")
End Sub

Public Sub ReadPocketsLocations(ArrayName As String)

Dim StringArray(4, 13, 8) As String
Dim file1 As String
Dim fname As String
Dim ret As Integer
Dim TempDate  As String
Dim TempUpdate As String
Dim TempCountry As String
Dim TempFactory As String
Dim TempMachineNumber As String
Dim tempArrayName As String
Dim dummy As String

Dim shelf As Integer
Dim column As Integer
Dim diameter As Integer
Dim ii As Integer

On Error GoTo label

    If ArrayName = "DrillLocations" Then
    
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\DrillLocations.csv"
        Open fname For Input As #file1
        
        Line Input #file1, dummy  ''read the entire line''Read the header

        For shelf = 1 To 3
            For column = 1 To TotalDRILL
                For diameter = 1 To 7
                    Input #file1, dummy   ''read just one cell
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).X
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).Y
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).z
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).Rx
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).Ry
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).Rz
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).Dist
                    Input #file1, DrillLocations(shelf, column).diameter(diameter).Alfa
                Next
            Next
            
        Next
            Close #file1
    End If

    
    If ArrayName = "HSKLocations" Then
    
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\HSKlocations.csv"
        Open fname For Input As #file1
        Line Input #file1, dummy  ''          read the entire line
        For shelf = 1 To 3
            For column = 1 To TotalHSK
                Input #file1, dummy   ''      read just one cell
                Input #file1, HSKLocations(shelf, column).X
                Input #file1, HSKLocations(shelf, column).Y
                Input #file1, HSKLocations(shelf, column).z
                Input #file1, HSKLocations(shelf, column).Rx
                Input #file1, HSKLocations(shelf, column).Ry
                Input #file1, HSKLocations(shelf, column).Rz
                Input #file1, HSKLocations(shelf, column).Dist
                Input #file1, HSKLocations(shelf, column).Alfa
            Next
        Next
        Close #file1
    End If
    
    If ArrayName = "RoundLocations" Then
    
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\RoundLocations.csv"
        Open fname For Input As #file1
        Line Input #file1, dummy  ''read the entire line''Read the header
        For shelf = 1 To 3
            For column = 1 To TotalROUND
                For diameter = 1 To 8
                    Input #file1, dummy   ''      read just one cell
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).X
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).Y
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).z
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).Rx
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).Ry
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).Rz
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).Dist
                    Input #file1, RoundLocations(shelf, column).diameter(diameter).Alfa
                Next diameter
            Next
        Next
        Close #file1
    End If
    
    Exit Sub
    
label:
    ret = MsgBox("Error while loading Pocket Locaions." _
    & vbCrLf & "try to load : " & ArrayName & vbCrLf _
        & "the error is :" & Err.Description _
            , vbExclamation, "ReadPocketsLocations()")
    
End Sub



Public Sub WriteGeneralLocations()

''1.this function write all general locations from memory into file.

Dim filenumber As Integer
Dim Today As Date
Dim now As Date
Dim country As String
Dim Factory As String
Dim MachineNumber As Integer
Dim ret As Integer
Dim ii As Integer
Dim GeneralLocationPath As String
Dim file1 As String
Dim fname As String

On Error GoTo error

    If AppToolType = HSK Then
        fname = App.path & "\WorkDirectory\Data\HSKGeneralLocations.csv"
        
    ElseIf AppToolType = Drill Then
        fname = App.path & "\WorkDirectory\Data\DrillGeneralLocations.csv"
        
    ElseIf AppToolType = Round Then
        fname = App.path & "\WorkDirectory\Data\RoundGeneralLocations.csv"
    End If
    
        filenumber = FreeFile
        Open fname For Output As #filenumber
        ''
        Print #filenumber, "General Location", ",", "x", ",", "y", ",", "z", ",", "Rx", ",", "Ry", ",", "Rz"
        
        For ii = 1 To AmountOfGeneralLocations
            GeneralLocation(ii).name = CStr(ii)
        Next
        
        For ii = 1 To AmountOfGeneralLocations
            Print #filenumber, _
             GeneralLocation(ii).name _
            ; ",", GeneralLocation(ii).X _
            ; ",", GeneralLocation(ii).Y _
            ; ",", GeneralLocation(ii).z _
            ; ",", GeneralLocation(ii).Rx _
            ; ",", GeneralLocation(ii).Ry _
            ; ",", GeneralLocation(ii).Rz
        Next
        Close filenumber
Exit Sub
error:
ret = MsgBox("Error while saving general locations " & vbCrLf _
        & "the error is :" & Err.Description _
            , vbExclamation, "modcsvfile.WriteGeneralLocations()")
End Sub


Public Sub ReadGeneralLocation()
''
''
Debug.Print "ReadGeneralLocation()"
''
Dim StringArray(4, 13, 8) As String
Dim file1 As String
Dim fname As String
Dim ret As Integer
Dim TempDate  As String
Dim TempUpdate As String
Dim TempCountry As String
Dim TempFactory As String
Dim TempMachineNumber As String
Dim tempArrayName As String
Dim dummy As String

Dim shelf As Integer
Dim column As Integer
Dim diameter As Integer
Dim ii As Integer

    On Error GoTo label
    
    If AppToolType = HSK Then
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\HSKGeneralLocations.csv"
        
    ElseIf AppToolType = Drill Then
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\DrillGeneralLocations.csv"
        
    ElseIf AppToolType = Round Then
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\RoundGeneralLocations.csv"
    End If
    
    Open fname For Input As #file1

    Line Input #file1, dummy  ''read one line.read the entire header.''Read the header
 
    For ii = 1 To AmountOfGeneralLocations
    
        Input #file1, dummy   ''                read just one cell-the name of the position
        Input #file1, GeneralLocation(ii).X ''  read just one cell
        Input #file1, GeneralLocation(ii).Y ''  read just one cell
        Input #file1, GeneralLocation(ii).z ''  read just one cell
        Input #file1, GeneralLocation(ii).Rx '' read just one cell
        Input #file1, GeneralLocation(ii).Ry '' read just one cell
        Input #file1, GeneralLocation(ii).Rz '' read just one cell

    Next

    Close #file1
    
    Exit Sub
    
    
label:
    ret = MsgBox("Error while reading General Locations" & vbCrLf _
            & " the error is :" & Err.Description & vbCrLf _
            & "the system was reading " & ii - 1 & " positions ." _
                , vbExclamation, "ModCsvFile.ReadGeneralLocation()")

End Sub

Public Sub ReadAutomationStatus()

Debug.Print "ReadAutomationStatus()"

Dim file1 As String
Dim fname As String
Dim ret As Integer
Dim shelf As Integer
Dim column As Integer
Dim diameter As Integer

Dim ii As Integer
Dim dummy As String
    
    On Error GoTo label
    ReadAllWorkPiece
    If AppToolType = HSK Then
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\HSKStatus.csv"
        Open fname For Input As #file1
        Line Input #file1, dummy  ''read the entire line
        For shelf = 1 To 3
            For column = 1 To 10
                    Input #file1, AutomationStatus(shelf, column).name '''read one cell
                    Input #file1, AutomationStatus(shelf, column).shelf
                    Input #file1, AutomationStatus(shelf, column).column
                    Input #file1, AutomationStatus(shelf, column).pocket
                    Input #file1, AutomationStatus(shelf, column).diameter
                    Input #file1, AutomationStatus(shelf, column).CurrentTool
                    Input #file1, AutomationStatus(shelf, column).Status
                    Input #file1, AutomationStatus(shelf, column).WorkPiece
                    Input #file1, AutomationStatus(shelf, column).ProgramNumber
            Next
        Next
        Close #file1
    End If
    
    If AppToolType = Drill Then
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\DrillStatus.csv"
        Open fname For Input As #file1
        Line Input #file1, dummy  ''read the entire line
        For shelf = 1 To 3
            For column = 1 To 12
                    Input #file1, AutomationStatus(shelf, column).name '''read one cell
                    Input #file1, AutomationStatus(shelf, column).shelf
                    Input #file1, AutomationStatus(shelf, column).column
                    Input #file1, AutomationStatus(shelf, column).pocket
                    Input #file1, AutomationStatus(shelf, column).diameter
                    Input #file1, AutomationStatus(shelf, column).CurrentTool
                    Input #file1, AutomationStatus(shelf, column).Status
                    Input #file1, AutomationStatus(shelf, column).WorkPiece
                    Input #file1, AutomationStatus(shelf, column).ProgramNumber
            Next
            
        Next
        Close #file1
    End If
    
    If AppToolType = Round Then
        file1 = FreeFile
        fname = App.path & "\WorkDirectory\Data\RoundStatus.csv"
        Open fname For Input As #file1
        Line Input #file1, dummy  ''read the entire line
        For shelf = 1 To 3
            For column = 1 To 12
                    Input #file1, AutomationStatus(shelf, column).name '''read one cell
                    Input #file1, AutomationStatus(shelf, column).shelf
                    Input #file1, AutomationStatus(shelf, column).column
                    Input #file1, AutomationStatus(shelf, column).pocket
                    Input #file1, AutomationStatus(shelf, column).diameter
                    Input #file1, AutomationStatus(shelf, column).CurrentTool
                    Input #file1, AutomationStatus(shelf, column).Status
                    Input #file1, AutomationStatus(shelf, column).WorkPiece
                    Input #file1, AutomationStatus(shelf, column).ProgramNumber
            Next
            
        Next
        Close #file1
    End If
        
    Exit Sub
label:
    
    ret = MsgBox("Error while reading automation status " _
    & vbCrLf & "tool type is : " & AppToolType _
        & vbCrLf & "the error is : " & Err.Description _
            , vbExclamation, "ReadAutomationStatus()")

End Sub




Public Sub ReadAllWorkPiece()
''1.this function read all work piece data from file into the memory.
''2.the function read data from CSV file into the : AllWP() array.
''3.the function called from the 'ModMain' too.
Debug.Print "ReadAllWorkPiece()"
 
Dim ret As Integer
Dim file1 As String
Dim fname As String
Dim TempDate  As String
Dim TempUpdate As String
Dim TempCountry As String
Dim TempFactory As String
Dim TempMachineNumber As String
Dim tempArrayName As String
Dim dummy As String
Dim ii As Integer

    On Error GoTo label
    file1 = FreeFile
    fname = App.path & "\WorkDirectory\Data\ALLWorkPiece.csv"
    Open fname For Input As #file1
    Line Input #file1, dummy  ''read the entire line
    For ii = 1 To 50
        Input #file1, AllWP(ii).LineNumber   ''read just one cell
        Input #file1, AllWP(ii).WPNumber
        Input #file1, AllWP(ii).NCProgram
        Input #file1, AllWP(ii).ToolAmount
        Input #file1, AllWP(ii).ToolAmountLeft
        Input #file1, AllWP(ii).ToolDiameter
        Input #file1, AllWP(ii).WPToolType
    Next
    Close #file1
    
    Exit Sub
label:
    ret = MsgBox("error while reading work piece file  " _
        & vbCrLf & "the error is :" & Err.Description _
            , vbExclamation, "ReadAllWorkPiece()")
    
End Sub




