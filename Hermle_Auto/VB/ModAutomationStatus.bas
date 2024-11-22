Attribute VB_Name = "ModAutomationStatus"
Option Explicit

Public Sub AutomationStatusInitialize()
''1. the function initilize the parameters of the status array
''2. the function take no parameters.
''3. the function return nothing.

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim digit As String

Dim LastColumn As Integer

    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If
        
        
    For shelf = 1 To 3
        For column = 1 To LastColumn
            AutomationStatus(shelf, column).shelf = shelf
            AutomationStatus(shelf, column).column = column
            AutomationStatus(shelf, column).pocket = 1
            If column <= 9 Then
                digit = "0"
            ElseIf column >= 10 Then
                digit = ""
            End If
            AutomationStatus(shelf, column).name = CStr(shelf) & digit & CStr(column)
        Next
    Next

 Call SaveAutomation
 
End Sub

Public Sub SetAutomationDefault()

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer

Dim LastColumn As Integer

    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If

    For shelf = 1 To 3
        For column = 1 To LastColumn
            AutomationStatus(shelf, column).Status = empty_
            AutomationStatus(shelf, column).WorkPiece = 0
            AutomationStatus(shelf, column).ProgramNumber = 0
            AutomationStatus(shelf, column).diameter = 1
            
            If AppToolType = Drill Then
                AutomationStatus(shelf, column).CurrentTool = Drill
            ElseIf AppToolType = HSK Then
                AutomationStatus(shelf, column).CurrentTool = HSK
            End If
                
        Next
    Next
    
    Call SaveAutomation

End Sub

Public Sub SetAutomationCurrentTool(ByVal Tool As tooltype)

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim LastColumn As Integer


    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If
        
    For shelf = 1 To 3
        If AppShelvs(shelf).ShelfToolType = AppToolType Then
            For column = 1 To LastColumn
                AutomationStatus(shelf, column).CurrentTool = Tool
            Next
        End If
    Next
    
    Call SaveAutomation
    
End Sub


Public Sub SetAutomationStatus(ByVal MyStatus As PocketStatus)

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim LastColumn As Integer


    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If

    For shelf = 1 To 3
        For column = 1 To LastColumn
            AutomationStatus(shelf, column).Status = MyStatus
        Next
    Next
    
    Call SaveAutomation

End Sub

Public Sub SetAutomationWorkPiece(ByVal MyWorkPiece As Integer)

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim LastColumn As Integer


    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If

    For shelf = 1 To 3
        For column = 1 To LastColumn
            AutomationStatus(shelf, column).WorkPiece = MyWorkPiece
        Next
    Next
    
    Call SaveAutomation

End Sub

Public Sub SetPocketStatus(ByVal SinglePocketStatus As PocketStatus, ByVal PocketName As String)

Dim shelf As Integer
Dim column As Integer

    If (PocketName = "") Then
         Exit Sub
    End If
    
    shelf = CInt(Left(PocketName, 1))
    column = CInt(Right(Left(PocketName, 3), 2))
    AutomationStatus(shelf, column).Status = SinglePocketStatus
    
    Call SaveAutomation
End Sub

Public Sub SetPocketDiameter(ByVal MyDiameter As Double, ByVal PocketName As String)

Dim shelf As Integer
Dim column As Integer
    
    If (PocketName = "") Then
         Exit Sub
    End If
    

    If AppToolType = HSK Then ''make sure the input value is correct.
        If column > TotalHSK Then
            Exit Sub
        End If
    ElseIf AppToolType = Drill Then ''make sure the input value is correct.
        If column > TotalDRILL Then
            Exit Sub
        End If
    ElseIf AppToolType = Round Then ''make sure the input value is correct.
        If column > TotalROUND Then
            Exit Sub
        End If
    End If
    
    shelf = CInt(Left(PocketName, 1))
    column = CInt(Right(Left(PocketName, 3), 2))
    AutomationStatus(shelf, column).diameter = MyDiameter
    
    Call SaveAutomation

End Sub
Public Sub SetPocketCurrentTool(ByVal myCurrentTool As Double, ByVal shelf As Integer, ByVal column As Integer)


    If shelf > 3 Then ''make sure the input value is correct.
        Exit Sub
    End If

    If AppToolType = HSK Then ''make sure the input value is correct.
        If column > TotalHSK Then
            Exit Sub
        End If
    ElseIf AppToolType = Drill Then ''make sure the input value is correct.
        If column > TotalDRILL Then
            Exit Sub
        End If
    End If
    
    AutomationStatus(shelf, column).CurrentTool = myCurrentTool
    
    ''Call SaveAutomation

End Sub

Public Sub SetPocketWrokPiece(ByVal MyWorkPiece As Double, ByVal PocketName As String)
''1 this function set WorkPiece number for specific Pocket.
''2 the function get 2 parameters:
''  the pocket and the the workpiece number.
''3 the function return no parameter.
''4 the function save the AutomationStatus Array into file.
''5 the function update the array in the memory.

Dim shelf As Integer
Dim column As Integer

    If (PocketName = "") Then
         Exit Sub
    End If
    

    If AppToolType = HSK Then ''make sure the input value is correct.
        If column > TotalHSK Then
            Exit Sub
        End If
    ElseIf AppToolType = Drill Then ''make sure the input value is correct.
        If column > TotalDRILL Then
            Exit Sub
        End If
    ElseIf AppToolType = Round Then ''make sure the input value is correct.
        If column > TotalROUND Then
            Exit Sub
        End If
    End If
        
    shelf = CInt(Left(PocketName, 1))
    column = CInt(Right(Left(PocketName, 3), 2))
    AutomationStatus(shelf, column).WorkPiece = MyWorkPiece
    
    Call SaveAutomation

End Sub
Public Sub SetPocketProgramNumber(ByVal MyProgramNumber As Double, ByVal PocketName As String)

Dim shelf As Integer
Dim column As Integer


    If (PocketName = "") Then
         Exit Sub
    End If
    
    
    If AppToolType = HSK Then       ''make sure the input value is correct.
        If column > TotalHSK Then
            Exit Sub
        End If
    ElseIf AppToolType = Drill Then ''make sure the input value is correct.
        If column > TotalDRILL Then
            Exit Sub
        End If
    ElseIf AppToolType = Round Then ''make sure the input value is correct.
        If column > TotalROUND Then
            Exit Sub
        End If
    End If
    
    shelf = CInt(Left(PocketName, 1))
    column = CInt(Right(Left(PocketName, 3), 2))
    AutomationStatus(shelf, column).ProgramNumber = MyProgramNumber
    
    Call SaveAutomation

End Sub


Public Sub SetAutomationProgramNumber(ByVal MyProgramNumber As Double)

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim LastColumn As Integer


    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If
    
    For shelf = 1 To 3
        For column = 1 To LastColumn
            AutomationStatus(shelf, column).ProgramNumber = MyProgramNumber
        Next
    Next
    
    Call SaveAutomation

End Sub


Public Sub SetAutomationDiameter(ByVal MyDiameter As Double)

Dim shelf As Integer
Dim column As Integer
Dim pocket As Integer
Dim LastColumn As Integer


    If AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Drill Then
        LastColumn = TotalDRILL
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If

    For shelf = 1 To 3
        For column = 1 To LastColumn
            AutomationStatus(shelf, column).diameter = MyDiameter
        Next
    Next
    
    Call SaveAutomation

End Sub

Public Function FindPocket _
    (MyDiameter As Integer, _
        myCurrentTool As tooltype, _
            MyStatus As PocketStatus, _
                MyWorkPiece As Integer, _
                    MyProgramNumber As Double) As String

Dim shelf As Integer
Dim column As Integer
Dim digit As String
Dim diameter As Integer
Dim LastColumn As Integer
Dim BoolNeighbor As Boolean
    
    FindPocket = "0"
    
    If AppToolType = Drill Then
        LastColumn = TotalDRILL
        myCurrentTool = Drill
         
    ElseIf AppToolType = HSK Then
        LastColumn = TotalHSK
        myCurrentTool = HSK
        
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
        myCurrentTool = Round
    End If

    AppToolsWPiece = fMainForm.txtWorkPiece(1).Text
    MyWorkPiece = AppToolsWPiece
    
    For shelf = 1 To 3
        If AppShelvs(shelf).ShelfToolType = AppToolType Then
            For column = 1 To LastColumn
               
                If (column <= 9) Then
                    digit = "0"
                Else
                    digit = ""
                End If
                
               If _
                (AutomationStatus(shelf, column).diameter = MyDiameter) And _
                (AutomationStatus(shelf, column).CurrentTool = myCurrentTool) And _
                (AutomationStatus(shelf, column).Status = MyStatus) Then
                
                BoolNeighbor = ModAutomationStatus.GetPocketNeighborsStatus(CStr(shelf) & digit & CStr(column), occupied)
                
                     If BoolNeighbor = True Then
                     
                    ''(AutomationStatus(shelf, column).WorkPiece = MyWorkPiece)
                    
                         If AppToolType = Drill Then
                             FindPocket = CStr(shelf) & digit & CStr(column) & "." & CStr(MyDiameter) ''if pocket found
                             Exit Function
                             
                         ElseIf AppToolType = HSK Then
                             FindPocket = CStr(shelf) & digit & CStr(column)
                             Exit Function
                             
                         ElseIf AppToolType = Round Then
                             FindPocket = CStr(shelf) & digit & CStr(column) & "." & CStr(MyDiameter) ''if pocket found
                             Exit Function
                             
                         End If
                     Else
                         FindPocket = "0" ''if pocket not found
                    End If
               End If
            Next
        End If
    Next
'''(AutomationStatus(shelf, column).ProgramNumber = MyProgramNumber)
'''Public AutomationStatus(4, 13) As PocketProperties
End Function



Public Sub SetGroupWorkPiece _
    (ByVal WPNumber As Integer, _
        ByVal FirstPocket As Integer, _
            ByVal ToolAmount As Integer)
            
Dim shelf As Integer
Dim column As Integer
Dim jj As Integer

    shelf = Left(CStr(FirstPocket), 1)
    column = Right(FirstPocket, 2)
    
    For jj = 1 To ToolAmount
        AutomationStatus(shelf, column).WorkPiece = WPNumber
        
        If column <= 11 Then
            column = column + 1
        End If
        
        If column = 12 Then
            shelf = shelf + 1
            column = 1
        End If
        
    Next
    
    
End Sub

Public Sub SetGroupProgramNumber _
    (ByVal ProgramNumber As Integer, _
        ByVal FirstPocket As Integer, _
            ByVal ToolAmount As Integer)
            
                        
Dim shelf As Integer
Dim column As Integer
Dim jj As Integer

    shelf = Left(CStr(FirstPocket), 1)
    column = Right(FirstPocket, 2)
    
    For jj = 1 To ToolAmount
        AutomationStatus(shelf, column).ProgramNumber = ProgramNumber
        
        If column <= 11 Then
            column = column + 1
        End If
        
        If column = 12 Then
            shelf = shelf + 1
            column = 1
        End If
        
    Next
    
End Sub

Public Sub SetGroupDiameter _
    (ByVal diameter As Integer, _
        ByVal FirstPocket As Integer, _
            ByVal ToolAmount As Integer)
            
                        
Dim shelf As Integer
Dim column As Integer
Dim jj As Integer

    shelf = Left(CStr(FirstPocket), 1)
    column = Right(FirstPocket, 2)
    
    For jj = 1 To ToolAmount
        AutomationStatus(shelf, column).diameter = diameter
        
        If column <= 11 Then
            column = column + 1
        End If
        
        If column = 12 Then
            shelf = shelf + 1
            column = 1
        End If
        
    Next
    
End Sub
    




Public Sub DecodePocketFromString _
    (ByVal MyString As String, _
        ByRef MyShelf As Integer, _
            ByRef MyColumn As Integer, _
                ByRef MyPocket As Integer)
                
        If MyString = "" Then
            Exit Sub
        End If
                
    If AppToolType = Drill Then
        MyShelf = CInt(Left(MyString, 1))
        MyColumn = CInt(Right(Left(MyString, 3), 2))
        MyPocket = CInt(Right(MyString, 1))
    
    ElseIf AppToolType = HSK Then
        MyShelf = CInt(Left(MyString, 1))
        MyColumn = CInt(Right(Left(MyString, 3), 2))
        MyPocket = 1
        
    ElseIf AppToolType = Round Then
        MyShelf = CInt(Left(MyString, 1))
        MyColumn = CInt(Right(Left(MyString, 3), 2))
        MyPocket = CInt(Right(MyString, 1))
        
    End If
                

End Sub


Public Function FindEmptyPocket() As String
''1.the function find empty pocket in the automation.
''2.the f serch empty pocket in shelf witch is Enable.
''3.the f return a String
''      if empty pocket found     :the name of the Pocket.
''      if empty pocket not found :"0".

Dim shelf As Integer
Dim column As Integer
Dim digit As String
Dim diameter As Integer
Dim LastColumn As Integer
Dim BoolNeighbor As Boolean

    If AppToolType = Drill Then ''set numbers of columns
        LastColumn = TotalDRILL
    ElseIf AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If
    
    FindEmptyPocket = "0"
     
    For shelf = 1 To 3
        If (AppShelvs(shelf).ShelfToolType = AppToolType) Then
            For column = 1 To LastColumn
                AppDiameter = AllWP(AppWPIndex).ToolDiameter
                diameter = AppDiameter ''the diameter should be as the App diameter.
                
                If (column <= 9) Then
                  digit = "0"
                Else
                  digit = ""
                End If
                
                ''if empty pocket was found
                If ((AutomationStatus(shelf, column).Status = PocketStatus.empty_)) Then
                    BoolNeighbor = ModAutomationStatus.GetPocketNeighborsStatus(CStr(shelf) & digit & CStr(column), occupied)
                    If BoolNeighbor = True Then
                        ''build the name of the pocket
                        If AppToolType = Drill Then
                            FindEmptyPocket = CStr(shelf) & digit & CStr(column) & "." & CStr(diameter)
                        ElseIf AppToolType = HSK Then
                            FindEmptyPocket = CStr(shelf) & digit & CStr(column)
                        ElseIf AppToolType = Round Then
                            FindEmptyPocket = CStr(shelf) & digit & CStr(column) & "." & CStr(diameter)
                        End If
                        Exit Function
                    Else
                        FindEmptyPocket = "0" ''if pocket not found
                    End If
                End If
            Next column
        End If
    Next shelf
    
End Function



Public Function GetPocketStatus(ByVal shelf As Integer, ByVal column As Integer) As PocketStatus
    
Dim MyStatus As PocketStatus

'    If AppShelvs(shelf).ShelfEnable = False Then
'        GetPocketStatus = occupied
'        Exit Function
'    End If

    If AppToolType = HSK Then
        MyStatus = AutomationStatus(shelf, column).Status
        
    ElseIf AppToolType = Drill Then
        MyStatus = AutomationStatus(shelf, column).Status
        
    ElseIf AppToolType = Round Then
        MyStatus = AutomationStatus(shelf, column).Status
        
    End If
    
    GetPocketStatus = MyStatus

End Function

Public Sub SetAutomationByShelvs()
''1.this function set automation Values by the configuration given by the user.
''2.the function write the data from the memory into the TextFile.


Dim ii As Integer
Dim jj As Integer
Dim kk As Integer
Dim digit As String

Dim LastColumn As Integer

    If AppToolType = Drill Then ''set numbers of columns
        LastColumn = TotalDRILL
    ElseIf AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If

    For ii = 1 To 3
        For jj = 1 To LastColumn
        
            AutomationStatus(ii, jj).shelf = ii ''                              shelf
            AutomationStatus(ii, jj).column = jj ''                             column
            AutomationStatus(ii, jj).pocket = AppShelvs(ii).DefaultPocket ''    pocket
            AutomationStatus(ii, jj).diameter = AppShelvs(ii).DefaultDiameter ''diameter
            AutomationStatus(ii, jj).CurrentTool = AppShelvs(ii).ShelfToolType ''tooltype
            AutomationStatus(ii, jj).Status = AppShelvs(ii).ShelfStatus ''      status
            AutomationStatus(ii, jj).ProgramNumber = 0 ''                       program Number
            AutomationStatus(ii, jj).WorkPiece = 0 ''                           WorkPiece
            
            ''name
            If jj > 9 Then
                digit = ""
            Else
                digit = "0"
            End If
            
            AutomationStatus(ii, jj).name = CStr(ii) & digit & CStr(jj)

        Next
    Next
    
    Call ModCsvFile.SaveAutomation ''   save automation into file.
    
End Sub



Public Sub SetAutomationAllOver _
    (ByVal MyDiameter As Integer, _
        ByVal myCurrentTool As tooltype, _
            ByVal MyStatus As PocketStatus, _
                ByVal MyWP As Integer, _
                    ByVal MyProgNumber As Double)

Dim MyShelf As Integer
Dim MyColumn As Integer
Dim column As Integer
Dim digit As String

Dim LastColumn As Integer

    If AppToolType = Drill Then ''set numbers of columns
        LastColumn = TotalDRILL
    ElseIf AppToolType = HSK Then
        LastColumn = TotalHSK
    ElseIf AppToolType = Round Then
        LastColumn = TotalROUND
    End If


    For MyShelf = 1 To 3
        For MyColumn = 1 To LastColumn
        
            AutomationStatus(MyShelf, MyColumn).shelf = MyShelf
            AutomationStatus(MyShelf, MyColumn).column = MyColumn
            AutomationStatus(MyShelf, MyColumn).diameter = MyDiameter
            AutomationStatus(MyShelf, MyColumn).CurrentTool = myCurrentTool
            AutomationStatus(MyShelf, MyColumn).Status = MyStatus
            AutomationStatus(MyShelf, MyColumn).WorkPiece = MyWP
            AutomationStatus(MyShelf, MyColumn).ProgramNumber = MyProgNumber
            
            If column <= 9 Then
                digit = "0"
            ElseIf column >= 10 Then
                digit = ""
            End If
            AutomationStatus(MyShelf, MyColumn).name = CStr(MyShelf) & digit & CStr(MyColumn)
            
        Next
    Next
    
    Call SaveAutomation
    
End Sub

Public Function GetPocketNeighborsStatus(ByVal MyPocket As String, ByVal MyStatus As PocketStatus) As Boolean

''1.the function read the neighbors status.
''2.the function return TRUE  if both neighbors are the same as wanted.
''3.the function return FALSE if both neighbors are NOT as wanted.
''4.the function return FALSE if the Current pocket is 101/201/301.

Dim TempPocket As Integer
Dim TempStatus As PocketStatus
Dim shelf As Integer
Dim column As Integer

    If AppToolType <> HSK Then
        GetPocketNeighborsStatus = True
        Exit Function
    End If
        
        
    If AppDiameter <= 150 Then
        GetPocketNeighborsStatus = True
        Exit Function
    End If
    
    If MyPocket = "" Then
        Exit Function
    End If
    
    
    GetPocketNeighborsStatus = False
    
    TempPocket = CInt(MyPocket)
    shelf = Left(CInt(MyPocket), 1)
    column = Right(CInt(MyPocket), 2)
    
    If ((TempPocket = 101) Or (TempPocket = 201) Or (TempPocket = 301)) Then
         GetPocketNeighborsStatus = False
        Exit Function
    End If
    
    
    If ((TempPocket <> 101) And (TempPocket <> 201) And (TempPocket <> 301)) Then
        TempStatus = GetPocketStatus(shelf, column - 1)
        If TempStatus = MyStatus Then
            GetPocketNeighborsStatus = True
        Else
            GetPocketNeighborsStatus = False
            Exit Function
        End If
    End If
    
    If ((TempPocket <> 110) And (TempPocket <> 210) And (TempPocket <> 310)) Then
        TempStatus = GetPocketStatus(shelf, column + 1)
        If TempStatus = MyStatus Then
            GetPocketNeighborsStatus = True
        Else
            GetPocketNeighborsStatus = False
            Exit Function
        End If
    End If

End Function

Public Sub SetPocketNeighborsStatus(ByVal MyPocket As String, ByVal MyStatus As PocketStatus)

Dim TempPocket As Integer

    If AppToolType = HSK Then
        '''do nothing
    Else
        Exit Sub
    End If
        
    If AppDiameter > 150 Then
        ''do nothing
    Else
        Exit Sub
    End If
        
    TempPocket = CInt(MyPocket) - 1
    If ((TempPocket <> 101) And (TempPocket <> 201) And (TempPocket <> 301)) Then
        Call ModAutomationStatus.SetPocketStatus(MyStatus, CStr(TempPocket))
    End If
    
    TempPocket = CInt(MyPocket) + 1
    If ((TempPocket <> 110) And (TempPocket <> 210) And (TempPocket <> 310)) Then
        Call ModAutomationStatus.SetPocketStatus(MyStatus, CStr(TempPocket))
    End If

End Sub

Public Sub AutomationStatusReset()

    If AppToolType = HSK Then
        Call SetAutomationAllOver(0, HSK, empty_, 0, 0)
    ElseIf AppToolType = Drill Then
        Call SetAutomationAllOver(0, Drill, empty_, 0, 0)
    ElseIf AppToolType = Round Then
        Call SetAutomationAllOver(0, Round, empty_, 0, 0)
    End If
        

    
End Sub

Public Sub SetBigHSKStatus()
    
    Call SetPocketStatus(occupied, 101)
    Call SetPocketStatus(occupied, 103)
    Call SetPocketStatus(occupied, 105)
    Call SetPocketStatus(occupied, 107)
    Call SetPocketStatus(occupied, 109)
    
    
    Call SetPocketStatus(occupied, 201)
    Call SetPocketStatus(occupied, 203)
    Call SetPocketStatus(occupied, 205)
    Call SetPocketStatus(occupied, 207)
    Call SetPocketStatus(occupied, 209)
    
    Call SetPocketStatus(occupied, 301)
    Call SetPocketStatus(occupied, 303)
    Call SetPocketStatus(occupied, 305)
    Call SetPocketStatus(occupied, 307)
    Call SetPocketStatus(occupied, 309)
    
End Sub
''*****************************************************************
''    empty_ = 1 ''       no tool in pocket
''    machined = 2 ''     tool after GOOD proccess in pocket
''    Mask = 3 ''         pocket not good.can not be used.
''    Spare = 4 ''        **spare status.for future use.
''    occupied = 5 ''     pocket currently/temporarly can not be used.
''    reserved = 6 ''     tool is being proccessed in the machined
''    Broken = 7 ''       pocket is full.tool after BAD proccess
''    Unmachined = 8 ''   pocket is full.tool before proccess
''*****************************************************************
'''AutomationStatus(3, 12) As PocketProperties

'    name           As String       v
'    shelf          As Integer      v
'    column         As Integer      v
'    pocket         As Integer
'    diameter       As Integer      v
'    CurrentTool    As tooltype     v
'    Status         As PocketStatus v
'    WorkPiece      As Integer      v
'    ProgramNumber  As Double       v

''*****************************************************************

