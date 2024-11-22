Attribute VB_Name = "modAllWorkPiece"
Option Explicit

Public Sub AllWPReset()
''
''1.the function beeing called from:CmdResetAllWP_Click()
Dim i As Integer

    For i = 1 To 50
    
        AllWP(i).LineNumber = 0
        AllWP(i).WPNumber = 0
        AllWP(i).NCProgram = 0
        AllWP(i).ToolAmount = 0
        AllWP(i).ToolAmountLeft = 0
        AllWP(i).ToolDiameter = 0
        AllWP(i).WPToolType = "0"
        
    Next
    
    Call SaveAllWP
    
End Sub

Public Sub WorkPieceReset(ByVal MyLineNumber As Integer)
''
''1.the function beeing called from: BtnDelOrder_Click()
''
    ModGUI.ListHMIUpdate ("Delete WorkPiece Index    : " & MyLineNumber)
    ModGUI.ListHMIUpdate ("       WorkPiece number   : " & CStr(AllWP(MyLineNumber).WPNumber))
    ModGUI.ListHMIUpdate ("       WorkPiece Amount   : " & CStr(AllWP(MyLineNumber).ToolAmount))
    ModGUI.ListHMIUpdate ("       WorkPiece Diameter : " & CStr(AllWP(MyLineNumber).ToolDiameter))
    ModGUI.ListHMIUpdate ("       WorkPiece NCProgram : " & CStr(AllWP(MyLineNumber).NCProgram))

    AllWP(MyLineNumber).WPNumber = 0
    AllWP(MyLineNumber).NCProgram = 0
    AllWP(MyLineNumber).ToolAmount = 0
    AllWP(MyLineNumber).ToolAmountLeft = 0
    AllWP(MyLineNumber).ToolDiameter = 0
    AllWP(MyLineNumber).WPToolType = 0
    AllWP(MyLineNumber).WPStatus = empty_
    
    Call SaveAllWP
    
End Sub

Public Sub WorkPieceEdit(ByVal MyLineNumber As Integer, ByRef TempWP As WorkPiece)

''
''1. the function beeing called from: BtnAddWrokPiece_Click()
''

    AllWP(MyLineNumber).WPNumber = TempWP.WPNumber
    AllWP(MyLineNumber).LineNumber = TempWP.LineNumber
    AllWP(MyLineNumber).NCProgram = TempWP.NCProgram
    AllWP(MyLineNumber).ToolAmount = TempWP.ToolAmount
    AllWP(MyLineNumber).ToolAmountLeft = TempWP.ToolAmountLeft
    AllWP(MyLineNumber).ToolDiameter = TempWP.ToolDiameter
    AllWP(MyLineNumber).WPToolType = TempWP.WPToolType
    
    Call SaveAllWP


End Sub

Public Function WPExist() As Boolean
''1. the function check if there is minimum 1 WP in the table.
''2. the function take no arguments.
''3. the function return TRUE or FALSE.

    If AllWP(1).WPNumber = 0 Then
        WPExist = False
    ElseIf AllWP(1).WPNumber <> 0 Then
        WPExist = True
    End If
    
End Function



Public Function IsLegalDiameter(ByVal MyDiameter As String) As Boolean
''1. the function check if the parameter "MyDiameter" is a Legal number for application diameter.
''2. the function get one parameter :a number to be check.
''3. the function return one parameter :true or false.


Dim ii As Integer
    
    If MyDiameter = "" Then
        Exit Function
    End If
    
    IsLegalDiameter = False

    If AppToolType = Drill Then
        For ii = 1 To 7
            If CInt(MyDiameter) = ii Then
                 IsLegalDiameter = True
            End If
        Next
        
    ElseIf AppToolType = HSK Then
        If ((CInt(MyDiameter) = 100) Or _
            (CInt(MyDiameter) = 200) Or _
            (CInt(MyDiameter) = 300)) Then
            IsLegalDiameter = True
        End If
        
    ElseIf AppToolType = Round Then
        For ii = 1 To 8
            If CInt(MyDiameter) = ii Then
                 IsLegalDiameter = True
            End If
        Next
        
    End If
    
End Function


Public Sub ReOrderAllWPiece()
''1.this function reorder the AllWPiece array.
''2.main use after delete one order.

Dim line As Integer
Dim ret As Integer

On Error GoTo error
    
    For line = 1 To 40
    
        ''check if the current line is empty
        If AllWP(line).WPNumber = 0 Then
        
        ''push the next line one step up
            ''AllWP(line).LineNumber = AllWP(line).LineNumber
            AllWP(line).NCProgram = AllWP(line + 1).NCProgram
            AllWP(line).ToolAmount = AllWP(line + 1).ToolAmount
            AllWP(line).ToolAmountLeft = AllWP(line + 1).ToolAmountLeft
            AllWP(line).ToolDiameter = AllWP(line + 1).ToolDiameter
            AllWP(line).WPNumber = AllWP(line + 1).WPNumber
            AllWP(line).WPToolType = AllWP(line + 1).WPToolType
            AllWP(line).WPStatus = AllWP(line + 1).WPStatus
            
        ''erase the next line
            AllWP(line + 1).LineNumber = AllWP(line + 1).LineNumber
            AllWP(line + 1).NCProgram = 0
            AllWP(line + 1).ToolAmount = 0
            AllWP(line + 1).ToolAmountLeft = 0
            AllWP(line + 1).ToolDiameter = 0
            AllWP(line + 1).WPNumber = 0
            AllWP(line + 1).WPToolType = 0
            AllWP(line + 1).WPStatus = PocketStatus.empty_
        
        End If
    Next
   

Exit Sub
error:
    ret = MsgBox("Error while reorder AllWp()" & vbCrLf _
   & "the error is : " & Err.Description _
   , vbExclamation, _
    "Reorder AllWorkPieces")
    
End Sub

Public Function CheckToolAmount() As Boolean
''1.the function check the total amount of  tools
''  in all off the worlPieces.
''2.if the amount is legal the function return TRUE
''3.if the amount is Illigal the f return FALSE.

Dim jj As Integer
Dim amount As Integer

    amount = 0
    
    For jj = 1 To 50
        amount = amount + AllWP(jj).ToolAmount
    Next
    
    If AppToolType = Drill Then
        If amount > (3 * TotalDRILL) Then
            CheckToolAmount = False
            Exit Function
        Else
            CheckToolAmount = True
            Exit Function
        End If
            
            
            
    ElseIf AppToolType = HSK Then
        If amount > (3 * TotalHSK) Then
            CheckToolAmount = False
            Exit Function
        Else
            CheckToolAmount = True
            Exit Function
        End If
    
    
    ElseIf AppToolType = Round Then
        If amount > (3 * TotalROUND) Then
            CheckToolAmount = False
            Exit Function
        Else
            CheckToolAmount = True
            Exit Function
        End If
    
    End If
    
    
End Function

Public Function AllDiametersTheSame() As Boolean
''1.the function check if all the diameters in the AllWp() array are the same.
''2.if all the same return TRUE
''3.if all not the same return FALSE

Dim kk As Integer
Dim BoolTemp As Boolean
    
    For kk = 1 To 50
        If AllWP(kk).ToolDiameter <> 0 Then
        
        
            If AllWP(kk).ToolDiameter = AllWP(1).ToolDiameter Then
                BoolTemp = True
                AllDiametersTheSame = BoolTemp
                
            ElseIf AllWP(kk).ToolDiameter <> AllWP(1).ToolDiameter Then
                BoolTemp = False
                AllDiametersTheSame = BoolTemp
                Exit Function
            End If
            
        End If
        
    Next

End Function















