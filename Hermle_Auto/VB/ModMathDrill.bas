Attribute VB_Name = "ModMathDrill"
Option Explicit

''DrillLoadConsts()
Dim H(7) As Double
Dim Rc(7) As Double

'' DrillSetGeometric
Dim Delta_x As Double
Dim Delta_y As Double
Dim C As Double
Dim Gamma_c As Double
Dim s(12) As Double
Dim Alfa As Double
Dim Alfa_1 As Double
Dim D_Alfa As Double
Dim D As Double
Dim Gamma_1 As Double
Dim Gamma(12) As Double
Dim Lambda(12) As Double


Public Sub DrillSetGeometric(ByVal shelf As Integer)
''
''1.the function Compute the commongeometric for the entire shelf.
''  C        - the string length between 101.1 to 112.1
''  Gamma_c  - is the angle between the Two Theoretic radius. R1 and R12.
''  S(jj)     - are the strings between Diameter(jj) and diameter(1).
'   Alfa
'   Alfa_1
'   D_Alfa
'   D
'   Gamma_1
'   Gamma(jj)
'   Lambda(jj)
'2.the function receive Shelf number.
'3.the function return nothing.

Dim jj As Integer
Dim Rconst As Double
Dim R1  As Double
Dim R12 As Double
Dim fdbk As Integer

On Error GoTo error

    Delta_x = Abs(DrillLocations(shelf, 1).diameter(1).X) + Abs(DrillLocations(shelf, 12).diameter(1).X)
    Delta_y = Abs(Abs(DrillLocations(shelf, 1).diameter(1).Y) - Abs(DrillLocations(shelf, 12).diameter(1).Y))
    C = Sqr(Delta_x ^ 2 + Delta_y ^ 2)
    Rconst = Rc(1)
    R1 = DrillLocations(shelf, 1).diameter(1).Dist
    R12 = DrillLocations(shelf, 12).diameter(1).Dist
    
    Gamma_c = ACos((-C ^ 2 + Rconst ^ 2 + Rconst ^ 2) / (2 * Rconst * Rconst)) / (2 * PI) * 360 ''calculate the real value
    '''Gamma_c = 78.54 theoretic
    
    For jj = 1 To 12
        s(jj) = 2 * Rconst * Sin(0.5 * ((Gamma_c) / (TotalDRILL - 1)) * (jj - 1) / 360 * (2 * PI))
    Next
    
    Alfa = (180 - Gamma_c) / 2
    Alfa_1 = ACos((R1 ^ 2 + s(12) ^ 2 - R12 ^ 2) / (2 * R1 * s(12))) / (2 * PI) * 360
    D_Alfa = Alfa - Alfa_1
    D = Sqr(Rconst ^ 2 + R1 ^ 2 - 2 * Rconst * R1 * Cos(D_Alfa / 360 * 2 * PI))
    Gamma_1 = ACos((Rconst ^ 2 + D ^ 2 - R1 ^ 2) / (2 * Rconst * D)) / (2 * PI) * 360

    For jj = 2 To 12
        Gamma(jj) = ACos((-s(jj) ^ 2 + Rconst ^ 2 + Rconst ^ 2) / (2 * Rconst * Rconst)) / (2 * PI) * 360
    Next
    
    For jj = 2 To 12
        If (((R1 < Rconst) And (R12 < Rconst)) Or ((R1 > Rconst) And (R12 < Rconst))) Then
            Lambda(jj) = Gamma_1 - Gamma(jj)
        ElseIf (((R1 < Rconst) And (R12 > Rconst)) Or ((R1 > Rconst) And (R12 > Rconst))) Then
            Lambda(jj) = Gamma_1 + Gamma(jj)
        End If
    Next
        
        
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillSetGeometric() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillSetGeometric")

End Sub

Public Sub DrillShelfCalculation(ByVal shelf As Integer)
''
''1.this function compute all locations for a giver shelf number.
''
Dim fdbk As Integer
On Error GoTo error


    Call DrillLoadConsts
    Call DrillCalcFirstDiameter(shelf)
    Call DrillCalcLastDiameter(shelf)
    Call DrillSetGeometric(shelf)
    
    Call DrillCalculateAllIndex_1(shelf, 1, Rc(1))

    Call DrillCalcDiameter(shelf, 2, 16.5)
    Call DrillCalculateAllIndex(shelf, 2, Rc(2))

    Call DrillCalcDiameter(shelf, 3, 35.5)
    Call DrillCalculateAllIndex(shelf, 3, Rc(3))

    Call DrillCalcDiameter(shelf, 4, 58)
    Call DrillCalculateAllIndex(shelf, 4, Rc(4))

    Call DrillCalc5thDiameter(shelf, Rc(5), Lambda(2))
    Call DrillCalculateAllIndex(shelf, 5, Rc(5))

    Call DrillCalcDiameter_2(shelf, 6, Rc(5) - Rc(6))
    Call DrillCalculateAllIndex(shelf, 6, Rc(6))

    Call DrillCalcDiameter_2(shelf, 7, Rc(5) - Rc(7))
    Call DrillCalculateAllIndex(shelf, 7, Rc(7))


    
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillShelfCalculation() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillShelfCalculation")

End Sub


Public Sub DrillCalcFirstDiameter(ByVal shelf As Integer)
''
''1.the function compute the NAME,DISTANCE,ANGLE of the first diameter in the First Pocket.

Dim xx As Double
Dim yy As Double
Dim fdbk As Integer
Dim temp As Double

On Error GoTo error

        ''name
        DrillLocations(shelf, 1).diameter(1).name = CStr(shelf) & "01" & ".1"
        
        ''distance
        xx = DrillLocations(shelf, 1).diameter(1).X
        yy = DrillLocations(shelf, 1).diameter(1).Y
       
        DrillLocations(shelf, 1).diameter(1).Dist = Sqr(xx ^ 2 + yy ^ 2)
        
        ''Alfa-calculate the Angle of the 1_st Pocket from the X-axis.
        If DrillLocations(shelf, 1).diameter(1).X >= 0 Then
            DrillLocations(shelf, 1).diameter(1).Alfa = Atn(Abs((DrillLocations(shelf, 1).diameter(1).Y)) / DrillLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
        ElseIf DrillLocations(shelf, 1).diameter(1).X < 0 Then
            DrillLocations(shelf, 1).diameter(1).Alfa = 180 - Atn(DrillLocations(shelf, 1).diameter(1).Y / DrillLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
        End If
        
        Call SaveArray("DrillLocations")
        
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillCalcFirstDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillCalcFirstDiameter")


End Sub

Public Sub DrillCalcLastDiameter(ByVal shelf As Integer)
''
''1.the function compute the NAME,DISTANCE,ANGLE of the First diameter(1) in the last Pocket (12).

Dim xx As Double
Dim yy As Double
Dim fdbk As Integer

On Error GoTo error
    ''name
    DrillLocations(shelf, 12).diameter(1).name = CStr(shelf) & "12" & ".1"
    
    ''distance
    xx = DrillLocations(shelf, 12).diameter(1).X
    yy = DrillLocations(shelf, 12).diameter(1).Y
    DrillLocations(shelf, 12).diameter(1).Dist = Sqr(xx ^ 2 + yy ^ 2)
    
    
    ''Alfa - calculate the Angle of the 12_th Pocket from the X-axis.
    If DrillLocations(shelf, 12).diameter(1).X >= 0 Then
        DrillLocations(shelf, 12).diameter(1).Alfa = Atn(Abs(DrillLocations(shelf, 12).diameter(1).Y) / DrillLocations(shelf, 12).diameter(1).X) * 360 / (2 * PI)
    ElseIf DrillLocations(shelf, 12).diameter(1).X < 0 Then
        DrillLocations(shelf, 12).diameter(1).Alfa = (-1) * (180 - Atn(Abs((DrillLocations(shelf, 12).diameter(1).Y)) / (DrillLocations(shelf, 12).diameter(1).X)) * 360 / (2 * PI))
    End If
    
    Call SaveArray("DrillLocations")
                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillCalcLastDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillCalcLastDiameter")

End Sub

Public Sub DrillCalculateAllIndex_1(ByVal shelf As Integer, ByVal diameter As Integer, ByVal Rc As Integer)
''
''1. this function compute Parameters for Diameter 1 for the Entire shelf.
''

Dim column As Integer
Dim Sign(12) As Integer
Dim Sign_2 As Integer
Dim R(12) As Double
Dim Betta(12) As Double
Dim Tetta(12) As Double
Dim fdbk As Integer
Dim dZ As Double
Dim dRx As Double
Dim dRy As Double
Dim R1 As Double
Dim digit As String


    R1 = DrillLocations(shelf, 1).diameter(1).Dist
    For column = 2 To 11
        R(column) = Sqr(D ^ 2 + Rc ^ 2 - 2 * D * Rc * Cos(Lambda(column) / 360 * 2 * PI))
        Betta(column) = ACos((R1 ^ 2 + R(column) ^ 2 - s(column) ^ 2) / (2 * R1 * R(column))) / (2 * PI) * 360
    Next
    
    ''name
    For column = 2 To 11
        If column < 10 Then
            digit = "0"
        Else
            digit = ""
        End If
        DrillLocations(shelf, column).diameter(1).name = CStr(shelf) & digit & CStr(column) & "." & CStr(1)
    Next
    
    ''assign value to the distance
    For column = 2 To 11
        DrillLocations(shelf, column).diameter(1).Dist = R(column)
    Next
    
    ''calculate alfa for all pockets.
    For column = 2 To 11
        DrillLocations(shelf, column).diameter(1).Alfa = Betta(column) + DrillLocations(shelf, 1).diameter(1).Alfa
    Next
    
    ''calculate the dZ parameter
    dZ = Abs(DrillLocations(shelf, 1).diameter(1).z - DrillLocations(shelf, 12).diameter(1).z) / (TotalDRILL - 1)
    
    ''calculate the Z coordinate
    If DrillLocations(shelf, 12).diameter(1).z > DrillLocations(shelf, 1).diameter(1).z Then
        For column = 2 To 11 ''calculate the z if the 12 pocket is higher then the 1.
            DrillLocations(shelf, column).diameter(1).z = DrillLocations(shelf, column - 1).diameter(1).z + dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z > DrillLocations(shelf, 12).diameter(1).z Then
        For column = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, column).diameter(1).z = DrillLocations(shelf, column - 1).diameter(1).z - dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z = DrillLocations(shelf, 12).diameter(1).z Then
        For column = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, column).diameter(1).z = DrillLocations(shelf, column - 1).diameter(1).z
        Next
    End If
    
    
    ''calculate the X coordinate
    For column = 2 To 11
        If DrillLocations(shelf, column).diameter(1).Alfa > -90 Then
            DrillLocations(shelf, column).diameter(1).X = DrillLocations(shelf, column).diameter(1).Dist * Cos(DrillLocations(shelf, column).diameter(1).Alfa * ((2 * PI) / (360)))
        ElseIf (DrillLocations(shelf, column).diameter(1).Alfa = -90) Then
            DrillLocations(shelf, column).diameter(1).X = DrillLocations(shelf, column).diameter(1).Dist
        ElseIf (DrillLocations(shelf, column).diameter(1).Alfa < -90) Then
            DrillLocations(shelf, column).diameter(1).X = (-1) * DrillLocations(shelf, column).diameter(1).Dist * Cos((180 - DrillLocations(shelf, column).diameter(1).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    
    
  ''calculate the y coordinate
    For column = 2 To 11
        If DrillLocations(shelf, column).diameter(1).Alfa < 90 Then
            DrillLocations(shelf, column).diameter(1).Y = -1 * DrillLocations(shelf, column).diameter(1).Dist * Sin(DrillLocations(shelf, column).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, column).diameter(1).Alfa = 90) Then
            DrillLocations(shelf, column).diameter(1).Y = -1 * DrillLocations(shelf, column).diameter(1).Dist
            
        ElseIf (DrillLocations(shelf, column).diameter(1).Alfa > 90) Then
            DrillLocations(shelf, column).diameter(1).Y = -1 * DrillLocations(shelf, column).diameter(1).Dist * Sin((DrillLocations(shelf, column).diameter(1).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    

    ''calculate the dRx
    dRx = Abs(DrillLocations(shelf, 1).diameter(1).Rx - DrillLocations(shelf, 12).diameter(1).Rx) / (TotalDRILL - 1)
    
    ''calculate the Rx
    If DrillLocations(shelf, 12).diameter(1).Rx > DrillLocations(shelf, 1).diameter(1).Rx Then
        For column = 2 To 11
            DrillLocations(shelf, column).diameter(1).Rx = DrillLocations(shelf, column - 1).diameter(1).Rx + dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx > DrillLocations(shelf, 12).diameter(1).Rx Then
        For column = 2 To 11
            DrillLocations(shelf, column).diameter(1).Rx = DrillLocations(shelf, column - 1).diameter(1).Rx - dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx = DrillLocations(shelf, 12).diameter(1).Rx Then
        For column = 2 To 11
            DrillLocations(shelf, column).diameter(1).Rx = DrillLocations(shelf, column - 1).diameter(1).Rx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(DrillLocations(shelf, 1).diameter(1).Ry - DrillLocations(shelf, 12).diameter(1).Ry) / (TotalDRILL - 1)
    
    ''calculate the Ry
    If DrillLocations(shelf, 12).diameter(1).Ry > DrillLocations(shelf, 1).diameter(1).Ry Then
        For column = 2 To 11
            DrillLocations(shelf, column).diameter(1).Ry = DrillLocations(shelf, column - 1).diameter(1).Ry + dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry > DrillLocations(shelf, 12).diameter(1).Ry Then
        For column = 2 To 11
            DrillLocations(shelf, column).diameter(1).Ry = DrillLocations(shelf, column - 1).diameter(1).Ry - dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry = DrillLocations(shelf, 12).diameter(1).Ry Then
        For column = 2 To 11
            DrillLocations(shelf, column).diameter(1).Ry = DrillLocations(shelf, column - 1).diameter(1).Ry
        Next
    End If
    
    ''calculate Tetta
    For column = 2 To 11
        Tetta(column) = ACos((-D ^ 2 + R(column) ^ 2 + Rc ^ 2) / (2 * R(column) * Rc)) / (2 * PI) * 360
    Next
    
    ''calculate the sign for the Tetta
    If ((R(1) > Rc) And (R(12) < Rc)) Then
        For column = 1 To 6
            Sign(column) = 1
        Next
        For column = 7 To 12
            Sign(column) = 1
        Next
        Sign_2 = 1
    
    ElseIf ((R(1) < Rc) And (R(12) < Rc)) Then
        For column = 1 To 6
            Sign(column) = 1
        Next
        For column = 7 To 12
            Sign(column) = -1
        Next
        Sign_2 = 1
        
    ElseIf ((R(1) < Rc) And (R(12) > Rc)) Then
        For column = 1 To 6
            Sign(column) = -1
        Next
        For column = 7 To 12
            Sign(column) = -1
        Next
        Sign_2 = -1
    
    ElseIf ((R(1) > Rc) And (R(12) > Rc)) Then
        For column = 1 To 6
            Sign(column) = -1
        Next
        For column = 7 To 12
            Sign(column) = 1
        Next
        Sign_2 = -1
    End If
    
    ''calculate the Rz
    For column = 2 To 11
        DrillLocations(shelf, column).diameter(1).Rz = 180 - DrillLocations(shelf, 1).diameter(1).Alfa - Betta(column) + Sign(column) * Tetta(column) + Sign_2 * D_Alfa
    Next

    Call SaveArray("DrillLocations")

                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillCalculateAllIndex_1() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillCalculateAllIndex_1")

End Sub


Public Sub DrillCalculateAllIndex(ByVal shelf As Integer, ByVal diameter As Integer, ByVal Rc As Double)

Dim jj As Integer
Dim R(12) As Double
Dim Betta(12) As Double
Dim Tetta(12) As Double
Dim Sign(12) As Integer
Dim Sign_2 As Integer
Dim fdbk As Integer
Dim dZ As Double
Dim dRx As Double
Dim dRy As Double
Dim Index As Integer
Dim digit As String
Dim R1 As Double


    R1 = DrillLocations(shelf, 1).diameter(diameter).Dist
    Index = diameter
    
        
    For jj = 1 To 12
        s(jj) = 2 * Rc * Sin(0.5 * ((Gamma_c) / (TotalDRILL - 1)) * (jj - 1) / 360 * (2 * PI))
    Next
    
    For jj = 2 To 12
        R(jj) = Sqr(D ^ 2 + Rc ^ 2 - 2 * D * Rc * Cos(Lambda(jj) / 360 * (2 * PI)))
        Betta(jj) = ACos((R1 ^ 2 + R(jj) ^ 2 - s(jj) ^ 2) / (2 * R1 * R(jj))) / (2 * PI) * 360
    Next
    
    ''name
    For jj = 1 To 12
        If jj < 10 Then
            digit = "0"
        Else
            digit = ""
        End If
        DrillLocations(shelf, jj).diameter(Index).name = CStr(shelf) & digit & CStr(jj) & "." & CStr(Index)
    Next
    
    ''Distance
    For jj = 2 To 12
        DrillLocations(shelf, jj).diameter(Index).Dist = R(jj)
    Next
    
    ''Alfa.
    For jj = 2 To 12
        DrillLocations(shelf, jj).diameter(Index).Alfa = Betta(jj) + DrillLocations(shelf, 1).diameter(Index).Alfa
    Next
    
    ''calculate the dZ parameter
    dZ = Abs(DrillLocations(shelf, 1).diameter(1).z - DrillLocations(shelf, 12).diameter(1).z) / (TotalDRILL - 1)
    
    ''calculate the Z coordinate
    If DrillLocations(shelf, 12).diameter(1).z > DrillLocations(shelf, 1).diameter(1).z Then
        For jj = 2 To 12                     ''calculate the z if the 12 pocket is higher then the 1.
            DrillLocations(shelf, jj).diameter(Index).z = DrillLocations(shelf, jj - 1).diameter(Index).z + dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z > DrillLocations(shelf, 12).diameter(1).z Then
        For jj = 2 To 12                     ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, jj).diameter(Index).z = DrillLocations(shelf, jj - 1).diameter(Index).z - dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z = DrillLocations(shelf, 12).diameter(1).z Then
        For jj = 2 To 12                    ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, jj).diameter(Index).z = DrillLocations(shelf, jj - 1).diameter(Index).z
        Next
    End If
    
    
    ''calculate the X coordinate
    For jj = 2 To 12
        If DrillLocations(shelf, jj).diameter(1).Alfa > -90 Then
            DrillLocations(shelf, jj).diameter(Index).X = _
                DrillLocations(shelf, jj).diameter(Index).Dist * Cos(DrillLocations(shelf, jj).diameter(Index).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, jj).diameter(1).Alfa = -90) Then
            DrillLocations(shelf, jj).diameter(Index).X = DrillLocations(shelf, jj).diameter(Index).Dist
            
        ElseIf (DrillLocations(shelf, jj).diameter(1).Alfa < -90) Then
            DrillLocations(shelf, jj).diameter(Index).X = _
                (-1) * DrillLocations(shelf, jj).diameter(Index).Dist * Cos((180 - DrillLocations(shelf, jj).diameter(Index).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    
  ''calculate the y coordinate
    For jj = 2 To 12
        If DrillLocations(shelf, jj).diameter(1).Alfa < 90 Then
            DrillLocations(shelf, jj).diameter(Index).Y = _
                -1 * DrillLocations(shelf, jj).diameter(Index).Dist * Sin(DrillLocations(shelf, jj).diameter(Index).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, jj).diameter(1).Alfa = 90) Then
            DrillLocations(shelf, jj).diameter(Index).Y = -1 * DrillLocations(shelf, jj).diameter(Index).Dist
            
        ElseIf (DrillLocations(shelf, jj).diameter(1).Alfa > 90) Then
            DrillLocations(shelf, jj).diameter(Index).Y = _
                -1 * DrillLocations(shelf, jj).diameter(Index).Dist * Sin((DrillLocations(shelf, jj).diameter(Index).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    

    ''calculate the dRx
    dRx = Abs(DrillLocations(shelf, 1).diameter(1).Rx - DrillLocations(shelf, 12).diameter(1).Rx) / (TotalDRILL - 1)
    
    ''calculate the Rx
    If DrillLocations(shelf, 12).diameter(1).Rx > DrillLocations(shelf, 1).diameter(1).Rx Then
       For jj = 2 To 12
            DrillLocations(shelf, jj).diameter(Index).Rx = DrillLocations(shelf, jj - 1).diameter(Index).Rx + dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx > DrillLocations(shelf, 12).diameter(1).Rx Then
       For jj = 2 To 12
            DrillLocations(shelf, jj).diameter(Index).Rx = DrillLocations(shelf, jj - 1).diameter(Index).Rx - dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx = DrillLocations(shelf, 12).diameter(1).Rx Then
        For jj = 2 To 12
            DrillLocations(shelf, jj).diameter(Index).Rx = DrillLocations(shelf, jj - 1).diameter(Index).Rx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(DrillLocations(shelf, 1).diameter(1).Ry - DrillLocations(shelf, 12).diameter(1).Ry) / (TotalDRILL - 1)
    
    ''calculate the Ry
    If DrillLocations(shelf, 12).diameter(1).Ry > DrillLocations(shelf, 1).diameter(1).Ry Then
        For jj = 2 To 12
            DrillLocations(shelf, jj).diameter(Index).Ry = DrillLocations(shelf, jj - 1).diameter(Index).Ry + dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry > DrillLocations(shelf, 12).diameter(1).Ry Then
        For jj = 2 To 12
            DrillLocations(shelf, jj).diameter(Index).Ry = DrillLocations(shelf, jj - 1).diameter(Index).Ry - dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry = DrillLocations(shelf, 12).diameter(1).Ry Then
        For jj = 2 To 12
            DrillLocations(shelf, jj).diameter(Index).Ry = DrillLocations(shelf, jj - 1).diameter(Index).Ry
        Next
    End If
    
    ''calculate Tetta
    For jj = 2 To 12
        Tetta(jj) = ACos((-D ^ 2 + R(jj) ^ 2 + Rc ^ 2) / (2 * R(jj) * Rc)) / (2 * PI) * 360
    Next
    
    ''calculate the sign for the Tetta
    If ((R(1) > Rc) And (R(12) < Rc)) Then
        For jj = 1 To 6
            Sign(jj) = 1
        Next
       For jj = 7 To 12
            Sign(jj) = 1
        Next
        Sign_2 = 1
    
    ElseIf ((R(1) < Rc) And (R(12) < Rc)) Then
       For jj = 1 To 6
            Sign(jj) = 1
        Next
        For jj = 7 To 12
            Sign(jj) = -1
        Next
        Sign_2 = 1
        
    ElseIf ((R(1) < Rc) And (R(12) > Rc)) Then
        For jj = 1 To 6
            Sign(jj) = -1
        Next
        For jj = 7 To 12
            Sign(jj) = -1
        Next
        Sign_2 = -1
    
    ElseIf ((R(1) > Rc) And (R(12) > Rc)) Then
        For jj = 1 To 6
            Sign(jj) = -1
        Next
        For jj = 7 To 12
            Sign(jj) = 1
        Next
        Sign_2 = -1
    End If
    
    ''calculate the Rz
    For jj = 2 To 12
        DrillLocations(shelf, jj).diameter(Index).Rz = 180 - DrillLocations(shelf, 1).diameter(Index).Alfa - Betta(jj) + Sign(jj) * Tetta(jj) + Sign_2 * D_Alfa
    Next

    Call SaveArray("DrillLocations")
                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillCalculateAllIndex() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillCalculateAllIndex")

End Sub


Public Sub DrillCalcDiameter(ByVal shelf As Integer, ByVal Index As Integer, ByVal H As Double)
''
''1.the function compute the parameters for a specific pocket,Index,in Column Number 1.
''
Dim Tetta As Double
Dim R1 As Double
Dim R12  As Double
Dim CurrRad As Double
Dim Omega As Double
Dim fdbk As Integer
Dim jj As Integer
Dim Sign As Integer

On Error GoTo error

    R1 = DrillLocations(shelf, 1).diameter(1).Dist
    R12 = DrillLocations(shelf, 12).diameter(1).Dist
    
    Tetta = ACos((D ^ 2 - R1 ^ 2 - Rc(1) ^ 2) / (-2 * R1 * Rc(1))) / (2 * PI) * 360
    CurrRad = Sqr(R1 ^ 2 + H ^ 2 - 2 * R1 * H * Cos(Tetta / 360 * (2 * PI)))
    Omega = ACos(((H ^ 2) - (R1 ^ 2) - (CurrRad ^ 2)) / (-2 * R1 * CurrRad)) / (2 * PI) * 360
    
    ''name
    DrillLocations(shelf, 1).diameter(Index).name = CStr(shelf) & "01." & Index
    
    ''distance
    DrillLocations(shelf, 1).diameter(Index).Dist = CurrRad

    ''sign
    If (R1 < Rc(1)) And (R12 < Rc(1)) Then
        Sign = -1
    ElseIf (R1 < Rc(1)) And (R12 > Rc(1)) Then
        Sign = 1
    ElseIf (R1 > Rc(1)) And (R12 < Rc(1)) Then
        Sign = 1
    ElseIf (R1 > Rc(1)) And (R12 > Rc(1)) Then
        Sign = -1
    End If
    
    ''alfa
    DrillLocations(shelf, 1).diameter(Index).Alfa = DrillLocations(shelf, 1).diameter(1).Alfa + Omega
    '' X
    If DrillLocations(shelf, 1).diameter(Index).Alfa > -90 Then
        DrillLocations(shelf, 1).diameter(Index).X = DrillLocations(shelf, 1).diameter(Index).Dist * Cos(DrillLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa = -90) Then
        DrillLocations(shelf, 1).diameter(Index).X = DrillLocations(shelf, 1).diameter(Index).Dist
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa < -90) Then
        DrillLocations(shelf, 1).diameter(Index).X = (-1) * DrillLocations(shelf, 1).diameter(Index).Dist * Cos((180 - DrillLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If
    
    ''calculate the y coordinate
    If DrillLocations(shelf, 1).diameter(Index).Alfa < 90 Then
        DrillLocations(shelf, 1).diameter(Index).Y = -1 * DrillLocations(shelf, 1).diameter(Index).Dist * Sin(DrillLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa = 90) Then
        DrillLocations(shelf, 1).diameter(Index).Y = -1 * DrillLocations(shelf, 1).diameter(Index).Dist
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa > 90) Then
        DrillLocations(shelf, 1).diameter(Index).Y = -1 * DrillLocations(shelf, 1).diameter(Index).Dist * Sin((DrillLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If

    ''Z
    DrillLocations(shelf, 1).diameter(Index).z = DrillLocations(shelf, 1).diameter(1).z
    
    ''Rx
    DrillLocations(shelf, 1).diameter(Index).Rx = DrillLocations(shelf, 1).diameter(1).Rx
    
    ''Ry
    DrillLocations(shelf, 1).diameter(Index).Ry = DrillLocations(shelf, 1).diameter(1).Ry
    
    ''Rz
    DrillLocations(shelf, 1).diameter(Index).Rz = DrillLocations(shelf, 1).diameter(1).Rz
    
    
''    name As String
''    Dist As Double
''    Alfa As Double
''    X As Double
''    Y As Double
''    z As Double
''    Rx As Double
''    Ry As Double
''    Rz As Double

                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillCalcDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillCalcDiameter")

End Sub

Public Sub DrillCalc5thDiameter(ByVal shelf As Integer, ByVal Rc As Double, ByVal Delta As Double)
''
Dim R5 As Double
Dim MyAlfa As Double


    ''radius
    R5 = Sqr(D ^ 2 + Rc ^ 2 - 2 * D * Rc * Cos((Delta) / 360 * 2 * PI))
    DrillLocations(shelf, 1).diameter(5).Dist = R5
    
    ''alfa
    MyAlfa = DrillLocations(shelf, 1).diameter(1).Alfa + 3.57
    DrillLocations(shelf, 1).diameter(5).Alfa = MyAlfa
    
    ''name
    DrillLocations(shelf, 1).diameter(5).name = CStr(shelf) & "01." & CStr(5)

    '' X
    If DrillLocations(shelf, 1).diameter(5).Alfa > -90 Then
        DrillLocations(shelf, 1).diameter(5).X = _
            DrillLocations(shelf, 1).diameter(5).Dist * Cos(DrillLocations(shelf, 1).diameter(5).Alfa * ((2 * PI) / (360)))
    ElseIf (DrillLocations(shelf, 1).diameter(5).Alfa = -90) Then
        DrillLocations(shelf, 1).diameter(5).X = DrillLocations(shelf, 1).diameter(5).Dist
    ElseIf (DrillLocations(shelf, 1).diameter(5).Alfa < -90) Then
        DrillLocations(shelf, 1).diameter(5).X = _
            (-1) * DrillLocations(shelf, 1).diameter(5).Dist * Cos((180 - DrillLocations(shelf, 1).diameter(5).Alfa) * ((2 * PI) / (360)))
    End If
    
    ''calculate the y coordinate
    If DrillLocations(shelf, 1).diameter(5).Alfa < 90 Then
        DrillLocations(shelf, 1).diameter(5).Y = _
            -1 * DrillLocations(shelf, 1).diameter(5).Dist * Sin(DrillLocations(shelf, 1).diameter(5).Alfa * ((2 * PI) / (360)))
    ElseIf (DrillLocations(shelf, 1).diameter(5).Alfa = 90) Then
        DrillLocations(shelf, 1).diameter(5).Y = -1 * DrillLocations(shelf, 1).diameter(5).Dist
    ElseIf (DrillLocations(shelf, 1).diameter(5).Alfa > 90) Then
        DrillLocations(shelf, 1).diameter(5).Y = _
            -1 * DrillLocations(shelf, 1).diameter(5).Dist * Sin((DrillLocations(shelf, 1).diameter(5).Alfa) * ((2 * PI) / (360)))
    End If


    ''Z
    DrillLocations(shelf, 1).diameter(5).z = DrillLocations(shelf, 1).diameter(1).z
    
    ''Rx
    DrillLocations(shelf, 1).diameter(5).Rx = DrillLocations(shelf, 1).diameter(1).Rx
    
    ''Ry
    DrillLocations(shelf, 1).diameter(5).Ry = DrillLocations(shelf, 1).diameter(1).Ry
    
    ''Rz
    DrillLocations(shelf, 1).diameter(5).Rz = DrillLocations(shelf, 1).diameter(1).Rz - 3.57
    
    Call SaveArray("DrillLocations")
    
End Sub
Public Sub DrillLoadConsts()

''
Dim fdbk As Integer

'''    H(1) = 0
'''    H(2) = -16.5
'''    H(3) = -19
'''    H(4) = -22.5
'''    H(5) = -22.5
'''    H(6) = -22.5
'''    H(7) = -22.5
    
    Rc(1) = 1173
    Rc(2) = 1156.5
    Rc(3) = 1137.5
    Rc(4) = 1115
    ''Rc(5) = 1186
    Rc(5) = 1183
    ''Rc(6) = 1156
    Rc(6) = 1157.5
    Rc(7) = 1126
   
                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillLoadConsts() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillLoadConsts")

End Sub




Public Sub DrillCalcDiameter_2(ByVal shelf As Integer, ByVal Index As Integer, ByVal H As Double)
''
''1.the function compute the parameters for a specific pocket,Index,in Column Number 1.
''
Dim Tetta As Double
Dim R1 As Double
Dim R12  As Double
Dim CurrRad As Double
Dim Omega As Double
Dim fdbk As Integer
Dim jj As Integer
Dim Sign As Integer

On Error GoTo error

    R1 = DrillLocations(shelf, 1).diameter(5).Dist
    R12 = DrillLocations(shelf, 12).diameter(5).Dist
    
    Tetta = ACos((D ^ 2 - R1 ^ 2 - Rc(5) ^ 2) / (-2 * R1 * Rc(5))) / (2 * PI) * 360
    CurrRad = Sqr(R1 ^ 2 + H ^ 2 - 2 * R1 * H * Cos(Tetta / 360 * 2 * PI))
    Omega = ACos(((H ^ 2) - (R1 ^ 2) - (CurrRad ^ 2)) / (-2 * R1 * CurrRad)) / (2 * PI) * 360
    
    ''name
    DrillLocations(shelf, 1).diameter(Index).name = CStr(shelf) & "01." & Index
    
    ''distance
    DrillLocations(shelf, 1).diameter(Index).Dist = CurrRad

    ''sign
    If (R1 < Rc(1)) And (R12 < Rc(1)) Then
        Sign = -1
    ElseIf (R1 < Rc(1)) And (R12 > Rc(1)) Then
        Sign = 1
    ElseIf (R1 > Rc(1)) And (R12 < Rc(1)) Then
        Sign = 1
    ElseIf (R1 > Rc(1)) And (R12 > Rc(1)) Then
        Sign = -1
    End If
    
    ''alfa
    DrillLocations(shelf, 1).diameter(Index).Alfa = DrillLocations(shelf, 1).diameter(5).Alfa + Omega
    
    '' X
    If DrillLocations(shelf, 1).diameter(Index).Alfa > -90 Then
        DrillLocations(shelf, 1).diameter(Index).X = DrillLocations(shelf, 1).diameter(Index).Dist * Cos(DrillLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa = -90) Then
        DrillLocations(shelf, 1).diameter(Index).X = DrillLocations(shelf, 1).diameter(Index).Dist
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa < -90) Then
        DrillLocations(shelf, 1).diameter(Index).X = (-1) * DrillLocations(shelf, 1).diameter(Index).Dist * Cos((180 - DrillLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If
    
    ''calculate the y coordinate
    If DrillLocations(shelf, 1).diameter(Index).Alfa < 90 Then
        DrillLocations(shelf, 1).diameter(Index).Y = -1 * DrillLocations(shelf, 1).diameter(Index).Dist * Sin(DrillLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa = 90) Then
        DrillLocations(shelf, 1).diameter(Index).Y = -1 * DrillLocations(shelf, 1).diameter(Index).Dist
    ElseIf (DrillLocations(shelf, 1).diameter(Index).Alfa > 90) Then
        DrillLocations(shelf, 1).diameter(Index).Y = -1 * DrillLocations(shelf, 1).diameter(Index).Dist * Sin((DrillLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If

    ''Z
    DrillLocations(shelf, 1).diameter(Index).z = DrillLocations(shelf, 1).diameter(1).z
    
    ''Rx
    DrillLocations(shelf, 1).diameter(Index).Rx = DrillLocations(shelf, 1).diameter(1).Rx
    
    ''Ry
    DrillLocations(shelf, 1).diameter(Index).Ry = DrillLocations(shelf, 1).diameter(1).Ry
    
    ''Rz
    DrillLocations(shelf, 1).diameter(Index).Rz = DrillLocations(shelf, 1).diameter(5).Rz
    
    
''    name As String
''    Dist As Double
''    Alfa As Double
''    X As Double
''    Y As Double
''    z As Double
''    Rx As Double
''    Ry As Double
''    Rz As Double

                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in DrillCalcDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "DrillCalcDiameter")

End Sub

