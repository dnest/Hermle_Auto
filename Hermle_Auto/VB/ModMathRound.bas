Attribute VB_Name = "ModMathRound"
Option Explicit

''RoundLoadConsts()
Dim H(8) As Double
Dim Rc(8) As Double

'' RoundSetGeometric
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


Public Sub RoundSetGeometric(ByVal shelf As Integer)
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

    Delta_x = Abs(RoundLocations(shelf, 1).diameter(1).X) + Abs(RoundLocations(shelf, 12).diameter(1).X)
    Delta_y = Abs(Abs(RoundLocations(shelf, 1).diameter(1).Y) - Abs(RoundLocations(shelf, 12).diameter(1).Y))
    C = Sqr(Delta_x ^ 2 + Delta_y ^ 2)
    Rconst = Rc(1)
    R1 = RoundLocations(shelf, 1).diameter(1).Dist
    R12 = RoundLocations(shelf, 12).diameter(1).Dist
    
    Gamma_c = ACos((-C ^ 2 + Rconst ^ 2 + Rconst ^ 2) / (2 * Rconst * Rconst)) / (2 * PI) * 360 ''calculate the real value
    '''Gamma_c = 78.54 theoretic
    
    For jj = 1 To 12
        s(jj) = 2 * Rconst * Sin(0.5 * ((Gamma_c) / (TotalROUND - 1)) * (jj - 1) / 360 * (2 * PI))
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
    fdbk = MsgBox(vbCrLf & "error in RoundSetGeometric() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundSetGeometric")

End Sub

Public Sub RoundShelfCalculation(ByVal shelf As Integer)
''
''1.this function compute all locations for a giver shelf number.
''
Dim fdbk As Integer
On Error GoTo error


    Call RoundLoadConsts
    Call RoundCalcFirstDiameter(shelf)
    Call RoundCalcLastDiameter(shelf)
    Call RoundSetGeometric(shelf)
    
    Call RoundCalculateAllIndex_1(shelf, 1, Rc(1))

    Call RoundCalcDiameter(shelf, 2, H(2))
    Call RoundCalculateAllIndex(shelf, 2, Rc(2))

    Call RoundCalcDiameter(shelf, 3, H(3))
    Call RoundCalculateAllIndex(shelf, 3, Rc(3))

    Call RoundCalcDiameter(shelf, 4, H(4))
    Call RoundCalculateAllIndex(shelf, 4, Rc(4))

    Call RoundCalc5thDiameter(shelf, Rc(5), Lambda(2))
    Call RoundCalculateAllIndex(shelf, 5, Rc(5))

    Call RoundCalcDiameter_2(shelf, 6, H(6))
    Call RoundCalculateAllIndex(shelf, 6, Rc(6))

    Call RoundCalcDiameter_2(shelf, 7, H(7))
    Call RoundCalculateAllIndex(shelf, 7, Rc(7))

    Call RoundCalcDiameter_2(shelf, 8, H(8))
    Call RoundCalculateAllIndex(shelf, 8, Rc(8))

    
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in RoundShelfCalculation() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundShelfCalculation")

End Sub


Public Sub RoundCalcFirstDiameter(ByVal shelf As Integer)
''
''1.the function compute the NAME,DISTANCE,ANGLE of the first diameter in the First Pocket.

Dim xx As Double
Dim yy As Double
Dim fdbk As Integer
Dim temp As Double

On Error GoTo error

        ''name
        RoundLocations(shelf, 1).diameter(1).name = CStr(shelf) & "01" & ".1"
        
        ''distance
        xx = RoundLocations(shelf, 1).diameter(1).X
        yy = RoundLocations(shelf, 1).diameter(1).Y
       
        RoundLocations(shelf, 1).diameter(1).Dist = Sqr(xx ^ 2 + yy ^ 2)
        
        ''Alfa-calculate the Angle of the 1_st Pocket from the X-axis.
        If RoundLocations(shelf, 1).diameter(1).X >= 0 Then
            RoundLocations(shelf, 1).diameter(1).Alfa = Atn(Abs((RoundLocations(shelf, 1).diameter(1).Y)) / RoundLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
        ElseIf RoundLocations(shelf, 1).diameter(1).X < 0 Then
            RoundLocations(shelf, 1).diameter(1).Alfa = 180 - Atn(RoundLocations(shelf, 1).diameter(1).Y / RoundLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
        End If
        
        Call SaveArray("RoundLocations")
        
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in RoundCalcFirstDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundCalcFirstDiameter")


End Sub

Public Sub RoundCalcLastDiameter(ByVal shelf As Integer)
''
''1.the function compute the NAME,DISTANCE,ANGLE of the First diameter(1) in the last Pocket (12).

Dim xx As Double
Dim yy As Double
Dim fdbk As Integer

On Error GoTo error
    ''name
    RoundLocations(shelf, 12).diameter(1).name = CStr(shelf) & "12" & ".1"
    
    ''distance
    xx = RoundLocations(shelf, 12).diameter(1).X
    yy = RoundLocations(shelf, 12).diameter(1).Y
    RoundLocations(shelf, 12).diameter(1).Dist = Sqr(xx ^ 2 + yy ^ 2)
    
    
    ''Alfa - calculate the Angle of the 12_th Pocket from the X-axis.
    If RoundLocations(shelf, 12).diameter(1).X >= 0 Then
        RoundLocations(shelf, 12).diameter(1).Alfa = Atn(Abs(RoundLocations(shelf, 12).diameter(1).Y) / RoundLocations(shelf, 12).diameter(1).X) * 360 / (2 * PI)
    ElseIf RoundLocations(shelf, 12).diameter(1).X < 0 Then
        RoundLocations(shelf, 12).diameter(1).Alfa = (-1) * (180 - Atn(Abs((RoundLocations(shelf, 12).diameter(1).Y)) / (RoundLocations(shelf, 12).diameter(1).X)) * 360 / (2 * PI))
    End If
    
    Call SaveArray("RoundLocations")
                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in RoundCalcLastDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundCalcLastDiameter")

End Sub

Public Sub RoundCalculateAllIndex_1(ByVal shelf As Integer, ByVal diameter As Integer, ByVal Rc As Integer)
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


    R1 = RoundLocations(shelf, 1).diameter(1).Dist
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
        RoundLocations(shelf, column).diameter(1).name = CStr(shelf) & digit & CStr(column) & "." & CStr(1)
    Next
    
    ''assign value to the distance
    For column = 2 To 11
        RoundLocations(shelf, column).diameter(1).Dist = R(column)
    Next
    
    ''calculate alfa for all pockets.
    For column = 2 To 11
        RoundLocations(shelf, column).diameter(1).Alfa = Betta(column) + RoundLocations(shelf, 1).diameter(1).Alfa
    Next
    
    ''calculate the dZ parameter
    dZ = Abs(RoundLocations(shelf, 1).diameter(1).z - RoundLocations(shelf, 12).diameter(1).z) / (TotalROUND - 1)
    
    ''calculate the Z coordinate
    If RoundLocations(shelf, 12).diameter(1).z > RoundLocations(shelf, 1).diameter(1).z Then
        For column = 2 To 11 ''calculate the z if the 12 pocket is higher then the 1.
            RoundLocations(shelf, column).diameter(1).z = RoundLocations(shelf, column - 1).diameter(1).z + dZ
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).z > RoundLocations(shelf, 12).diameter(1).z Then
        For column = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            RoundLocations(shelf, column).diameter(1).z = RoundLocations(shelf, column - 1).diameter(1).z - dZ
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).z = RoundLocations(shelf, 12).diameter(1).z Then
        For column = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            RoundLocations(shelf, column).diameter(1).z = RoundLocations(shelf, column - 1).diameter(1).z
        Next
    End If
    
    
    ''calculate the X coordinate
    For column = 2 To 11
        If RoundLocations(shelf, column).diameter(1).Alfa > -90 Then
            RoundLocations(shelf, column).diameter(1).X = RoundLocations(shelf, column).diameter(1).Dist * Cos(RoundLocations(shelf, column).diameter(1).Alfa * ((2 * PI) / (360)))
        ElseIf (RoundLocations(shelf, column).diameter(1).Alfa = -90) Then
            RoundLocations(shelf, column).diameter(1).X = RoundLocations(shelf, column).diameter(1).Dist
        ElseIf (RoundLocations(shelf, column).diameter(1).Alfa < -90) Then
            RoundLocations(shelf, column).diameter(1).X = (-1) * RoundLocations(shelf, column).diameter(1).Dist * Cos((180 - RoundLocations(shelf, column).diameter(1).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    
    
  ''calculate the y coordinate
    For column = 2 To 11
        If RoundLocations(shelf, column).diameter(1).Alfa < 90 Then
            RoundLocations(shelf, column).diameter(1).Y = -1 * RoundLocations(shelf, column).diameter(1).Dist * Sin(RoundLocations(shelf, column).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, column).diameter(1).Alfa = 90) Then
            RoundLocations(shelf, column).diameter(1).Y = -1 * RoundLocations(shelf, column).diameter(1).Dist
            
        ElseIf (RoundLocations(shelf, column).diameter(1).Alfa > 90) Then
            RoundLocations(shelf, column).diameter(1).Y = -1 * RoundLocations(shelf, column).diameter(1).Dist * Sin((RoundLocations(shelf, column).diameter(1).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    

    ''calculate the dRx
    dRx = Abs(RoundLocations(shelf, 1).diameter(1).Rx - RoundLocations(shelf, 12).diameter(1).Rx) / (TotalROUND - 1)
    
    ''calculate the Rx
    If RoundLocations(shelf, 12).diameter(1).Rx > RoundLocations(shelf, 1).diameter(1).Rx Then
        For column = 2 To 11
            RoundLocations(shelf, column).diameter(1).Rx = RoundLocations(shelf, column - 1).diameter(1).Rx + dRx
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Rx > RoundLocations(shelf, 12).diameter(1).Rx Then
        For column = 2 To 11
            RoundLocations(shelf, column).diameter(1).Rx = RoundLocations(shelf, column - 1).diameter(1).Rx - dRx
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Rx = RoundLocations(shelf, 12).diameter(1).Rx Then
        For column = 2 To 11
            RoundLocations(shelf, column).diameter(1).Rx = RoundLocations(shelf, column - 1).diameter(1).Rx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(RoundLocations(shelf, 1).diameter(1).Ry - RoundLocations(shelf, 12).diameter(1).Ry) / (TotalROUND - 1)
    
    ''calculate the Ry
    If RoundLocations(shelf, 12).diameter(1).Ry > RoundLocations(shelf, 1).diameter(1).Ry Then
        For column = 2 To 11
            RoundLocations(shelf, column).diameter(1).Ry = RoundLocations(shelf, column - 1).diameter(1).Ry + dRy
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Ry > RoundLocations(shelf, 12).diameter(1).Ry Then
        For column = 2 To 11
            RoundLocations(shelf, column).diameter(1).Ry = RoundLocations(shelf, column - 1).diameter(1).Ry - dRy
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Ry = RoundLocations(shelf, 12).diameter(1).Ry Then
        For column = 2 To 11
            RoundLocations(shelf, column).diameter(1).Ry = RoundLocations(shelf, column - 1).diameter(1).Ry
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
        RoundLocations(shelf, column).diameter(1).Rz = 180 - RoundLocations(shelf, 1).diameter(1).Alfa - Betta(column) + Sign(column) * Tetta(column) + Sign_2 * D_Alfa
    Next

    Call SaveArray("RoundLocations")

                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in RoundCalculateAllIndex_1() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundCalculateAllIndex_1")

End Sub


Public Sub RoundCalculateAllIndex(ByVal shelf As Integer, ByVal diameter As Integer, ByVal Rc As Double)

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


    R1 = RoundLocations(shelf, 1).diameter(diameter).Dist
    Index = diameter
    
        
    For jj = 1 To 12
        s(jj) = 2 * Rc * Sin(0.5 * ((Gamma_c) / (TotalROUND - 1)) * (jj - 1) / 360 * (2 * PI))
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
        RoundLocations(shelf, jj).diameter(Index).name = CStr(shelf) & digit & CStr(jj) & "." & CStr(Index)
    Next
    
    ''Distance
    For jj = 2 To 12
        RoundLocations(shelf, jj).diameter(Index).Dist = R(jj)
    Next
    
    ''Alfa.
    For jj = 2 To 12
        RoundLocations(shelf, jj).diameter(Index).Alfa = Betta(jj) + RoundLocations(shelf, 1).diameter(Index).Alfa
    Next
    
    ''calculate the dZ parameter
    dZ = Abs(RoundLocations(shelf, 1).diameter(1).z - RoundLocations(shelf, 12).diameter(1).z) / (TotalROUND - 1)
    
    ''calculate the Z coordinate
    If RoundLocations(shelf, 12).diameter(1).z > RoundLocations(shelf, 1).diameter(1).z Then
        For jj = 2 To 12                     ''calculate the z if the 12 pocket is higher then the 1.
            RoundLocations(shelf, jj).diameter(Index).z = RoundLocations(shelf, jj - 1).diameter(Index).z + dZ
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).z > RoundLocations(shelf, 12).diameter(1).z Then
        For jj = 2 To 12                     ''calculate the z if the 1 pocket is higher then the 12.
            RoundLocations(shelf, jj).diameter(Index).z = RoundLocations(shelf, jj - 1).diameter(Index).z - dZ
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).z = RoundLocations(shelf, 12).diameter(1).z Then
        For jj = 2 To 12                    ''calculate the z if the 1 pocket is higher then the 12.
            RoundLocations(shelf, jj).diameter(Index).z = RoundLocations(shelf, jj - 1).diameter(Index).z
        Next
    End If
    
    
    ''calculate the X coordinate
    For jj = 2 To 12
        If RoundLocations(shelf, jj).diameter(1).Alfa > -90 Then
            RoundLocations(shelf, jj).diameter(Index).X = _
                RoundLocations(shelf, jj).diameter(Index).Dist * Cos(RoundLocations(shelf, jj).diameter(Index).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, jj).diameter(1).Alfa = -90) Then
            RoundLocations(shelf, jj).diameter(Index).X = RoundLocations(shelf, jj).diameter(Index).Dist
            
        ElseIf (RoundLocations(shelf, jj).diameter(1).Alfa < -90) Then
            RoundLocations(shelf, jj).diameter(Index).X = _
                (-1) * RoundLocations(shelf, jj).diameter(Index).Dist * Cos((180 - RoundLocations(shelf, jj).diameter(Index).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    
  ''calculate the y coordinate
    For jj = 2 To 12
        If RoundLocations(shelf, jj).diameter(1).Alfa < 90 Then
            RoundLocations(shelf, jj).diameter(Index).Y = _
                -1 * RoundLocations(shelf, jj).diameter(Index).Dist * Sin(RoundLocations(shelf, jj).diameter(Index).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, jj).diameter(1).Alfa = 90) Then
            RoundLocations(shelf, jj).diameter(Index).Y = -1 * RoundLocations(shelf, jj).diameter(Index).Dist
            
        ElseIf (RoundLocations(shelf, jj).diameter(1).Alfa > 90) Then
            RoundLocations(shelf, jj).diameter(Index).Y = _
                -1 * RoundLocations(shelf, jj).diameter(Index).Dist * Sin((RoundLocations(shelf, jj).diameter(Index).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    

    ''calculate the dRx
    dRx = Abs(RoundLocations(shelf, 1).diameter(1).Rx - RoundLocations(shelf, 12).diameter(1).Rx) / (TotalROUND - 1)
    
    ''calculate the Rx
    If RoundLocations(shelf, 12).diameter(1).Rx > RoundLocations(shelf, 1).diameter(1).Rx Then
       For jj = 2 To 12
            RoundLocations(shelf, jj).diameter(Index).Rx = RoundLocations(shelf, jj - 1).diameter(Index).Rx + dRx
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Rx > RoundLocations(shelf, 12).diameter(1).Rx Then
       For jj = 2 To 12
            RoundLocations(shelf, jj).diameter(Index).Rx = RoundLocations(shelf, jj - 1).diameter(Index).Rx - dRx
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Rx = RoundLocations(shelf, 12).diameter(1).Rx Then
        For jj = 2 To 12
            RoundLocations(shelf, jj).diameter(Index).Rx = RoundLocations(shelf, jj - 1).diameter(Index).Rx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(RoundLocations(shelf, 1).diameter(1).Ry - RoundLocations(shelf, 12).diameter(1).Ry) / (TotalROUND - 1)
    
    ''calculate the Ry
    If RoundLocations(shelf, 12).diameter(1).Ry > RoundLocations(shelf, 1).diameter(1).Ry Then
        For jj = 2 To 12
            RoundLocations(shelf, jj).diameter(Index).Ry = RoundLocations(shelf, jj - 1).diameter(Index).Ry + dRy
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Ry > RoundLocations(shelf, 12).diameter(1).Ry Then
        For jj = 2 To 12
            RoundLocations(shelf, jj).diameter(Index).Ry = RoundLocations(shelf, jj - 1).diameter(Index).Ry - dRy
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Ry = RoundLocations(shelf, 12).diameter(1).Ry Then
        For jj = 2 To 12
            RoundLocations(shelf, jj).diameter(Index).Ry = RoundLocations(shelf, jj - 1).diameter(Index).Ry
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
        RoundLocations(shelf, jj).diameter(Index).Rz = 180 - RoundLocations(shelf, 1).diameter(Index).Alfa - Betta(jj) + Sign(jj) * Tetta(jj) + Sign_2 * D_Alfa
    Next

    Call SaveArray("RoundLocations")
                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in RoundCalculateAllIndex() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundCalculateAllIndex")

End Sub


Public Sub RoundCalcDiameter(ByVal shelf As Integer, ByVal Index As Integer, ByVal H As Double)
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

    R1 = RoundLocations(shelf, 1).diameter(1).Dist
    R12 = RoundLocations(shelf, 12).diameter(1).Dist
    
    Tetta = ACos((D ^ 2 - R1 ^ 2 - Rc(1) ^ 2) / (-2 * R1 * Rc(1))) / (2 * PI) * 360
    CurrRad = Sqr(R1 ^ 2 + H ^ 2 - 2 * R1 * H * Cos(Tetta / 360 * (2 * PI)))
    Omega = ACos(((H ^ 2) - (R1 ^ 2) - (CurrRad ^ 2)) / (-2 * R1 * CurrRad)) / (2 * PI) * 360
    
    ''name
    RoundLocations(shelf, 1).diameter(Index).name = CStr(shelf) & "01." & Index
    
    ''distance
    RoundLocations(shelf, 1).diameter(Index).Dist = CurrRad

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
    RoundLocations(shelf, 1).diameter(Index).Alfa = RoundLocations(shelf, 1).diameter(1).Alfa + Omega
    '' X
    If RoundLocations(shelf, 1).diameter(Index).Alfa > -90 Then
        RoundLocations(shelf, 1).diameter(Index).X = RoundLocations(shelf, 1).diameter(Index).Dist * Cos(RoundLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa = -90) Then
        RoundLocations(shelf, 1).diameter(Index).X = RoundLocations(shelf, 1).diameter(Index).Dist
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa < -90) Then
        RoundLocations(shelf, 1).diameter(Index).X = (-1) * RoundLocations(shelf, 1).diameter(Index).Dist * Cos((180 - RoundLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If
    
    ''calculate the y coordinate
    If RoundLocations(shelf, 1).diameter(Index).Alfa < 90 Then
        RoundLocations(shelf, 1).diameter(Index).Y = -1 * RoundLocations(shelf, 1).diameter(Index).Dist * Sin(RoundLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa = 90) Then
        RoundLocations(shelf, 1).diameter(Index).Y = -1 * RoundLocations(shelf, 1).diameter(Index).Dist
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa > 90) Then
        RoundLocations(shelf, 1).diameter(Index).Y = -1 * RoundLocations(shelf, 1).diameter(Index).Dist * Sin((RoundLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If

    ''Z
    RoundLocations(shelf, 1).diameter(Index).z = RoundLocations(shelf, 1).diameter(1).z
    
    ''Rx
    RoundLocations(shelf, 1).diameter(Index).Rx = RoundLocations(shelf, 1).diameter(1).Rx
    
    ''Ry
    RoundLocations(shelf, 1).diameter(Index).Ry = RoundLocations(shelf, 1).diameter(1).Ry
    
    ''Rz
    RoundLocations(shelf, 1).diameter(Index).Rz = RoundLocations(shelf, 1).diameter(1).Rz
    
    
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
    fdbk = MsgBox(vbCrLf & "error in RoundCalcDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundCalcDiameter")

End Sub

Public Sub RoundCalc5thDiameter(ByVal shelf As Integer, ByVal Rc As Double, ByVal Delta As Double)
''
Dim R5 As Double
Dim MyAlfa As Double


    ''radius
    R5 = Sqr(D ^ 2 + Rc ^ 2 - 2 * D * Rc * Cos((Delta) / 360 * 2 * PI))
    RoundLocations(shelf, 1).diameter(5).Dist = R5
    
    ''alfa
    MyAlfa = RoundLocations(shelf, 1).diameter(1).Alfa + 3.57
    RoundLocations(shelf, 1).diameter(5).Alfa = MyAlfa
    
    ''name
    RoundLocations(shelf, 1).diameter(5).name = CStr(shelf) & "01." & CStr(5)

    '' X
    If RoundLocations(shelf, 1).diameter(5).Alfa > -90 Then
        RoundLocations(shelf, 1).diameter(5).X = _
            RoundLocations(shelf, 1).diameter(5).Dist * Cos(RoundLocations(shelf, 1).diameter(5).Alfa * ((2 * PI) / (360)))
    ElseIf (RoundLocations(shelf, 1).diameter(5).Alfa = -90) Then
        RoundLocations(shelf, 1).diameter(5).X = RoundLocations(shelf, 1).diameter(5).Dist
    ElseIf (RoundLocations(shelf, 1).diameter(5).Alfa < -90) Then
        RoundLocations(shelf, 1).diameter(5).X = _
            (-1) * RoundLocations(shelf, 1).diameter(5).Dist * Cos((180 - RoundLocations(shelf, 1).diameter(5).Alfa) * ((2 * PI) / (360)))
    End If
    
    ''calculate the y coordinate
    If RoundLocations(shelf, 1).diameter(5).Alfa < 90 Then
        RoundLocations(shelf, 1).diameter(5).Y = _
            -1 * RoundLocations(shelf, 1).diameter(5).Dist * Sin(RoundLocations(shelf, 1).diameter(5).Alfa * ((2 * PI) / (360)))
    ElseIf (RoundLocations(shelf, 1).diameter(5).Alfa = 90) Then
        RoundLocations(shelf, 1).diameter(5).Y = -1 * RoundLocations(shelf, 1).diameter(5).Dist
    ElseIf (RoundLocations(shelf, 1).diameter(5).Alfa > 90) Then
        RoundLocations(shelf, 1).diameter(5).Y = _
            -1 * RoundLocations(shelf, 1).diameter(5).Dist * Sin((RoundLocations(shelf, 1).diameter(5).Alfa) * ((2 * PI) / (360)))
    End If


    ''Z
    RoundLocations(shelf, 1).diameter(5).z = RoundLocations(shelf, 1).diameter(1).z
    
    ''Rx
    RoundLocations(shelf, 1).diameter(5).Rx = RoundLocations(shelf, 1).diameter(1).Rx
    
    ''Ry
    RoundLocations(shelf, 1).diameter(5).Ry = RoundLocations(shelf, 1).diameter(1).Ry
    
    ''Rz
    RoundLocations(shelf, 1).diameter(5).Rz = RoundLocations(shelf, 1).diameter(1).Rz - 3.57
    
    Call SaveArray("RoundLocations")
    
End Sub
Public Sub RoundLoadConsts()

''
Dim fdbk As Integer

    H(1) = 0
    H(2) = 35
    H(3) = 65
    H(4) = 90
    H(5) = 0
    H(6) = 25
    H(7) = 55
    H(8) = 90
    
    Rc(1) = 1200
    Rc(2) = 1165
    Rc(3) = 1135
    Rc(4) = 1110
    Rc(5) = 1215
    Rc(6) = 1190
    Rc(7) = 1160
    Rc(8) = 1125
   
                
Exit Sub
error:
    fdbk = MsgBox(vbCrLf & "error in RoundLoadConsts() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundLoadConsts")

End Sub




Public Sub RoundCalcDiameter_2(ByVal shelf As Integer, ByVal Index As Integer, ByVal H As Double)
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

    R1 = RoundLocations(shelf, 1).diameter(5).Dist
    R12 = RoundLocations(shelf, 12).diameter(5).Dist
    
    Tetta = ACos((D ^ 2 - R1 ^ 2 - Rc(5) ^ 2) / (-2 * R1 * Rc(5))) / (2 * PI) * 360
    CurrRad = Sqr(R1 ^ 2 + H ^ 2 - 2 * R1 * H * Cos(Tetta / 360 * 2 * PI))
    Omega = ACos(((H ^ 2) - (R1 ^ 2) - (CurrRad ^ 2)) / (-2 * R1 * CurrRad)) / (2 * PI) * 360
    
    ''name
    RoundLocations(shelf, 1).diameter(Index).name = CStr(shelf) & "01." & Index
    
    ''distance
    RoundLocations(shelf, 1).diameter(Index).Dist = CurrRad

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
    RoundLocations(shelf, 1).diameter(Index).Alfa = RoundLocations(shelf, 1).diameter(5).Alfa + Omega
    
    '' X
    If RoundLocations(shelf, 1).diameter(Index).Alfa > -90 Then
        RoundLocations(shelf, 1).diameter(Index).X = RoundLocations(shelf, 1).diameter(Index).Dist * Cos(RoundLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa = -90) Then
        RoundLocations(shelf, 1).diameter(Index).X = RoundLocations(shelf, 1).diameter(Index).Dist
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa < -90) Then
        RoundLocations(shelf, 1).diameter(Index).X = (-1) * RoundLocations(shelf, 1).diameter(Index).Dist * Cos((180 - RoundLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If
    
    ''calculate the y coordinate
    If RoundLocations(shelf, 1).diameter(Index).Alfa < 90 Then
        RoundLocations(shelf, 1).diameter(Index).Y = -1 * RoundLocations(shelf, 1).diameter(Index).Dist * Sin(RoundLocations(shelf, 1).diameter(Index).Alfa * ((2 * PI) / (360)))
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa = 90) Then
        RoundLocations(shelf, 1).diameter(Index).Y = -1 * RoundLocations(shelf, 1).diameter(Index).Dist
    ElseIf (RoundLocations(shelf, 1).diameter(Index).Alfa > 90) Then
        RoundLocations(shelf, 1).diameter(Index).Y = -1 * RoundLocations(shelf, 1).diameter(Index).Dist * Sin((RoundLocations(shelf, 1).diameter(Index).Alfa) * ((2 * PI) / (360)))
    End If

    ''Z
    RoundLocations(shelf, 1).diameter(Index).z = RoundLocations(shelf, 1).diameter(1).z
    
    ''Rx
    RoundLocations(shelf, 1).diameter(Index).Rx = RoundLocations(shelf, 1).diameter(1).Rx
    
    ''Ry
    RoundLocations(shelf, 1).diameter(Index).Ry = RoundLocations(shelf, 1).diameter(1).Ry
    
    ''Rz
    RoundLocations(shelf, 1).diameter(Index).Rz = RoundLocations(shelf, 1).diameter(5).Rz
    
    
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
    fdbk = MsgBox(vbCrLf & "error in RoundCalcDiameter() " & vbCrLf & " the error is :" & Err.Description, vbExclamation, "RoundCalcDiameter")

End Sub



