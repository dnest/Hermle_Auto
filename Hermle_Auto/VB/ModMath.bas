Attribute VB_Name = "ModMath"
Option Explicit

Dim Tetta(12) As Double




Public Sub HSKPocketInterpolation_original(shelf As Integer)

Dim i As Integer
Dim dDist As Double
Dim dAngle As Double
Dim dRx As Double
Dim dRy As Double
Dim dRz As Double
Dim dZ As Double
Dim dY As Double
Dim dX As Double


        
    ''builed the name of the pocket.
    HSKLocations(shelf, 1).name = CStr(shelf) & "01"
    HSKLocations(shelf, 10).name = CStr(shelf) & "10"
    For i = 2 To 9
        HSKLocations(shelf, i).name = CStr(shelf) & "0" & CStr(i)
    Next
        
    ''calculate the distance of the pocket from the RobotBase
    HSKLocations(shelf, 1).Dist = RootSquere(HSKLocations(shelf, 1).X, HSKLocations(shelf, 1).Y)
    HSKLocations(shelf, 10).Dist = RootSquere(HSKLocations(shelf, 10).X, HSKLocations(shelf, 10).Y)
    dDist = Abs((HSKLocations(shelf, 1).Dist - HSKLocations(shelf, 10).Dist)) / (TotalHSK - 1) ''calcule the DeltaRadius
    If HSKLocations(shelf, 10).Dist > HSKLocations(shelf, 1).Dist Then
        For i = 2 To 9 ''calculate the Dist
            HSKLocations(shelf, i).Dist = HSKLocations(shelf, i - 1).Dist + dDist
        Next
    ElseIf HSKLocations(shelf, 1).Dist > HSKLocations(shelf, 10).Dist Then
        For i = 2 To 9 ''calculate the Dist
            HSKLocations(shelf, i).Dist = HSKLocations(shelf, i - 1).Dist - dDist
        Next
    End If
    
 
     ''calculate the angle
    If HSKLocations(shelf, 1).X >= 0 Then ''calculate the Angle of the first Pocket from the X-axis.
        HSKLocations(shelf, 1).Alfa = Atn(Abs(HSKLocations(shelf, 1).Y) / HSKLocations(shelf, 1).X) * 360 / (2 * PI)
    ElseIf HSKLocations(shelf, 1).X < 0 Then
        HSKLocations(shelf, 1).Alfa = 180 - Atn(HSKLocations(shelf, 1).Y / HSKLocations(shelf, 1).X) * 360 / (2 * PI)
    End If
    
    If HSKLocations(shelf, 10).X >= 0 Then ''calculate the Angle of the 10_th Pocket from the X-axis.
        HSKLocations(shelf, 10).Alfa = Atn(HSKLocations(shelf, 10).Y / HSKLocations(shelf, 10).X) * 360 / (2 * PI)
    ElseIf HSKLocations(shelf, 10).X < 0 Then
        HSKLocations(shelf, 10).Alfa = 180 - Atn(Abs(HSKLocations(shelf, 10).Y) / Abs(HSKLocations(shelf, 10).X)) * 360 / (2 * PI)
    End If
    
    dAngle = Abs(HSKLocations(shelf, 10).Alfa - HSKLocations(shelf, 1).Alfa) / (TotalHSK - 1) ''calculate the DeltaAngle
    For i = 2 To 9
        HSKLocations(shelf, i).Alfa = HSKLocations(shelf, i - 1).Alfa + dAngle
    Next

    ''calculate the dRx
    dRx = Abs(HSKLocations(shelf, 10).Rx - HSKLocations(shelf, 1).Rx) / (TotalHSK - 1)
    
    If HSKLocations(shelf, 10).Rx > HSKLocations(shelf, 1).Rx Then
        For i = 2 To 9 ''calculate the Rx
            HSKLocations(shelf, i).Rx = HSKLocations(shelf, i - 1).Rx + dRx
        Next
    ElseIf HSKLocations(shelf, 1).Rx > HSKLocations(shelf, 10).Rx Then
        For i = 2 To 9 ''calculate the Rx
            HSKLocations(shelf, i).Rx = HSKLocations(shelf, i - 1).Rx - dRx
        Next
    ElseIf HSKLocations(shelf, 1).Rx = HSKLocations(shelf, 10).Rx Then
        For i = 2 To 9 ''calculate the Rx
            HSKLocations(shelf, i).Rx = HSKLocations(shelf, i - 1).Rx
        Next
    End If
    
    
    ''calculate the dRy
    dRy = Abs(HSKLocations(shelf, 10).Ry - HSKLocations(shelf, 1).Ry) / (TotalHSK - 1)
    
    If HSKLocations(shelf, 10).Ry > HSKLocations(shelf, 1).Ry Then
        For i = 2 To 9 ''calculate the Ry
            HSKLocations(shelf, i).Ry = HSKLocations(shelf, i - 1).Ry + dRy
        Next
    ElseIf HSKLocations(shelf, 1).Ry > HSKLocations(shelf, 10).Ry Then
        For i = 2 To 9 ''calculate the Ry
            HSKLocations(shelf, i).Ry = HSKLocations(shelf, i - 1).Ry - dRy
        Next
    ElseIf HSKLocations(shelf, 1).Ry = HSKLocations(shelf, 10).Ry Then
        For i = 2 To 9 ''calculate the Ry
            HSKLocations(shelf, i).Ry = HSKLocations(shelf, i - 1).Ry
        Next
    End If
    
    ''calculate the dRz
     dRz = 8
     For i = 2 To 9
        HSKLocations(shelf, i).Rz = HSKLocations(shelf, i - 1).Rz - dRz
     Next

    ''calculate the dZ parameter
    dZ = Abs(HSKLocations(shelf, 1).z - HSKLocations(shelf, 10).z) / (TotalHSK - 1)
    
     ''calculate the Z coordinate
    If HSKLocations(shelf, 10).z >= HSKLocations(shelf, 1).z Then
        For i = 2 To 9 ''calculate the Z
            HSKLocations(shelf, i).z = HSKLocations(shelf, i - 1).z + dZ
        Next
    ElseIf HSKLocations(shelf, 1).z > HSKLocations(shelf, 10).z Then
        For i = 2 To 9 ''calculate the Z
            HSKLocations(shelf, i).z = HSKLocations(shelf, i - 1).z - dZ
        Next
    End If
    
    ''calculate the X coordinate
    For i = 2 To 9
        If HSKLocations(shelf, i).Alfa < 90 Then
            HSKLocations(shelf, i).X = HSKLocations(shelf, i).Dist * Cos(HSKLocations(shelf, i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (HSKLocations(shelf, i).Alfa = 90) Then
            HSKLocations(shelf, i).X = 0
            
        ElseIf (HSKLocations(shelf, i).Alfa > 90) Then
            HSKLocations(shelf, i).X = (-1) * HSKLocations(shelf, i).Dist * Cos((180 - HSKLocations(shelf, i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 9
    
        If HSKLocations(shelf, i).Alfa < 90 Then
            HSKLocations(shelf, i).Y = -1 * HSKLocations(shelf, i).Dist * Sin(HSKLocations(shelf, i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (HSKLocations(shelf, i).Alfa = 90) Then
            HSKLocations(shelf, i).Y = -1 * HSKLocations(shelf, i).Dist
            
        ElseIf (HSKLocations(shelf, i).Alfa > 90) Then
            HSKLocations(shelf, i).Y = -1 * HSKLocations(shelf, i).Dist * Sin((HSKLocations(shelf, i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
      
End Sub

Public Sub HSKPocketInterpolation(shelf As Integer)

Dim R(10) As Double
Dim Lambda(10) As Double
Dim Gamma(10) As Double
Dim s(10) As Double
Dim Betta(10) As Double
Dim Tetta(10) As Double
Dim Sign(10) As Integer
Dim Sign_2 As Integer
Dim Alfa As Double
Dim Alfa_1 As Double
Dim D_Alfa As Double
Dim Gamma_c As Double
Dim Gamma_1 As Double
Dim D As Double
Dim C As Double
Dim i As Integer
Dim Rc As Double
Dim dRz As Double
Dim DeltaX As Double
Dim DeltaY As Double
Dim dZ As Double
Dim dRx As Double
Dim dRy As Double
Dim ret As Integer
Dim stam As Double


On Error GoTo error

    DeltaX = Abs(HSKLocations(shelf, 1).X) + Abs(HSKLocations(shelf, 10).X)
    DeltaY = Abs(HSKLocations(shelf, 1).Y) - Abs(HSKLocations(shelf, 10).Y)
    C = Sqr(DeltaX ^ 2 + DeltaY ^ 2)
    
    Rc = 1130
   
    ''Calculate the distance of the pocket from the RobotBase
    HSKLocations(shelf, 1).Dist = RootSquere(HSKLocations(shelf, 1).X, HSKLocations(shelf, 1).Y)
    HSKLocations(shelf, 10).Dist = RootSquere(HSKLocations(shelf, 10).X, HSKLocations(shelf, 10).Y)
    R(1) = HSKLocations(shelf, 1).Dist
    R(10) = HSKLocations(shelf, 10).Dist
    
    Gamma_c = ACos((-C ^ 2 + Rc ^ 2 + Rc ^ 2) / (2 * Rc * Rc)) / (2 * PI) * 360
    
    For i = 1 To 10
        s(i) = 2 * Rc * Sin(0.5 * ((Gamma_c) / (TotalHSK - 1)) * (i - 1) / 360 * (2 * PI))
    Next
    
    Alfa = (180 - Gamma_c) / 2
    Alfa_1 = ACos((R(1) ^ 2 + C ^ 2 - R(10) ^ 2) / (2 * R(1) * C)) / (2 * PI) * 360
    D_Alfa = Alfa - Alfa_1
    D = Sqr(Rc ^ 2 + R(1) ^ 2 - 2 * Rc * R(1) * Cos(D_Alfa / 360 * 2 * PI))
    Gamma_1 = ACos((Rc ^ 2 + D ^ 2 - R(1) ^ 2) / (2 * Rc * D)) / (2 * PI) * 360
    
    For i = 1 To 9
        Gamma(i) = ACos((-s(i) ^ 2 + Rc ^ 2 + Rc ^ 2) / (2 * Rc * Rc)) / (2 * PI) * 360
        If (((R(1) < Rc) And (R(10) < Rc)) Or ((R(1) > Rc) And (R(10) < Rc))) Then
            Lambda(i) = Gamma_1 - Gamma(i)
        ElseIf (((R(1) < Rc) And (R(10) > Rc)) Or ((R(1) > Rc) And (R(10) > Rc))) Then
            Lambda(i) = Gamma_1 + Gamma(i)
        End If
        
        R(i) = Sqr(D ^ 2 + Rc ^ 2 - 2 * D * Rc * Cos(Lambda(i) / 360 * 2 * PI))
        Betta(i) = ACos((R(1) ^ 2 + R(i) ^ 2 - s(i) ^ 2) / (2 * R(1) * R(i))) / (2 * PI) * 360
    Next
    
    ''assign value to the distance
    For i = 2 To 9
        HSKLocations(shelf, i).Dist = R(i)
    Next
    
    ''calculate the Angle of the first Pocket from the X-axis.
    If HSKLocations(shelf, 1).X >= 0 Then
        HSKLocations(shelf, 1).Alfa = Atn(Abs(HSKLocations(shelf, 1).Y) / HSKLocations(shelf, 1).X) * 360 / (2 * PI)
    ElseIf HSKLocations(shelf, 1).X < 0 Then
        HSKLocations(shelf, 1).Alfa = 180 - Atn(HSKLocations(shelf, 1).Y / HSKLocations(shelf, 1).X) * 360 / (2 * PI)
    End If
    
    ''calculate the Angle of the 10_th Pocket from the X-axis.
    If HSKLocations(shelf, 10).X >= 0 Then
        HSKLocations(shelf, 10).Alfa = Atn(HSKLocations(shelf, 10).Y / HSKLocations(shelf, 10).X) * 360 / (2 * PI)
    ElseIf HSKLocations(shelf, 10).X < 0 Then
        HSKLocations(shelf, 10).Alfa = 180 - Atn(Abs(HSKLocations(shelf, 10).Y) / Abs(HSKLocations(shelf, 10).X)) * 360 / (2 * PI)
    End If
    

    ''calculate alfa for all pockets.
    For i = 2 To 9
        HSKLocations(shelf, i).Alfa = Betta(i) + HSKLocations(shelf, 1).Alfa '''alfa(1)=~50.24
    Next


    ''calculate the dZ parameter
    dZ = Abs(HSKLocations(shelf, 1).z - HSKLocations(shelf, 10).z) / (TotalHSK - 1)
    
    ''calculate the Z coordinate
    If HSKLocations(shelf, 10).z >= HSKLocations(shelf, 1).z Then
        For i = 2 To 9 ''calculate the Z
            HSKLocations(shelf, i).z = HSKLocations(shelf, i - 1).z + dZ
        Next
    ElseIf HSKLocations(shelf, 1).z > HSKLocations(shelf, 10).z Then
        For i = 2 To 9 ''calculate the Z
            HSKLocations(shelf, i).z = HSKLocations(shelf, i - 1).z - dZ
        Next
    End If
    
    ''calculate the X coordinate
    For i = 2 To 9
        If HSKLocations(shelf, i).Alfa < 90 Then
            HSKLocations(shelf, i).X = HSKLocations(shelf, i).Dist * Cos(HSKLocations(shelf, i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (HSKLocations(shelf, i).Alfa = 90) Then
            HSKLocations(shelf, i).X = 0
            
        ElseIf (HSKLocations(shelf, i).Alfa > 90) Then
            HSKLocations(shelf, i).X = (-1) * HSKLocations(shelf, i).Dist * Cos((180 - HSKLocations(shelf, i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 9
        If HSKLocations(shelf, i).Alfa < 90 Then
            HSKLocations(shelf, i).Y = -1 * HSKLocations(shelf, i).Dist * Sin(HSKLocations(shelf, i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (HSKLocations(shelf, i).Alfa = 90) Then
            HSKLocations(shelf, i).Y = -1 * HSKLocations(shelf, i).Dist
            
        ElseIf (HSKLocations(shelf, i).Alfa > 90) Then
            HSKLocations(shelf, i).Y = -1 * HSKLocations(shelf, i).Dist * Sin((HSKLocations(shelf, i).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    
    ''builed the name of the pocket.
    HSKLocations(shelf, 1).name = CStr(shelf) & "01"
    HSKLocations(shelf, 10).name = CStr(shelf) & "10"
    For i = 2 To 9
        HSKLocations(shelf, i).name = CStr(shelf) & "0" & CStr(i)
    Next
      
      
    ''calculate the dRx
    dRx = Abs(HSKLocations(shelf, 10).Rx - HSKLocations(shelf, 1).Rx) / (TotalHSK - 1)
    
    ''calculate the Rx
    If HSKLocations(shelf, 10).Rx > HSKLocations(shelf, 1).Rx Then
        For i = 2 To 9
            HSKLocations(shelf, i).Rx = HSKLocations(shelf, i - 1).Rx + dRx
        Next
    ElseIf HSKLocations(shelf, 1).Rx > HSKLocations(shelf, 10).Rx Then
        For i = 2 To 9
            HSKLocations(shelf, i).Rx = HSKLocations(shelf, i - 1).Rx - dRx
        Next
    ElseIf HSKLocations(shelf, 1).Rx = HSKLocations(shelf, 10).Rx Then
        For i = 2 To 9
            HSKLocations(shelf, i).Rx = HSKLocations(shelf, i - 1).Rx
        Next
    End If
    
    
    ''calculate the dRy
    dRy = Abs(HSKLocations(shelf, 10).Ry - HSKLocations(shelf, 1).Ry) / (TotalHSK - 1)
    
    ''calculate the Ry
    If HSKLocations(shelf, 10).Ry > HSKLocations(shelf, 1).Ry Then
        For i = 2 To 9
            HSKLocations(shelf, i).Ry = HSKLocations(shelf, i - 1).Ry + dRy
        Next
    ElseIf HSKLocations(shelf, 1).Ry > HSKLocations(shelf, 10).Ry Then
        For i = 2 To 9
            HSKLocations(shelf, i).Ry = HSKLocations(shelf, i - 1).Ry - dRy
        Next
    ElseIf HSKLocations(shelf, 1).Ry = HSKLocations(shelf, 10).Ry Then
        For i = 2 To 9
            HSKLocations(shelf, i).Ry = HSKLocations(shelf, i - 1).Ry
        Next
    End If
    
    ''calculate Tetta
    For i = 1 To 9
        Tetta(i) = ACos((-(D ^ 2) + R(i) ^ 2 + Rc ^ 2) / (2 * R(i) * Rc)) / (2 * PI) * 360
    Next
    
    ''calculate the sign
    If ((R(1) > Rc) And (R(10) < Rc)) Then
        For i = 1 To 5
            Sign(i) = 1
        Next
        For i = 6 To 10
            Sign(i) = 1
        Next
        Sign_2 = 1
    
    ElseIf ((R(1) < Rc) And (R(10) < Rc)) Then
        For i = 1 To 5
            Sign(i) = 1
        Next
        For i = 6 To 10
            Sign(i) = -1
        Next
        Sign_2 = 1
        
    ElseIf ((R(1) < Rc) And (R(10) > Rc)) Then
        For i = 1 To 5
            Sign(i) = -1
        Next
        For i = 6 To 10
            Sign(i) = -1
        Next
        Sign_2 = -1
    
    ElseIf ((R(1) > Rc) And (R(10) > Rc)) Then
        For i = 1 To 5
            Sign(i) = -1
        Next
        For i = 6 To 10
            Sign(i) = 1
        Next
        Sign_2 = -1
    End If
    
    ''calculate the Rz
    For i = 2 To 9
        HSKLocations(shelf, i).Rz = 180 - HSKLocations(shelf, 1).Alfa - Betta(i) + Sign(i) * Tetta(i) + Sign_2 * D_Alfa
    Next


Exit Sub

error:
    ret = MsgBox("error  while calculate position in shelf number :" & shelf & vbCrLf _
    & "the error is :" & Err.Description & vbCrLf _
        , vbExclamation, "HSKPocketInterpolation()")
End Sub


Public Function RootSquere(ByVal aa As Double, ByVal bb As Double)
    
    aa = aa ^ 2
    bb = bb ^ 2
    RootSquere = Sqr(aa + bb)
    
End Function

Function ASin(Value As Double) As Double
    If Abs(Value) <> 1 Then
        ASin = Atn(Value / Sqr(1 - Value * Value))
    Else
        ASin = 1.5707963267949 * Sgn(Value)
    End If
End Function

' arc cosine
' error if NUMBER is outside the range [-1,1]

Function ACos(ByVal number As Double) As Double
    If Abs(number) <> 1 Then
        ACos = 1.5707963267949 - Atn(number / Sqr(1 - number * number))
    ElseIf number = -1 Then
        ACos = 3.14159265358979
    End If
    'elseif number=1 --> Acos=0 (implicit)
End Function


Public Sub DrillLocationsInterpolation_original(shelf As Integer)
''1.this function compute the parameters:x,y,z,Rx,Ry,Rz
''  of the FIRST pocket in every column.
''2.the function call to another function that compute the internal intepolation.

Dim i As Integer
Dim dDist As Double
Dim dAngle As Double
Dim dRx As Double
Dim dRy As Double
Dim dRz As Double
Dim dZ As Double
Dim dY As Double
Dim dX As Double
Dim PocketString As String
Dim digit As String
''
Dim column As Integer
Dim pocket As Integer


    ''builed the name of the first pocket in every column.
    For i = 1 To 12
        If i < 10 Then
             DrillLocations(shelf, i).diameter(1).name = CStr(shelf) & "0" & CStr(i) & ".1"
        ElseIf i >= 10 Then
            DrillLocations(shelf, i).diameter(1).name = CStr(shelf) & CStr(i) & ".1"
        End If
    Next
    
    ''calculate the distance of the pocket from the RobotBase
    DrillLocations(shelf, 1).diameter(1).Dist = RootSquere(DrillLocations(shelf, 1).diameter(1).X, DrillLocations(shelf, 1).diameter(1).Y)
    DrillLocations(shelf, 12).diameter(1).Dist = RootSquere(DrillLocations(shelf, 12).diameter(1).X, DrillLocations(shelf, 12).diameter(1).Y)
     
    dDist = Abs((DrillLocations(shelf, 1).diameter(1).Dist - DrillLocations(shelf, 12).diameter(1).Dist)) / (TotalDRILL - 1) ''calcule the DeltaRadius
    If DrillLocations(shelf, 12).diameter(1).Dist > DrillLocations(shelf, 1).diameter(1).Dist Then
        For i = 2 To 11 ''calculate the Dist
            DrillLocations(shelf, i).diameter(1).Dist = DrillLocations(shelf, i - 1).diameter(1).Dist + dDist
        Next
    ElseIf DrillLocations(shelf, 12).diameter(1).Dist < DrillLocations(shelf, 1).diameter(1).Dist Then
        For i = 2 To 11 ''calculate the Dist
           DrillLocations(shelf, i).diameter(1).Dist = DrillLocations(shelf, i - 1).diameter(1).Dist - dDist
        Next
    End If
    
     ''calculate the angle
    If DrillLocations(shelf, 1).diameter(1).X > 0 Then  ''calculate the Angle of the first Pocket from the X-axis.
        DrillLocations(shelf, 1).diameter(1).Alfa = Atn((DrillLocations(shelf, 1).diameter(1).Y) / DrillLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
    ElseIf DrillLocations(shelf, 1).diameter(1).X < 0 Then
        DrillLocations(shelf, 1).diameter(1).Alfa = 180 - Atn(DrillLocations(shelf, 1).diameter(1).Y / DrillLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
    End If
    
    If DrillLocations(shelf, 12).diameter(1).X >= 0 Then ''calculate the Angle of the 12_th Pocket from the X-axis.
        DrillLocations(shelf, 12).diameter(1).Alfa = Atn(DrillLocations(shelf, 12).diameter(1).Y / DrillLocations(shelf, 12).diameter(1).X) * 360 / (2 * PI)
    ElseIf DrillLocations(shelf, 12).diameter(1).X < 0 Then
        DrillLocations(shelf, 12).diameter(1).Alfa = (-1) * (180 - Atn((DrillLocations(shelf, 12).diameter(1).Y) / (DrillLocations(shelf, 12).diameter(1).X)) * 360 / (2 * PI))
    End If
    
    dAngle = Abs(DrillLocations(shelf, 1).diameter(1).Alfa - DrillLocations(shelf, 12).diameter(1).Alfa) / (TotalDRILL - 1) ''calculate the DeltaAngle
    For i = 2 To 11
        DrillLocations(shelf, i).diameter(1).Alfa = DrillLocations(shelf, i - 1).diameter(1).Alfa - dAngle
    Next
    
    ''calculate the dRx
    dRx = Abs(DrillLocations(shelf, 1).diameter(1).Rx - DrillLocations(shelf, 12).diameter(1).Rx) / (TotalDRILL - 1)
    
    If DrillLocations(shelf, 12).diameter(1).Rx > DrillLocations(shelf, 1).diameter(1).Rx Then
        For i = 2 To 11 ''calculate the Rx
            DrillLocations(shelf, i).diameter(1).Rx = DrillLocations(shelf, i - 1).diameter(1).Rx + dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx > DrillLocations(shelf, 12).diameter(1).Rx Then
        For i = 2 To 11 ''calculate the Rx
            DrillLocations(shelf, i).diameter(1).Rx = DrillLocations(shelf, i - 1).diameter(1).Rx - dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx = DrillLocations(shelf, 12).diameter(1).Rx Then
        For i = 2 To 11 ''calculate the Rx
            DrillLocations(shelf, i).diameter(1).Rx = DrillLocations(shelf, i - 1).diameter(1).Rx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(DrillLocations(shelf, 1).diameter(1).Ry - DrillLocations(shelf, 12).diameter(1).Ry) / (TotalDRILL - 1)
    If DrillLocations(shelf, 12).diameter(1).Ry > DrillLocations(shelf, 1).diameter(1).Ry Then
        For i = 2 To 11 ''calculate the Ry
            DrillLocations(shelf, i).diameter(1).Ry = DrillLocations(shelf, i - 1).diameter(1).Ry + dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry > DrillLocations(shelf, 12).diameter(1).Ry Then
        For i = 2 To 11 ''calculate the Ry
            DrillLocations(shelf, i).diameter(1).Ry = DrillLocations(shelf, i - 1).diameter(1).Ry - dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry = DrillLocations(shelf, 12).diameter(1).Ry Then
        For i = 2 To 11 ''calculate the Ry
            DrillLocations(shelf, i).diameter(1).Ry = DrillLocations(shelf, i - 1).diameter(1).Ry
        Next
    End If
    
    ''calculate the dRz
    dRz = 7.14
    For i = 2 To 11
        DrillLocations(shelf, i).diameter(1).Rz = DrillLocations(shelf, i - 1).diameter(1).Rz - dRz
    Next

    ''calculate the z
    dZ = Abs(DrillLocations(shelf, 1).diameter(1).z - DrillLocations(shelf, 12).diameter(1).z) / (TotalDRILL - 1)
    If DrillLocations(shelf, 12).diameter(1).z > DrillLocations(shelf, 1).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 12 pocket is higher then the 1.
            DrillLocations(shelf, i).diameter(1).z = DrillLocations(shelf, i - 1).diameter(1).z + dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z > DrillLocations(shelf, 12).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, i).diameter(1).z = DrillLocations(shelf, i - 1).diameter(1).z - dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z = DrillLocations(shelf, 12).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, i).diameter(1).z = DrillLocations(shelf, i - 1).diameter(1).z
        Next
    End If
    
    ''calculate the X coordinate
    For i = 2 To 11
        If DrillLocations(shelf, i).diameter(1).Alfa > -90 Then
            DrillLocations(shelf, i).diameter(1).X = DrillLocations(shelf, i).diameter(1).Dist * Cos(DrillLocations(shelf, i).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa = -90) Then
            DrillLocations(shelf, i).diameter(1).X = DrillLocations(shelf, i).diameter(1).Dist
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa < -90) Then
            DrillLocations(shelf, i).diameter(1).X = (-1) * DrillLocations(shelf, i).diameter(1).Dist * Cos((180 - DrillLocations(shelf, i).diameter(1).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 11
        If DrillLocations(shelf, i).diameter(1).Alfa > -90 Then
            DrillLocations(shelf, i).diameter(1).Y = DrillLocations(shelf, i).diameter(1).Dist * Sin(DrillLocations(shelf, i).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa = -90) Then
            DrillLocations(shelf, i).diameter(1).Y = DrillLocations(shelf, i).diameter(1).Dist
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa < -90) Then
            DrillLocations(shelf, i).diameter(1).Y = DrillLocations(shelf, i).diameter(1).Dist * Sin((180 - DrillLocations(shelf, i).diameter(1).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    pocket = 1
    ''For shelf = 1 To 3
        For column = 1 To 12
        
            If column >= 10 Then
                digit = ""
            Else
                digit = "0"
            End If
            
            PocketString = CStr(shelf) & digit & CStr(column) & "." & CStr(pocket)
            Call DrillPocketInterpolation(PocketString)
            
        Next
    ''Next
    
End Sub

Public Sub DrillLocationsInterpolation(shelf As Integer)
''1.this function perform calculation for specific shelf.

Dim MyColumn As Integer
Dim MyDiameter As Integer
Dim BoolDone As Boolean

   ''compute the location of the FIRST diameter in All Columns
    Call DrillCalculateFirstDiameter(shelf)
    
    ''compute the location of all other diameters in the shelf
    For MyColumn = 1 To 12
        For MyDiameter = 2 To 4
            BoolDone = DrillPocketsCalculation(shelf, MyColumn, MyDiameter)
            BoolDone = DrillPocketsCalculation(shelf, MyColumn, MyDiameter)
            BoolDone = DrillPocketsCalculation(shelf, MyColumn, MyDiameter)
            BoolDone = DrillPocketsCalculation(shelf, MyColumn, MyDiameter)
            BoolDone = DrillPocketsCalculation(shelf, MyColumn, MyDiameter)
            BoolDone = DrillPocketsCalculation(shelf, MyColumn, MyDiameter)
        Next
    Next

End Sub
    
Public Sub DrillCalculateFirstDiameter(ByVal shelf As Integer)
''1.the function compute the location [X,Y,Z,Rx,Ry,Rz]
''  of the FIRST diameter in every column in a specific shelf.

Dim R(12) As Double
Dim Lambda(12) As Double
Dim Gamma(12) As Double
Dim s(12) As Double
Dim Betta(12) As Double
Dim Sign(12) As Integer
Dim Sign_2 As Integer
Dim Alfa As Double
Dim Alfa_1 As Double
Dim D_Alfa As Double
Dim Gamma_c As Double
Dim Gamma_1 As Double
Dim D As Double
Dim C As Double
Dim i As Integer
Dim Rc As Double
Dim dRz As Double
Dim Delta_x As Double
Dim Delta_y As Double
Dim dZ As Double
Dim dRx As Double
Dim dRy As Double

 
    Delta_x = Abs(DrillLocations(shelf, 1).diameter(1).X) + Abs(DrillLocations(shelf, 12).diameter(1).X)
    Delta_y = Abs(Abs(DrillLocations(shelf, 1).diameter(1).Y) - Abs(DrillLocations(shelf, 12).diameter(1).Y))
    C = Sqr(Delta_x ^ 2 + Delta_y ^ 2)

    Rc = 1115

    ''calculate the distance of the pocket from the RobotBase
    DrillLocations(shelf, 1).diameter(1).Dist = RootSquere(DrillLocations(shelf, 1).diameter(1).X, DrillLocations(shelf, 1).diameter(1).Y)
    DrillLocations(shelf, 12).diameter(1).Dist = RootSquere(DrillLocations(shelf, 12).diameter(1).X, DrillLocations(shelf, 12).diameter(1).Y)
    R(1) = DrillLocations(shelf, 1).diameter(1).Dist
    R(12) = DrillLocations(shelf, 12).diameter(1).Dist
    
    Gamma_c = ACos((-C ^ 2 + Rc ^ 2 + Rc ^ 2) / (2 * Rc * Rc)) / (2 * PI) * 360 ''calculate the real value
    '''Gamma_c = 78.54 theoretic
    
    For i = 1 To 12
        s(i) = 2 * Rc * Sin(0.5 * ((Gamma_c) / (TotalDRILL - 1)) * (i - 1) / 360 * (2 * PI))
    Next
    
    Alfa = (180 - Gamma_c) / 2
    Alfa_1 = ACos((R(1) ^ 2 + s(12) ^ 2 - R(12) ^ 2) / (2 * R(1) * s(12))) / (2 * PI) * 360
    D_Alfa = Alfa - Alfa_1
    D = Sqr(Rc ^ 2 + R(1) ^ 2 - 2 * Rc * R(1) * Cos(D_Alfa / 360 * 2 * PI))
    Gamma_1 = ACos((Rc ^ 2 + D ^ 2 - R(1) ^ 2) / (2 * Rc * D)) / (2 * PI) * 360
    
    For i = 2 To 11
        Gamma(i) = ACos((-s(i) ^ 2 + Rc ^ 2 + Rc ^ 2) / (2 * Rc * Rc)) / (2 * PI) * 360
        If (((R(1) < Rc) And (R(12) < Rc)) Or ((R(1) > Rc) And (R(12) < Rc))) Then
            Lambda(i) = Gamma_1 - Gamma(i)
        ElseIf (((R(1) < Rc) And (R(12) > Rc)) Or ((R(1) > Rc) And (R(12) > Rc))) Then
            Lambda(i) = Gamma_1 + Gamma(i)
        End If
        
        R(i) = Sqr(D ^ 2 + Rc ^ 2 - 2 * D * Rc * Cos(Lambda(i) / 360 * 2 * PI))
        Betta(i) = ACos((R(1) ^ 2 + R(i) ^ 2 - s(i) ^ 2) / (2 * R(1) * R(i))) / (2 * PI) * 360
    Next
    
    ''assign value to the distance
    For i = 2 To 11
        DrillLocations(shelf, i).diameter(1).Dist = R(i)
    Next
    
    ''calculate the angle of the 1st pocket
    If DrillLocations(shelf, 1).diameter(1).X > 0 Then  ''calculate the Angle of the first Pocket from the X-axis.
        DrillLocations(shelf, 1).diameter(1).Alfa = Atn(Abs(DrillLocations(shelf, 1).diameter(1).Y) / DrillLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
    ElseIf DrillLocations(shelf, 1).diameter(1).X < 0 Then
        DrillLocations(shelf, 1).diameter(1).Alfa = 180 - Atn(DrillLocations(shelf, 1).diameter(1).Y / DrillLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
    End If
    
     ''calculate the angle of the 12st pocket
    If DrillLocations(shelf, 12).diameter(1).X >= 0 Then ''calculate the Angle of the 12_th Pocket from the X-axis.
        DrillLocations(shelf, 12).diameter(1).Alfa = Atn(Abs(DrillLocations(shelf, 12).diameter(1).Y) / DrillLocations(shelf, 12).diameter(1).X) * 360 / (2 * PI)
    ElseIf DrillLocations(shelf, 12).diameter(1).X < 0 Then
        DrillLocations(shelf, 12).diameter(1).Alfa = 1 * (180 - Atn((DrillLocations(shelf, 12).diameter(1).Y) / (DrillLocations(shelf, 12).diameter(1).X)) * 360 / (2 * PI))
    End If
    
    
    ''calculate alfa for all pockets.
    For i = 2 To 11
        DrillLocations(shelf, i).diameter(1).Alfa = Betta(i) + DrillLocations(shelf, 1).diameter(1).Alfa
    Next
    
    ''calculate the dZ parameter
    dZ = Abs(DrillLocations(shelf, 1).diameter(1).z - DrillLocations(shelf, 12).diameter(1).z) / (TotalDRILL - 1)
    
    ''calculate the Z coordinate
    If DrillLocations(shelf, 12).diameter(1).z > DrillLocations(shelf, 1).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 12 pocket is higher then the 1.
            DrillLocations(shelf, i).diameter(1).z = DrillLocations(shelf, i - 1).diameter(1).z + dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z > DrillLocations(shelf, 12).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, i).diameter(1).z = DrillLocations(shelf, i - 1).diameter(1).z - dZ
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).z = DrillLocations(shelf, 12).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            DrillLocations(shelf, i).diameter(1).z = DrillLocations(shelf, i - 1).diameter(1).z
        Next
    End If
    
    
    ''calculate the X coordinate
    For i = 2 To 11
        If DrillLocations(shelf, i).diameter(1).Alfa > -90 Then
            DrillLocations(shelf, i).diameter(1).X = DrillLocations(shelf, i).diameter(1).Dist * Cos(DrillLocations(shelf, i).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa = -90) Then
            DrillLocations(shelf, i).diameter(1).X = DrillLocations(shelf, i).diameter(1).Dist
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa < -90) Then
            DrillLocations(shelf, i).diameter(1).X = (-1) * DrillLocations(shelf, i).diameter(1).Dist * Cos((180 - DrillLocations(shelf, i).diameter(1).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    
  ''calculate the y coordinate
    For i = 2 To 11
        If DrillLocations(shelf, i).diameter(1).Alfa < 90 Then
            DrillLocations(shelf, i).diameter(1).Y = -1 * DrillLocations(shelf, i).diameter(1).Dist * Sin(DrillLocations(shelf, i).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa = 90) Then
            DrillLocations(shelf, i).diameter(1).Y = -1 * DrillLocations(shelf, i).diameter(1).Dist
            
        ElseIf (DrillLocations(shelf, i).diameter(1).Alfa > 90) Then
            DrillLocations(shelf, i).diameter(1).Y = -1 * DrillLocations(shelf, i).diameter(1).Dist * Sin((DrillLocations(shelf, i).diameter(1).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    

    ''calculate the dRx
    dRx = Abs(DrillLocations(shelf, 1).diameter(1).Rx - DrillLocations(shelf, 12).diameter(1).Rx) / (TotalDRILL - 1)
    
    ''calculate the Rx
    If DrillLocations(shelf, 12).diameter(1).Rx > DrillLocations(shelf, 1).diameter(1).Rx Then
        For i = 2 To 11
            DrillLocations(shelf, i).diameter(1).Rx = DrillLocations(shelf, i - 1).diameter(1).Rx + dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx > DrillLocations(shelf, 12).diameter(1).Rx Then
        For i = 2 To 11
            DrillLocations(shelf, i).diameter(1).Rx = DrillLocations(shelf, i - 1).diameter(1).Rx - dRx
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Rx = DrillLocations(shelf, 12).diameter(1).Rx Then
        For i = 2 To 11
            DrillLocations(shelf, i).diameter(1).Rx = DrillLocations(shelf, i - 1).diameter(1).Rx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(DrillLocations(shelf, 1).diameter(1).Ry - DrillLocations(shelf, 12).diameter(1).Ry) / (TotalDRILL - 1)
    
    ''calculate the Ry
    If DrillLocations(shelf, 12).diameter(1).Ry > DrillLocations(shelf, 1).diameter(1).Ry Then
        For i = 2 To 11
            DrillLocations(shelf, i).diameter(1).Ry = DrillLocations(shelf, i - 1).diameter(1).Ry + dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry > DrillLocations(shelf, 12).diameter(1).Ry Then
        For i = 2 To 11
            DrillLocations(shelf, i).diameter(1).Ry = DrillLocations(shelf, i - 1).diameter(1).Ry - dRy
        Next
    ElseIf DrillLocations(shelf, 1).diameter(1).Ry = DrillLocations(shelf, 12).diameter(1).Ry Then
        For i = 2 To 11
            DrillLocations(shelf, i).diameter(1).Ry = DrillLocations(shelf, i - 1).diameter(1).Ry
        Next
    End If
    
    ''calculate Tetta
    For i = 1 To 12
        Tetta(i) = ACos((-D ^ 2 + R(i) ^ 2 + Rc ^ 2) / (2 * R(i) * Rc)) / (2 * PI) * 360
    Next
    
    ''calculate the sign
    If ((R(1) > Rc) And (R(12) < Rc)) Then
        For i = 1 To 6
            Sign(i) = 1
        Next
        For i = 7 To 12
            Sign(i) = 1
        Next
        Sign_2 = 1
    
    ElseIf ((R(1) < Rc) And (R(10) < Rc)) Then
        For i = 1 To 6
            Sign(i) = 1
        Next
        For i = 7 To 12
            Sign(i) = -1
        Next
        Sign_2 = 1
        
    ElseIf ((R(1) < Rc) And (R(12) > Rc)) Then
        For i = 1 To 6
            Sign(i) = -1
        Next
        For i = 7 To 12
            Sign(i) = -1
        Next
        Sign_2 = -1
    
    ElseIf ((R(1) > Rc) And (R(12) > Rc)) Then
        For i = 1 To 6
            Sign(i) = -1
        Next
        For i = 7 To 12
            Sign(i) = 1
        Next
        Sign_2 = -1
    End If
    
    ''calculate the Rz
    For i = 2 To 11
        DrillLocations(shelf, i).diameter(1).Rz = 180 - DrillLocations(shelf, 1).diameter(1).Alfa - Betta(i) + Sign(i) * Tetta(i) + Sign_2 * D_Alfa
    Next

End Sub
Public Function DrillPocketsCalculation(ByVal MyShelf As Integer, ByVal MyColumn As Integer, ByVal MyDiameter As Integer) As Boolean
''
''1.this function compute location [X,Y,Z,Rx,Ry,Rz] for a specific Diameter in entire shelf.
''2.the function receive the shelf number,the column number and the diameter number.

Dim InitRad As Double
Dim CurrentRad As Double
Dim Omega As Double
Dim H(7) As Double
Dim shelf As Integer
Dim column As Integer
Dim digit As String
Dim i As Integer

Dim aa As Double
Dim bb As Double


    On Error GoTo label
    shelf = MyShelf
    column = MyColumn
    ''constants.internal geometry.
    H(1) = 0
    H(2) = 22.5
    H(3) = 41.5
    H(4) = 58
    H(5) = 10
    H(6) = 41
    H(7) = 69
        
    InitRad = DrillLocations(shelf, MyColumn).diameter(1).Dist       '''the radius to the first diameter
    CurrentRad = Sqr(InitRad ^ 2 + H(MyDiameter) ^ 2 - 2 * InitRad * H(MyDiameter) * Cos((180 - Tetta(MyDiameter)) / 360 * 2 * PI))
    
    aa = (-H(MyDiameter) ^ 2 + InitRad ^ 2 + CurrentRad ^ 2)
    bb = (2 * InitRad * CurrentRad)
    
    Omega = ACos((-H(MyDiameter) ^ 2 + InitRad ^ 2 + CurrentRad ^ 2) / (2 * InitRad * CurrentRad)) / (2 * PI) * 360
    
    DrillLocations(shelf, MyColumn).diameter(MyDiameter).Dist = CurrentRad
    DrillLocations(shelf, MyColumn).diameter(MyDiameter).Alfa = _
        DrillLocations(shelf, MyColumn).diameter(MyDiameter).Alfa + Omega
   
    If CInt(column) <= 9 Then
        digit = "0"
    ElseIf CInt(column) >= 10 Then
        digit = ""
    End If
    

    ''name
    For i = 1 To 7
        DrillLocations(shelf, column).diameter(i).name = CStr(shelf) & digit & CStr(column) & "." & CStr(i)
    Next
            
    ''calculate the X coordinate
    For i = 2 To 7
        If DrillLocations(shelf, column).diameter(i).Alfa > -90 Then
            DrillLocations(shelf, column).diameter(i).X = DrillLocations(shelf, column).diameter(i).Dist * Cos(DrillLocations(shelf, column).diameter(i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa = -90) Then
            DrillLocations(shelf, column).diameter(i).X = DrillLocations(shelf, column).diameter(i).Dist
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa < -90) Then
            DrillLocations(shelf, column).diameter(i).X = (-1) * DrillLocations(shelf, column).diameter(i).Dist * Cos((180 - DrillLocations(shelf, column).diameter(i).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 7
        If DrillLocations(shelf, column).diameter(i).Alfa < 90 Then
            DrillLocations(shelf, column).diameter(i).Y = -1 * DrillLocations(shelf, column).diameter(i).Dist * Sin(DrillLocations(shelf, column).diameter(i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa = 90) Then
            DrillLocations(shelf, column).diameter(i).Y = 0
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa > 90) Then
            DrillLocations(shelf, column).diameter(i).Y = -1 * DrillLocations(shelf, column).diameter(i).Dist * Sin((DrillLocations(shelf, column).diameter(i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    ''*** Z ***
    For i = 2 To 7
        DrillLocations(shelf, column).diameter(i).z = DrillLocations(shelf, column).diameter(1).z
    Next
    
    ''*** Rx ***
    For i = 2 To 7
        DrillLocations(shelf, column).diameter(i).Rx = DrillLocations(shelf, column).diameter(1).Rx
    Next
    
    ''*** Ry ***
    For i = 2 To 7
       DrillLocations(shelf, column).diameter(i).Ry = DrillLocations(shelf, column).diameter(1).Ry
    Next
    
    ''*** Rz ***
    For i = 2 To 7
        DrillLocations(shelf, column).diameter(i).Rz = DrillLocations(shelf, column).diameter(1).Rz
    Next
    
    DrillPocketsCalculation = True
    Exit Function
    
label:
    '''MsgBox " Error in DrillPocketsCalculation" & Err.Description
    DrillPocketsCalculation = False
    
End Function

Public Sub DrillPocketInterpolation(ByVal PocketNumber As String)
''this function compute the parameters of each pocket
''in each column according to the parameters of HOLE number 1.

Dim Beta As Double
Dim dRadius(8) As Double
Dim i As Integer
Dim column As Integer
Dim shelf As Integer
Dim digit As String
Dim temp1 As Double
Dim temp2 As Double

On Error GoTo label
    ''constants. internal geometry.
    dRadius(1) = 0
    dRadius(2) = 22.5
    dRadius(3) = 41.5
    dRadius(4) = 58
    dRadius(5) = 10
    dRadius(6) = 41
    dRadius(7) = 69
    

    shelf = CInt(Left(PocketNumber, 1))
    column = CInt(Left(PocketNumber, 3))
    column = Right(column, 2)
    
    If CInt(column) <= 9 Then
        digit = "0"
    ElseIf CInt(column) >= 10 Then
        digit = ""
    End If
    

    ''name
    For i = 1 To 7
        DrillLocations(shelf, column).diameter(i).name = CStr(shelf) & digit & CStr(column) & "." & CStr(i)
    Next
            
    DrillLocations(shelf, column).diameter(2).Dist = 1

    ''calculate the X coordinate
    For i = 2 To 7
        If DrillLocations(shelf, column).diameter(i).Alfa > -90 Then
            DrillLocations(shelf, column).diameter(i).X = DrillLocations(shelf, column).diameter(i).Dist * Cos(DrillLocations(shelf, column).diameter(i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa = -90) Then
            DrillLocations(shelf, column).diameter(i).X = DrillLocations(shelf, column).diameter(i).Dist
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa < -90) Then
            DrillLocations(shelf, column).diameter(i).X = (-1) * DrillLocations(shelf, column).diameter(i).Dist * Cos((180 - DrillLocations(shelf, column).diameter(i).Alfa) * ((2 * PI) / (360)))
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 7
        If DrillLocations(shelf, column).diameter(i).Alfa < 90 Then
            DrillLocations(shelf, column).diameter(i).Y = -1 * DrillLocations(shelf, column).diameter(i).Dist * Sin(DrillLocations(shelf, column).diameter(i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa = 90) Then
            DrillLocations(shelf, column).diameter(i).Y = 0
            
        ElseIf (DrillLocations(shelf, column).diameter(i).Alfa > 90) Then
            DrillLocations(shelf, column).diameter(i).Y = -1 * DrillLocations(shelf, column).diameter(i).Dist * Sin((DrillLocations(shelf, column).diameter(i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    
    ''*** Rx ***
    For i = 2 To 7
        DrillLocations(shelf, column).diameter(i).Rx = DrillLocations(shelf, column).diameter(1).Rx
    Next
    
    ''*** Ry ***
    For i = 2 To 7
       DrillLocations(shelf, column).diameter(i).Ry = DrillLocations(shelf, column).diameter(1).Ry
    Next
    
    ''*** Rz ***
    For i = 2 To 7
        DrillLocations(shelf, column).diameter(i).Rz = DrillLocations(shelf, column).diameter(1).Rz
    Next
    
    ''*** Z ***
    For i = 2 To 7
        DrillLocations(shelf, column).diameter(i).z = DrillLocations(shelf, column).diameter(1).z
    Next
    
    
    
    
    Exit Sub
label:
    MsgBox "error in localterpolation"
    
End Sub


Public Sub DrillAllColumnSetDefault()

Dim i As Integer
Dim shelf As Integer
Dim column As Integer

    For shelf = 1 To 3
        For column = 1 To 12
             Call DrillPocketSetDefault(shelf, column)
        Next
    Next
     
End Sub

Public Sub DrillPocketSetDefault(shelf As Integer, column As Integer)

Dim i As Integer
Dim digit As String

    If column <= 9 Then
        digit = "0"
    ElseIf column >= 10 Then
        digit = ""
    End If
    

    For i = 1 To 7
    
        DrillLocations(shelf, column).diameter(i).name = CStr(shelf) & digit & CStr(column) & "." & CStr(i)
        
        DrillLocations(shelf, column).diameter(i).X = 0
        DrillLocations(shelf, column).diameter(i).Y = 0
        DrillLocations(shelf, column).diameter(i).z = 0
        
        DrillLocations(shelf, column).diameter(i).Rx = 0
        DrillLocations(shelf, column).diameter(i).Ry = 0
        DrillLocations(shelf, column).diameter(i).Rz = 0
        
        DrillLocations(shelf, column).diameter(i).Dist = 0
        DrillLocations(shelf, column).diameter(i).Alfa = 0
        
    Next
    
End Sub



Public Sub RoundLocationsInterpolation(shelf As Integer)
''1.this function compute the parameters:x,y,z,Rx,Ry,Rz
''  of the FIRST pocket in every column in RoundTool Shelf.
''2.the function call to another function that compute the internal locations of RoundTool Column.

Dim i As Integer
Dim dDist As Double
Dim dAngle As Double
Dim dRx As Double
Dim dRy As Double
Dim dRz As Double
Dim dZ As Double
Dim dY As Double
Dim dX As Double
Dim PocketString As String
Dim digit As String
''
Dim column As Integer
Dim pocket As Integer


    ''builed the name of the first pocket in every column.
    For column = 1 To 12
    
        If column < 10 Then
             RoundLocations(shelf, column).diameter(1).name = CStr(shelf) & "0" & CStr(column) & ".1"
        ElseIf column >= 10 Then
            RoundLocations(shelf, column).diameter(1).name = CStr(shelf) & CStr(column) & ".1"
        End If
        
    Next
    
    ''calculate the distance of the (first and the last) pocket from the RobotBase
    RoundLocations(shelf, 1).diameter(1).Dist = RootSquere(RoundLocations(shelf, 1).diameter(1).X, RoundLocations(shelf, 1).diameter(1).Y)
    RoundLocations(shelf, 12).diameter(1).Dist = RootSquere(RoundLocations(shelf, 12).diameter(1).X, RoundLocations(shelf, 12).diameter(1).Y)
     
     ''calcule the DeltaRadius
    dDist = Abs((RoundLocations(shelf, 1).diameter(1).Dist - RoundLocations(shelf, 12).diameter(1).Dist)) / (TotalROUND - 1)
    If RoundLocations(shelf, 12).diameter(1).Dist > RoundLocations(shelf, 1).diameter(1).Dist Then
        For i = 2 To 11 ''calculate the Dist
            RoundLocations(shelf, i).diameter(1).Dist = RoundLocations(shelf, i - 1).diameter(1).Dist + dDist
        Next
    ElseIf RoundLocations(shelf, 12).diameter(1).Dist < RoundLocations(shelf, 1).diameter(1).Dist Then
        For i = 2 To 11 ''calculate the Dist
           RoundLocations(shelf, i).diameter(1).Dist = RoundLocations(shelf, i - 1).diameter(1).Dist - dDist
        Next
    End If
    
     ''***calculate the angle***
     ''calculate the Angle of the 1_st Pocket from the X-axis.
    If RoundLocations(shelf, 1).diameter(1).X > 0 Then
        RoundLocations(shelf, 1).diameter(1).Alfa = Atn((RoundLocations(shelf, 1).diameter(1).Y) / RoundLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
    ElseIf RoundLocations(shelf, 1).diameter(1).X < 0 Then
        RoundLocations(shelf, 1).diameter(1).Alfa = 180 - Atn(RoundLocations(shelf, 1).diameter(1).Y / RoundLocations(shelf, 1).diameter(1).X) * 360 / (2 * PI)
    End If
    
    ''calculate the Angle of the 12_th Pocket from the X-axis.
    If RoundLocations(shelf, 12).diameter(1).X >= 0 Then
        RoundLocations(shelf, 12).diameter(1).Alfa = Atn(RoundLocations(shelf, 12).diameter(1).Y / RoundLocations(shelf, 12).diameter(1).X) * 360 / (2 * PI)
    ElseIf RoundLocations(shelf, 12).diameter(1).X < 0 Then
        RoundLocations(shelf, 12).diameter(1).Alfa = (-1) * (180 - Atn((RoundLocations(shelf, 12).diameter(1).Y) / (RoundLocations(shelf, 12).diameter(1).X)) * 360 / (2 * PI))
    End If
    
    dAngle = Abs(RoundLocations(shelf, 1).diameter(1).Alfa - RoundLocations(shelf, 12).diameter(1).Alfa) / (TotalROUND - 1) ''calculate the DeltaAngle
    For i = 2 To 11
        RoundLocations(shelf, i).diameter(1).Alfa = RoundLocations(shelf, i - 1).diameter(1).Alfa - dAngle
    Next
    
    ''calculate the dRx
    dRx = Abs(RoundLocations(shelf, 1).diameter(1).Rx - RoundLocations(shelf, 12).diameter(1).Rx) / (TotalROUND - 1)
    
    If RoundLocations(shelf, 12).diameter(1).Rx > RoundLocations(shelf, 1).diameter(1).Rx Then
        For i = 2 To 11 ''calculate the Rx
            RoundLocations(shelf, i).diameter(1).Rx = RoundLocations(shelf, i - 1).diameter(1).Rx + dRx
        Next
        
    ElseIf RoundLocations(shelf, 1).diameter(1).Rx = RoundLocations(shelf, 12).diameter(1).Rx Then
        For i = 2 To 11 ''calculate the Rx
            RoundLocations(shelf, i).diameter(1).Rx = RoundLocations(shelf, i - 1).diameter(1).Rx
        Next
        
        
    ElseIf RoundLocations(shelf, 1).diameter(1).Rx > RoundLocations(shelf, 12).diameter(1).Rx Then
        For i = 2 To 11 ''calculate the Rx
            RoundLocations(shelf, i).diameter(1).Rx = RoundLocations(shelf, i - 1).diameter(1).Rx - dRx
        Next
    End If
    
    ''calculate the dRy
    dRy = Abs(RoundLocations(shelf, 1).diameter(1).Ry - RoundLocations(shelf, 12).diameter(1).Ry) / (TotalROUND - 1)
    If RoundLocations(shelf, 12).diameter(1).Ry > RoundLocations(shelf, 1).diameter(1).Ry Then
        For i = 2 To 11 ''calculate the Ry
            RoundLocations(shelf, i).diameter(1).Ry = RoundLocations(shelf, i - 1).diameter(1).Ry + dRy
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Ry > RoundLocations(shelf, 12).diameter(1).Ry Then
        For i = 2 To 11 ''calculate the Ry
            RoundLocations(shelf, i).diameter(1).Ry = RoundLocations(shelf, i - 1).diameter(1).Ry - dRy
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).Ry = RoundLocations(shelf, 12).diameter(1).Ry Then
        For i = 2 To 11 ''calculate the Ry
            RoundLocations(shelf, i).diameter(1).Ry = RoundLocations(shelf, i - 1).diameter(1).Ry
        Next
    End If
    
    ''******
    ''  dZ
    ''******
    
    dRz = 7.14
    For i = 2 To 11
        RoundLocations(shelf, i).diameter(1).Rz = RoundLocations(shelf, i - 1).diameter(1).Rz - dRz
    Next
    
    ''calculate the z
    dZ = Abs(RoundLocations(shelf, 1).diameter(1).z - RoundLocations(shelf, 12).diameter(1).z) / (TotalROUND - 1)
    If RoundLocations(shelf, 12).diameter(1).z > RoundLocations(shelf, 1).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 12 pocket is higher then the 1.
            RoundLocations(shelf, i).diameter(1).z = RoundLocations(shelf, i - 1).diameter(1).z + dZ
        Next
    ElseIf RoundLocations(shelf, 1).diameter(1).z >= RoundLocations(shelf, 12).diameter(1).z Then
        For i = 2 To 11 ''calculate the z if the 1 pocket is higher then the 12.
            RoundLocations(shelf, i).diameter(1).z = RoundLocations(shelf, i - 1).diameter(1).z - dZ
        Next
    End If
    
    
    ''calculate the X coordinate
    For i = 2 To 11
        If RoundLocations(shelf, i).diameter(1).Alfa > -90 Then
            RoundLocations(shelf, i).diameter(1).X = RoundLocations(shelf, i).diameter(1).Dist * Cos(RoundLocations(shelf, i).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, i).diameter(1).Alfa = -90) Then
            RoundLocations(shelf, i).diameter(1).X = RoundLocations(shelf, i).diameter(1).Dist
            
        ElseIf (RoundLocations(shelf, i).diameter(1).Alfa < -90) Then
            RoundLocations(shelf, i).diameter(1).X = (-1) * RoundLocations(shelf, i).diameter(1).Dist * Cos((180 - RoundLocations(shelf, i).diameter(1).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 11
        If RoundLocations(shelf, i).diameter(1).Alfa > -90 Then
            RoundLocations(shelf, i).diameter(1).Y = RoundLocations(shelf, i).diameter(1).Dist * Sin(RoundLocations(shelf, i).diameter(1).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, i).diameter(1).Alfa = -90) Then
            RoundLocations(shelf, i).diameter(1).Y = RoundLocations(shelf, i).diameter(1).Dist
            
        ElseIf (RoundLocations(shelf, i).diameter(1).Alfa < -90) Then
            RoundLocations(shelf, i).diameter(1).Y = RoundLocations(shelf, i).diameter(1).Dist * Sin((180 - RoundLocations(shelf, i).diameter(1).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    pocket = 1
    ''For shelf = 1 To 3
        For column = 1 To TotalROUND
        
            If column >= 10 Then
                digit = ""
            Else
                digit = "0"
            End If
            
            PocketString = CStr(shelf) & digit & CStr(column) & "." & CStr(pocket)
            Call RoundPocketInterpolation(PocketString)
            
        Next
    ''Next


End Sub





Public Sub RoundPocketInterpolation(ByVal PocketNumber As String)
''1.this function compute the 6 Degree-of-freedom of each pocket
''  in each column according to the parameters of pocket number 1.
''2.the function recieve the name of the first pocket in the column.


Dim Beta As Double
Dim dRadius(9) As Double
Dim i As Integer
Dim column As Integer
Dim shelf As Integer
Dim digit As String


    ''constants. internal geometry.
    dRadius(1) = 0
    dRadius(2) = 25
    dRadius(3) = 55
    dRadius(4) = 90
    dRadius(5) = 15
    dRadius(6) = 50
    dRadius(7) = 80
    dRadius(8) = 105
    

    shelf = CInt(Left(PocketNumber, 1))
    
    column = CInt(Left(PocketNumber, 3))
    column = Right(column, 2)
    
    If CInt(column) <= 9 Then
        digit = "0"
    ElseIf CInt(column) >= 10 Then
        digit = ""
    End If
    
    Beta = (RoundLocations(shelf, 12).diameter(1).Alfa - RoundLocations(shelf, 1).diameter(1).Alfa) / (2 * (TotalROUND - 1))

    ''name
    For i = 2 To 8
        RoundLocations(shelf, column).diameter(i).name = CStr(shelf) & digit & CStr(column) & "." & CStr(i)
    Next
    
    ''distance
    For i = 2 To 8
        RoundLocations(shelf, column).diameter(i).Dist = RoundLocations(shelf, column).diameter(1).Dist + dRadius(i)
    Next
        
    ''angle
    For i = 2 To 8
        If i <= 4 Then
            RoundLocations(shelf, column).diameter(i).Alfa = RoundLocations(shelf, column).diameter(1).Alfa
        ElseIf i >= 5 Then
            RoundLocations(shelf, column).diameter(i).Alfa = RoundLocations(shelf, column).diameter(1).Alfa + Beta
        End If
    Next
        
    ''Rx
    For i = 2 To 8
        RoundLocations(shelf, column).diameter(i).Rx = RoundLocations(shelf, column).diameter(1).Rx
    Next
    
    ''Ry
    For i = 2 To 8
       RoundLocations(shelf, column).diameter(i).Ry = RoundLocations(shelf, column).diameter(1).Ry
    Next
    
    ''Rz
    For i = 2 To 8
        RoundLocations(shelf, column).diameter(i).Rz = RoundLocations(shelf, column).diameter(1).Rz
    Next
    

    ''Z
    For i = 2 To 8
        RoundLocations(shelf, column).diameter(i).z = RoundLocations(shelf, column).diameter(1).z
    Next
    
    ''calculate the X coordinate
    For i = 2 To 8
        If RoundLocations(shelf, column).diameter(i).Alfa > -90 Then
            RoundLocations(shelf, column).diameter(i).X = RoundLocations(shelf, column).diameter(i).Dist * Cos(RoundLocations(shelf, column).diameter(i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, column).diameter(i).Alfa = -90) Then
            RoundLocations(shelf, column).diameter(i).X = RoundLocations(shelf, column).diameter(i).Dist
            
        ElseIf (RoundLocations(shelf, column).diameter(i).Alfa < -90) Then
            RoundLocations(shelf, column).diameter(i).X = (-1) * RoundLocations(shelf, column).diameter(i).Dist * Cos((180 - RoundLocations(shelf, column).diameter(i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
    ''calculate the y coordinate
    For i = 2 To 8
        If RoundLocations(shelf, column).diameter(i).Alfa > -90 Then
            RoundLocations(shelf, column).diameter(i).Y = RoundLocations(shelf, column).diameter(i).Dist * Sin(RoundLocations(shelf, column).diameter(i).Alfa * ((2 * PI) / (360)))
            
        ElseIf (RoundLocations(shelf, column).diameter(i).Alfa = -90) Then
            RoundLocations(shelf, column).diameter(i).Y = 0
            
        ElseIf (RoundLocations(shelf, column).diameter(i).Alfa < -90) Then
            RoundLocations(shelf, column).diameter(i).Y = RoundLocations(shelf, column).diameter(i).Dist * Sin((180 - RoundLocations(shelf, column).diameter(i).Alfa) * ((2 * PI) / (360)))
            
        End If
    Next
    
End Sub















