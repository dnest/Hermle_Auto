Attribute VB_Name = "ModLogFile"
Option Explicit
Public Sub LogAddLine(ByVal MyString As String)

Dim ii As Integer

    If UseExternalFile = False Then
        Exit Sub
    End If

    For ii = 1 To 99
        StringLog(ii) = StringLog(ii + 1)
    Next
    StringLog(100) = MyString & "," & CStr(Format(now, "hh:mm:ss"))
    
End Sub


Public Sub WriteLogFile()

''1.the function save the LogFile into the HardDisk.
''2.the function save the last 100 acts.

Dim LogFilePath As String
Dim filenumber As Integer
Dim Today As Date
Dim now As Date
Dim country As String
Dim Factory As String
Dim MachineNumber As Integer
Dim jj As Integer


        LogFilePath = App.path & "\WorkDirectory\Data\"
        LogFilePath = LogFilePath & "LogFile.csv"
        
        filenumber = FreeFile
        Open LogFilePath For Output As #filenumber
    
        For jj = 1 To 100
            Print #filenumber, StringLog(jj)
        Next
        
        Close filenumber
End Sub
