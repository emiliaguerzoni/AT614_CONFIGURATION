Attribute VB_Name = "Module_GENERIC_TROVANOMECOMPUTER"
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long



Public Function TrovaNomeComputer() As String
    On Error GoTo Err_TrovaNomeComputer
    Dim b
    Dim VarRIS As Long
    Dim NomeComputer As String * 145
    Dim lunghezzastringa As Long

    NomeComputer = String(256, 0)
    If (GetComputerName(NomeComputer, Len(NomeComputer)) <> 0) Then
        TrovaNomeComputer = Left$(NomeComputer, InStr(NomeComputer, vbNullChar) - 1)
    End If
       

Exit_TrovaNomeComputer:
    Exit Function

Err_TrovaNomeComputer:
    MsgBox err.Number & " " & err.Description
    Exit Function
End Function


