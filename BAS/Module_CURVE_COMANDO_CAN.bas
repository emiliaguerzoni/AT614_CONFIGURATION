Attribute VB_Name = "Module_CURVE_COMANDO_CAN"
Option Explicit


Enum SetupConGrafico
    eOK_CURVE_COMANDO_CAN
    eError_CURVE_COMANDO_CAN
End Enum

'STATI DELL' MLT
Enum MltState
    Neutral = 0
    Extend = 1
    Retract = 2
    Float = 3
    Alarm = 4
End Enum

'STATI DELL' MLT UTILIZZATI NEL FILE DELLA CURVA DI COMANDO
Enum MltFileState
    FileExtendError = -4
    FileExtendFloat = -3
    FileExtendPercentual100 = -2
    FileExtendPercentual0 = -1
    FileNeutro = 0
    FileRetractPercentual0 = 1
    FileRetractPercentual100 = 2
    FileRetractFloat = 3
    FileRetractError = 4
End Enum

Type TagPoints
    x As Double
    y As Double
End Type

Type TagChannel
    Points(20) As TagPoints
End Type



' Return TRUE errore
Private Function WriteChannelSinglePoint(Channel As TagChannel, ByVal indice As Integer, ByVal x As Double, ByVal y As Double) As SetupConGrafico
    Dim bReturn As Boolean
    bReturn = False
    
    If (indice >= UBound(Channel.Points)) Then
        bReturn = True
    Else
        Channel.Points(indice).x = x
        Channel.Points(indice).y = y
    End If
    
    WriteChannelSinglePoint = bReturn
End Function


' Return TRUE errore
Private Function ReadChannelSinglePoint(Channel As TagChannel, ByVal indice As Integer, ByRef x As Double, ByRef y As Double) As SetupConGrafico
    Dim bReturn As Boolean
    bReturn = False

    If (indice >= UBound(Channel.Points)) Then
        bReturn = True
    Else
        x = Channel.Points(indice).x
        y = Channel.Points(indice).y
    End If
    ReadChannelSinglePoint = bReturn
End Function


' Return TRUE errore
Public Function Get_CURVECOMANDO_VALUE(Channel As TagChannel, ByVal CommandInVolts As Double, ByRef State As MltState, ByRef Percentual As Integer) As SetupConGrafico
    Dim i As Integer
    Dim FirstY As Integer
    Dim SecondY As Integer
    
    Dim bReturn As Boolean

    i = 0
    Do While (i < UBound(Channel.Points) - 1)
    
        If (CommandInVolts > Channel.Points(i).x And CommandInVolts <= Channel.Points(i + 1).x) Then
        
            Select Case (Channel.Points(i).y)
            Case FileExtendError
                State = Alarm
                Percentual = 0
            Case FileExtendFloat
                State = Float
                Percentual = 0
            Case FileExtendPercentual0
                State = Extend
                FirstY = 0
                If (Channel.Points(i + 1).y = FileExtendPercentual100) Then
                    SecondY = 250
                Else
                    SecondY = 0
                End If
                Percentual = GetAvcPercentual(Channel.Points(i).x, Channel.Points(i + 1).x, _
                                FirstY, SecondY, _
                                CommandInVolts)
            Case FileExtendPercentual100
                FirstY = 250
                If (Channel.Points(i + 1).y = FileExtendPercentual100) Then
                    SecondY = 250
                Else
                    SecondY = 0
                End If
                
                Percentual = GetAvcPercentual(Channel.Points(i).x, Channel.Points(i + 1).x, _
                                FirstY, SecondY, _
                                CommandInVolts)
                State = Extend
            Case FileNeutro
                State = Neutral
                Percentual = 0
            Case FileRetractPercentual0
                FirstY = 0
                If (Channel.Points(i + 1).y = FileRetractPercentual100) Then
                    SecondY = 250
                Else
                    SecondY = 0
                End If
                
                Percentual = GetAvcPercentual(Channel.Points(i).x, Channel.Points(i + 1).x, _
                                FirstY, SecondY, _
                                CommandInVolts)
                State = Retract
            Case FileRetractPercentual100
                FirstY = 250
                If (Channel.Points(i + 1).y = FileRetractPercentual100) Then
                    SecondY = 250
                Else
                    SecondY = 0
                End If
                
                Percentual = GetAvcPercentual(Channel.Points(i).x, Channel.Points(i + 1).x, _
                                FirstY, SecondY, _
                                CommandInVolts)
                State = Retract
            Case FileRetractFloat
                State = Float
                Percentual = 0
            Case FileRetractError
                State = Alarm
                Percentual = 0
            Case Else
                State = Alarm
                Percentual = 0
                bReturn = True
            End Select
            
            Exit Do
        End If
        
        i = i + 1
    Loop
    
    If (i >= UBound(Channel.Points) - 1) Then
        bReturn = True
    End If
    
        
    Get_CURVECOMANDO_VALUE = bReturn
End Function







'-------------------------------------------------------------------------------------
'funzione che trasforma il comando al MLT da valore in volt (0-5V) a valore in % (AVC)
'
' AVC = Auxiliary Valve Command
'-------------------------------------------------------------------------------------
Private Function GetAvcPercentual(ByVal X0 As Double, ByVal x1 As Double, ByVal Y0 As Double, ByVal Y1 As Double, ByVal x As Double) As Integer
    Dim m As Double
    Dim c As Double
    Dim y As Double
    Dim maxY As Double
    Dim minY As Double
    
    m = (Y1 - Y0) / (x1 - X0)
    c = Y0 - m * X0
    
    y = m * x + c
        
    If (Y1 > Y0) Then
        maxY = Y1
        minY = Y0
    Else
        maxY = Y0
        minY = Y1
    End If
    
       
    If (y > maxY) Then
        y = maxY
    ElseIf (y < minY) Then
        y = minY
    End If
    
    GetAvcPercentual = y
    
End Function


' Carica la configurazione del file delle curve di comando
' TRUE errore
Public Function Update_CURVECOMANDO(ByRef Channel As TagChannel, ByVal NomeFile As String) As Boolean
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim x As Double
    Dim y As Double
    Dim bReturn As Boolean

    Update_CURVECOMANDO = False

    Canal = FreeFile
    
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal
        
        Do While Not EOF(1)
            Line Input #Canal, stringa_tmp
            If (count >= 1) Then             ' salto la prima riga
                Dati = Split(stringa_tmp, ";")
                x = CDbl(Dati(0))
                y = CDbl(Dati(1))
                
                
                If (WriteChannelSinglePoint(Channel, count - 1, x, y) = eError_CURVE_COMANDO_CAN) Then
                    Update_CURVECOMANDO = True
                    Exit Function
                End If
            End If
            count = count + 1
        Loop
    End If
        
    Close Canal
    Exit Function
err:
    Debug.Print err.Description
    Update_CURVECOMANDO = True
    Close Canal
End Function

Public Sub Update_CURVECOMANDO_GRAPH(ByRef CWGraphImpostaCan As CWGraph, ByVal NomeFile As String)
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim x As Double
    Dim y As Double
    
    CWGraphImpostaCan.ClearData
    
    Canal = FreeFile
    
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal
        
        Do While Not EOF(1)
            Line Input #Canal, stringa_tmp
            If (count >= 1) Then             ' salto la prima riga
                Dati = Split(stringa_tmp, ";")
                x = CDbl(Dati(0))
                y = CDbl(Dati(1))
                CWGraphImpostaCan.ChartXvsY x, y
                
            End If
            count = count + 1
        Loop
    End If
        
    Close Canal
End Sub


