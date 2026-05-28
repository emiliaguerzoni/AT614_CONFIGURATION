Attribute VB_Name = "Module_CURVE_XY"
Option Explicit


Enum SetupConGraficoRampa
    eOK_Rampe
    eERROR_Rampe
End Enum



Type TagPointsRampa
    X As Double
    Y As Double
End Type

Type TagChannelRampa
    PointsNumber As Integer
    Points(20) As TagPointsRampa
End Type



' Return TRUE errore
Public Function WriteChannelSinglePointRampa(Channel As TagChannelRampa, ByVal indice As Integer, ByVal X As Double, ByVal Y As Double) As SetupConGraficoRampa
    Dim bReturn As Boolean
    bReturn = False
    
    If (indice >= UBound(Channel.Points)) Then
        bReturn = True
    Else
        Channel.Points(indice).X = X
        Channel.Points(indice).Y = Y
    End If
    
    WriteChannelSinglePointRampa = bReturn
End Function


' Return TRUE errore
Public Function AddChannelSinglePointRampa(ByRef Channel As TagChannelRampa, ByVal X As Double, ByVal Y As Double) As SetupConGraficoRampa
    Dim bReturn As SetupConGraficoRampa
    Dim indice As Integer
    
    bReturn = eOK_Rampe
    
    
    indice = Channel.PointsNumber
    
    Channel.Points(indice).X = X
    Channel.Points(indice).Y = Y

    Channel.PointsNumber = Channel.PointsNumber + 1
    AddChannelSinglePointRampa = bReturn
    Exit Function
err:
    AddChannelSinglePointRampa = eERROR_Rampe
End Function


Public Function ReadGraphPoints(ByRef Channel As TagChannelRampa, ByRef Index As Integer) As SetupConGraficoRampa
    On Error GoTo err:
    
    
    Index = Channel.PointsNumber - 1
        
    ReadGraphPoints = eOK_Rampe
    Exit Function
err:
    ReadGraphPoints = eERROR_Rampe
End Function


' Return TRUE errore
Public Function ReadChannelSinglePointRampa(ByRef Channel As TagChannelRampa, ByVal indice As Integer, ByRef X As Double, ByRef Y As Double) As SetupConGraficoRampa
    Dim bReturn As SetupConGraficoRampa
    
    bReturn = eOK_Rampe

    If (indice >= Channel.PointsNumber) Then
        bReturn = eERROR_Rampe
    Else
        X = Channel.Points(indice).X
        Y = Channel.Points(indice).Y
    End If
    
    ReadChannelSinglePointRampa = bReturn
End Function





'-------------------------------------------------------------------------------------
'funzione che trasforma il comando al MLT da valore in volt (0-5V) a valore in % (AVC)
'
' AVC = Auxiliary Valve Command
'-------------------------------------------------------------------------------------
Private Function GetCommand(ByVal X0 As Double, ByVal X1 As Double, ByVal Y0 As Double, ByVal Y1 As Double, ByVal X As Double) As Double
    Dim m As Double
    Dim c As Double
    Dim Y As Double
    Dim maxY As Double
    Dim minY As Double
    
    m = (Y1 - Y0) / (X1 - X0)
    c = Y0 - m * X0
    
    Y = m * X + c
        
    If (Y1 > Y0) Then
        maxY = Y1
        minY = Y0
    Else
        maxY = Y0
        minY = Y1
    End If
    
       
    If (Y > maxY) Then
        Y = maxY
    ElseIf (Y < minY) Then
        Y = minY
    End If
    
    GetCommand = Y
    
End Function


' Questa funzione carica il grafico di un file
Public Sub Update_CURVE_XY_GRAPH(ByRef CWGraphImpostaCan As CWGraph, ByVal NomeFile As String)
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim X As Double
    Dim Y As Double
    Dim XMAX As Double
    CWGraphImpostaCan.ClearData
    Canal = FreeFile
    
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal
        
        Do While Not EOF(1)
        
            Line Input #Canal, stringa_tmp
            
            If (count >= 1) Then             ' salto la prima riga
            
                Dati = Split(stringa_tmp, ";")
                
                X = CDbl(Dati(0))
                Y = CDbl(Dati(1))
                
                CWGraphImpostaCan.ChartXvsY X, Y
                
                If (X > XMAX) Then
                    XMAX = X
                End If
                
            End If
            count = count + 1
        Loop
    End If

    
    CWGraphImpostaCan.Axes(1).Maximum = XMAX
    Close Canal
End Sub

' Questa funzione carica il grafico di un file

Public Function Update_CURVE_XY(ByRef Channel As TagChannelRampa, ByVal NomeFile As String) As Boolean
    On Error GoTo err:
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim X As Double
    Dim Y As Double
    Dim XMAX As Double
    
    Update_CURVE_XY = False
    
    Canal = FreeFile
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal
        
        Do While Not EOF(1)
            Line Input #Canal, stringa_tmp
            If (count >= 1) Then             ' salto la prima riga
                Dati = Split(stringa_tmp, ";")
                X = CDbl(Dati(0))
                Y = CDbl(Dati(1))

                If (X > XMAX) Then
                    XMAX = X
                End If
                
                If (WriteChannelSinglePointRampa(Channel, count - 1, X, Y) = eERROR_Rampe) Then
                    Update_CURVE_XY = True
                    Exit Function
                End If
            End If
            count = count + 1
        Loop
    Else
        Update_CURVE_XY = True
        Exit Function
    End If
    
    Channel.PointsNumber = count - 1
        
    Close Canal
    Update_CURVE_XY = False
    Exit Function
err:

    Update_CURVE_XY = True
    Close Canal
End Function


Public Sub Write_CURVE_XY_GRAPH(ByRef Channel As TagChannelRampa, ByRef CWGraphImpostaCan As CWGraph, ByVal NomeFile As String)
    On Error GoTo err:
    Dim count As Integer
    Dim Str As String
    Dim X As Double
    Dim Y As Double
    Dim Canal As Integer
    Canal = FreeFile
    'If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Output As Canal
            
        Str = "Tempo;Tensione"
        Do While (ReadChannelSinglePointRampa(Channel, count, X, Y) = eOK_Rampe)
            Str = Str + vbCrLf + CStr(Channel.Points(count).X) + ";" + CStr(Channel.Points(count).Y)
            count = count + 1
        Loop
        
        Print #Canal, Str
    'End If
err:
Close Canal
End Sub



' Return TRUE errore
Public Function Get_CURVE_XY_VALUE(ByRef Channel As TagChannelRampa, ByVal Tempo_sec As Double, ByRef CommandInMilliVolts As Double) As SetupConGraficoRampa
    Dim i As Integer
    Dim FirstY As Integer
    Dim SecondY As Integer
    Dim bFindx As Boolean
    
    Dim bReturn As SetupConGraficoRampa
    
    bReturn = eOK_Rampe

    i = 0
    Do While (i < UBound(Channel.Points) - 1)
    
        If (Tempo_sec > Channel.Points(i).X And Tempo_sec <= Channel.Points(i + 1).X) Then
        
            CommandInMilliVolts = GetCommand(Channel.Points(i).X, Channel.Points(i + 1).X, Channel.Points(i).Y, Channel.Points(i + 1).Y, Tempo_sec)
            bFindx = True
        End If
        
        i = i + 1
    Loop
    
    If (bFindx = False) Then
        bReturn = eERROR_Rampe
    End If
    

        
    Get_CURVE_XY_VALUE = bReturn
End Function


