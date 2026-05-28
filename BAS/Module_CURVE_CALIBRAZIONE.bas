Attribute VB_Name = "Module_CURVE_CALIBRAZIONE"
Public Type TagPunto
    x As Double
    y As Double
End Type

' Struttura dati di scambio con il file di settaggio
Public Type TagCalibration
    Punto() As TagPunto             ' Vettore dei punti di calibrazione
    MaxPoint As Integer             ' Numero di punti trovati all'interno del file di calibrazione
    MaxYPoint As Double             ' valore del massimo punto trovato all'interno del file di calibrazione
    MinYPoint As Double             ' valore del minimo punto trovato all'interno del file di calibrazione
    UnitaMisura As String           ' unità di misura dell'asse Y
    FileName As String              ' Percorso completo del file di configurazione
End Type



' True errore
Public Function ApriCalibration(ByRef Context As TagCalibration) As Boolean
    Dim fileId As Long
    Dim LineaDati As String                 ' linea dei dati
    Dim Dati As Variant                     ' Dati
    Dim i As Integer
    Dim bFirstTime As Boolean
    Dim bReturn As Boolean
    Dim bFirstPoint As Boolean
    bReturn = True
    
    bFirstTime = True
    
    ReDim Context.Punto(0)
    Context.MaxPoint = 0
    
    On Error GoTo err:
    fileId = FreeFile

    Open Context.FileName For Input As #fileId
  
    Do While (EOF(fileId) = False)
        Line Input #fileId, LineaDati
        
        
        If (bFirstTime = True) Then
            bFirstTime = False
            Dati = Split(LineaDati, ";")                    ' Prelevo i dati
            Context.UnitaMisura = Dati(1)
        Else
            Dati = Split(LineaDati, ";")                    ' Prelevo i dati
          If (UBound(Dati) = 1) Then                      ' Controllo quanti campi ci sono attivi
              Context.Punto(Context.MaxPoint).x = Dati(0)
              Context.Punto(Context.MaxPoint).y = Dati(1)
              Context.MaxPoint = Context.MaxPoint + 1
              
              If (bFirstPoint = True) Then
                bFirstPoint = False
                Context.MaxYPoint = Dati(1)
                Context.MinYPoint = Dati(1)
              End If
              
              If (Context.MaxYPoint < Dati(1)) Then
                Context.MaxYPoint = Dati(1)
              End If
              
              If (Context.MinYPoint > Dati(1)) Then
                Context.MinYPoint = Dati(1)
              End If
              
              ReDim Preserve Context.Punto(Context.MaxPoint)
          End If
        
        End If
    Loop
    bReturn = False
    
    ApriCalibration = bReturn
    Close #fileId
    Exit Function
err:
    ApriCalibration = bReturn
    Debug.Print err.Description
    Close #fileId
End Function


Public Function SaveCalibration(ByRef Context As TagCalibration) As Boolean
    Dim fileId As Long
    Dim LineaDati As String                 ' linea dei dati
    Dim Dati As Variant                     ' Dati
    Dim i As Integer
    Dim bReturn As Boolean
    
    bReturn = True
    
    On Error GoTo err:
    fileId = FreeFile
    

    Open Context.FileName For Output As #fileId
    
    LineaDati = "PUNTO X;PUNTO Y"
    Print #fileId, LineaDati
    
    Do While (i < Context.MaxPoint)
    
        LineaDati = CStr(Context.Punto(i).x) + ";" + CStr(Context.Punto(i).y)
        Print #fileId, LineaDati
        
        i = i + 1
    Loop
    SaveCalibration = False
    Close #fileId
    Exit Function
err:
    SaveCalibration = True
    Close #fileId
End Function




' QUesta funzione controlla se il punto Y,X è al di sopra della curva limite associata
' Param - Context puntatore al contesto del modulo
' Param - Punto Punto della curva da controllare
' Param - curva tipologia di curva ( inferiore o superiore )
' Param - bStatus - Stato di ritorno , TRUE il punto è ERROR , altrimento FALSE

Public Function GetCalibrationPoint(ByRef Context As TagCalibration, ByVal ValueIn As Double, ByRef ValueOut As Double) As Boolean
    Dim Indice0 As Integer
    Dim Indice1 As Integer
    Dim bFind As Boolean
    Dim bExit As Boolean
    'Dim bFind As Boolean
    Dim min As Double
    Dim max As Double
    Dim m As Double
    Dim c As Double
    Dim y As Double
    
    Dim i As Integer
    bExit = False
    bFind = False
    i = 0
    If (Context.MaxPoint < 2) Then
        GetCalibrationPoint = True
    Else
    
        GetCalibrationPoint = False
        
        Do While (i < Context.MaxPoint And bExit = False)
            If (ValueIn >= Context.Punto(i).x And ValueIn <= Context.Punto(i + 1).x) Then       ' controllo se il punto ricercato
                bExit = True
                bFind = True
            Else
                i = i + 1
            End If
        Loop

    
        If (bFind = True) Then                  ' ho trovato il punto all'interno della curva
            Indice0 = i                         ' mi salvo il punto attuale
            Indice1 = i + 1                     ' mi salvo il punto successivo
            CalcolaKoefficientiCurva Context.Punto(Indice0), Context.Punto(Indice1), m, c, min, max
            ValueOut = m * ValueIn + c
        Else
            If (ValueIn < Context.Punto(0).x) Then
                Indice0 = 0                         ' mi salvo l'indice ultimo
                Indice1 = 1                         ' mi salvo l'indice ultimo
            
            ElseIf (ValueIn > Context.Punto(Context.MaxPoint - 1).x) Then
                Indice0 = Context.MaxPoint - 2                         ' mi salvo l'indice ultimo
                Indice1 = Context.MaxPoint - 1                         ' mi salvo l'indice ultimo
            End If
            
            CalcolaKoefficientiCurva Context.Punto(Indice0), Context.Punto(Indice1), m, c, min, max
            ValueOut = m * ValueIn + c
        End If
        
        
        
        
        
        
        
    End If
End Function

' Questa funzione calcola i coefficienti c , m della retta ed i limiti inferiori / superiori della spezzata
' param Punto0 - Punto 0 usato per il calcolo
' param Punto1 - Punto 1 usato per il calcolo
' param m koefficiente angolare della retta
' param c costante della retta
' param min valore minimo
' param max valore massimo
Private Sub CalcolaKoefficientiCurva(Punto0 As TagPunto, Punto1 As TagPunto, ByRef m As Double, ByRef c As Double, min As Double, max As Double)
    On Error Resume Next
    If (Punto0.x = Punto1.x) Then
        m = 0
        c = Punto0.x
    Else
        m = (Punto0.y - Punto1.y) / (Punto0.x - Punto1.x)
        c = Punto0.y - m * Punto0.x
    End If
    
    If (Punto0.y < Punto1.y) Then
        min = Punto0.y
        max = Punto1.y
    Else
        min = Punto1.y
        max = Punto0.y
    End If
End Sub



