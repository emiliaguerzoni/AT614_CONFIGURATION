Attribute VB_Name = "Module_CURVE_LIMITE"
Public Enum eInfSup
    eInferiore                          ' Curva inferiore
    eSuperiore                          ' Curva superiore
    eInferiore_2 = 2                         ' Curva inferiore
    eSuperiore_2 = 3                         ' Curva superiore
End Enum

Public Enum eCurveLimiteError
    eERROR_CurveLimite
    eOK_CurveLimite
End Enum


' Struttura dati del singolo punto della curva limite
Type PuntoCurveLimiteTag
    X As Double                     ' valore asse x del punto della curva limite
    Y As Double                     ' valore asse y del punto della curva limite
    m As Double
    c As Double
    min As Double
    max As Double
    
End Type


' Struttura dati di una singola curva limite
Type CurveLimiteTag
    Punto() As PuntoCurveLimiteTag                   ' vettore dei punti della curva limite
    'Indice0 As Integer                               ' Indice del punto da utilizzare per calcolare la curva spezzata
    'Indice1 As Integer                               ' Indice del punto da utilizzare per calcolare la curva spezzata
End Type

' Struttura dati dei parametri del modulo
Type CurveLimitePlusVariablesTag
    Curva(2) As CurveLimiteTag                     ' Struttura dati della curva inferiore
End Type

' Struttura dati usata dall'utente per configurare opportunamente il modulo
Type CurveLimitePlusParametersTag
    Indice(2) As Integer                            ' Indice del plot associata alla curve inferiore della classe CWGraph. Inizializzata correttamente permete di gestire la visualizzazione della curva
End Type


' struttura dati del contesto del modulo. Questa struttura dati deve essere inizializzata nel form principale e deve essere sempre passata al modulo
Type CurveLimitePlusTag
    Variables As CurveLimitePlusVariablesTag            ' Struttura dati delle variabili gestite dal modulo
    Parameter As CurveLimitePlusParametersTag           ' Struttura dati dei parametri gestiti dal modulo
End Type


' Questa funzione inizializza la struttura dati dei parametri del modulo software.
' Param CurveLimitePlus - Contesto del modulo, da inizializzare all'interno del form principale
' Param Parameter - Struttura dati dei parametri del modulo
Public Function CurveLimitePlusInit(ByRef Context As CurveLimitePlusTag, ByRef par As CurveLimitePlusParametersTag) As eCurveLimiteError
    Dim eReturn As eCurveLimiteError
    eReturn = eOK_CurveLimite
    
    Context.Parameter = par
    
    CurveLimitePlusInit = eReturn
End Function


' Questa funzione carica da un file la struttura dei punti delle curve limite
' Param CurveLimitePlus - Contesto del modulo, da inizializzare all'interno del form principale
' Param Graph - classe del grafico da gestire
' Param FileName - Nome del file da cui accedere per caricare i valori delle curve limite
Public Function CurveLimitePlusLoad(ByRef Context As CurveLimitePlusTag, ByVal NomeFile As String, ByVal Curva As eInfSup) As eCurveLimiteError
    On Error GoTo err:
    Dim eReturn As eCurveLimiteError
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim Xinf As Double
    Dim Yinf As Double
    Dim Xsup As Double
    Dim Ysup As Double
    Dim Punto As Integer
    
    eReturn = eOK_CurveLimite
    
    Canal = FreeFile
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal
        Punto = 0
        ReDim Preserve Context.Variables.Curva(Curva).Punto(0)
        Do While Not EOF(Canal)
            Line Input #Canal, stringa_tmp
            If (count >= 1) Then             ' salto la prima riga
                Dati = Split(stringa_tmp, ";")
                X = CDbl(Dati(0))
                Y = CDbl(Dati(1))
                
                Context.Variables.Curva(Curva).Punto(UBound(Context.Variables.Curva(Curva).Punto)).X = X             ' mi salvo il punto x inferiore
                Context.Variables.Curva(Curva).Punto(UBound(Context.Variables.Curva(Curva).Punto)).Y = Y             ' mi salvo il punto Y inferiore
                
                ReDim Preserve Context.Variables.Curva(Curva).Punto(UBound(Context.Variables.Curva(Curva).Punto) + 1)     ' ridimensiono il vettore dei punti inferiori
               
                
    
                ' Ricavo il valore massimo
                If (Y > Ysup) Then
                    Ysup = Y
                End If
                
                If (X > Xsup) Then
                    Xsup = X
                End If
                                                       
            End If
            count = count + 1
        Loop
    End If
        
    Close Canal                                                                 ' Chiudo il file
    
    If (UBound(Context.Variables.Curva(Curva).Punto) > 0) Then
        ReDim Preserve Context.Variables.Curva(Curva).Punto(UBound(Context.Variables.Curva(Curva).Punto) - 1)     ' ridimensiono il vettore dei punti inferiori
    End If

    CurveLimitePlusLoad = eReturn
    Exit Function
err:
    CurveLimitePlusLoad = eERROR_CurveLimite
    Debug.Print err.Description
    Close Canal                                                                 ' Chiudo il file
End Function


' Questa funzione carica da un file la struttura dei punti delle curve limite
' Param CurveLimitePlus - Contesto del modulo, da inizializzare all'interno del form principale
' Param Graph - classe del grafico da gestire
' Param FileName - Nome del file da cui accedere per caricare i valori delle curve limite
Public Function CurveLimitePlusLoad_GRAPH(ByRef Graph As CWGraph, ByVal NomeFile As String, ByVal Curva As eInfSup) As eCurveLimiteError
    On Error GoTo err:
    Dim eReturn As eCurveLimiteError
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim Xinf As Double
    Dim Yinf As Double
    Dim Xsup As Double
    Dim bFirstTime As Boolean
    Dim Ysup As Double
    Dim Punto As Integer
    Dim X As Double
    Dim Y As Double
    
    eReturn = eOK_CurveLimite
    bFirstTime = True
    Canal = FreeFile
    
    
  
    
    
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal
        Punto = 0
        Graph.Plots(Curva + 1).ClearData
        
        Do While Not EOF(Canal)
            Line Input #Canal, stringa_tmp
            If (count >= 1) Then             ' salto la prima riga
                Dati = Split(stringa_tmp, ";")
                
                X = CDbl(Dati(0))
                Y = CDbl(Dati(1))
                
                If (bFirstTime = True) Then
                    bFirstTime = False
                    Xinf = X
                    Xsup = X
                    Yinf = Y
                    Ysup = Y
                End If
                
                Graph.Plots(Curva + 1).ChartXvsY X, Y        ' Disegno il punto
    
                ' Ricavo il valore massimo
                If (Y > Ysup) Then
                    Ysup = Y
                End If
                
                If (X > Xsup) Then
                    Xsup = X
                End If
                                       
                ' Ricavo il valore massimo
                If (Y < Yinf) Then
                    Yinf = Y
                End If
                
                If (X < Xinf) Then
                    Xinf = X
                End If
                Punto = Punto + 1                                                   ' Passo al prossimo punto
            End If
            count = count + 1
        Loop
    End If
        
    Close Canal                                                                 ' Chiudo il file
    
    Graph.Axes(2).Maximum = Ysup + 1
    Graph.Axes(2).Minimum = Yinf - 1
    Graph.Axes(1).AutoScale = True
    ' Se la curva è la superiore allora aggiorno anche la scala del grafico
    If (Curva = eSuperiore) Then
          
    End If
    
    CurveLimitePlusLoad_GRAPH = eReturn
    Exit Function
err:
    CurveLimitePlusLoad_GRAPH = eERROR_CurveLimite
    Debug.Print err.Description
    Close Canal                                                                 ' Chiudo il file
End Function

' Questa funzione carica da un file la struttura dei punti delle curve limite
' Param CurveLimitePlus - Contesto del modulo, da inizializzare all'interno del form principale
' Param Graph - classe del grafico da gestire
' Param FileName - Nome del file da cui accedere per caricare i valori delle curve limite
Public Function CurveLimitePlusVisualize(ByRef Context As CurveLimitePlusTag, ByVal NomeFile As String, ByVal Curva As eInfSup) As eCurveLimiteError
    On Error GoTo err:
    Dim eReturn As eCurveLimiteError
    Dim stringa_tmp As String
    Dim Canal As Long
    Dim count As Integer
    Dim Dati As Variant
    Dim Xinf As Double
    Dim Yinf As Double
    Dim Xsup As Double
    Dim Ysup As Double

    
    eReturn = eOK_CurveLimite
    
    Canal = FreeFile
    If (Dir(NomeFile, vbNormal) <> "") Then
        'leggo il valore da file altrimenti resta quello di default
        
        Open NomeFile For Input As Canal

        'Graph.Plots(Context.Parameter.indice(Curva)).ClearData
        Do While Not EOF(Canal)
            Line Input #Canal, stringa_tmp
            If (count >= 1) Then             ' salto la prima riga
                Dati = Split(stringa_tmp, ";")
                X = CDbl(Dati(0))
                Y = CDbl(Dati(1))
                

            
    
                ' Ricavo il valore massimo
                If (Y > Ysup) Then
                    Ysup = Y
                End If
                                                                                        ' Passo al prossimo punto
            End If
            count = count + 1
        Loop
    End If
        
    Close Canal                                                                 ' Chiudo il file
    

    
    CurveLimitePlusVisualize = eReturn
    Exit Function
err:
    CurveLimitePlusVisualize = eERROR_CurveLimite
    Close Canal                                                                 ' Chiudo il file
End Function

' Questa funzione controlla se il valore passato è all'interno delle curve limite
' Param Context - Puntatore al contesto
' Param X - Valore asse X
' Param Y - Valore asse Y

Public Function CurveLimitePlusCheck(ByRef Context As CurveLimitePlusTag, X As Double, Y As Double) As eCurveLimiteError
    Dim Punto As PuntoCurveLimiteTag
    Dim eReturn As eCurveLimiteError
    
    Dim bStatusSup As Boolean
    Dim bStatusInf As Boolean
On Error GoTo err:
    Punto.X = X
    Punto.Y = Y
    eReturn = eOK_CurveLimite
    If (InternalCurveLimitePlusCheck(Context, Punto, eInferiore, bStatusSup) = eERROR_CurveLimite) Then
        eReturn = eERROR_CurveLimite
        Distributore.Text3 = "INF"
    ElseIf (InternalCurveLimitePlusCheck(Context, Punto, eSuperiore, bStatusInf) = eERROR_CurveLimite) Then
        eReturn = eERROR_CurveLimite
        Distributore.Text3 = "SUB"
    Else
        If (bStatusSup = False And bStatusInf = False) Then
            CurveLimitePlusCheck = eOK_CurveLimite
        End If
    End If
    
    CurveLimitePlusCheck = eReturn
    Exit Function
err:
    Debug.Print err.Description
    CurveLimitePlusCheck = eERROR_CurveLimite
End Function


' QUesta funzione controlla se il punto Y,X è al di sopra della curva limite associata
' Param - Context puntatore al contesto del modulo
' Param - Punto Punto della curva da controllare
' Param - curva tipologia di curva ( inferiore o superiore )
' Param - bStatus - Stato di ritorno , TRUE il punto è ERROR , altrimento FALSE

Private Function InternalCurveLimitePlusCheck(ByRef Context As CurveLimitePlusTag, Punto As PuntoCurveLimiteTag, Curva As eInfSup, ByRef bStatus As Boolean) As eCurveLimiteError
    On Error GoTo err:
    Dim Indice0 As Integer
    Dim Indice1 As Integer
    Dim bFind As Boolean
    Dim bExit As Boolean
    'Dim bFind As Boolean
    Dim min As Double
    Dim max As Double
    Dim m As Double
    Dim c As Double
    Dim Y As Double
    
    Dim i As Integer
    bExit = False
    bFind = False
    i = 1
    Do While (bExit = False)
        If (Punto.X >= Context.Variables.Curva(Curva).Punto(i).X And Punto.X < Context.Variables.Curva(Curva).Punto(i + 1).X) Then      ' controllo se il punto ricercato
            bExit = True
            bFind = True
        ElseIf (i >= UBound(Context.Variables.Curva(Curva).Punto) - 1) Then
            bExit = True
        Else
            i = i + 1
        End If
    Loop
    
    If (bFind = True) Then                  ' ho trovato il punto all'interno della curva
        Indice0 = i                         ' mi salvo il punto attuale
        Indice1 = i + 1                     ' mi salvo il punto successivo
    Else
        Indice0 = i                         ' mi salvo l'indice ultimo
        Indice1 = i                         ' mi salvo l'indice ultimo
    End If
    InternalCurveLimitePlusCalcolaKoefficientiCurva Context.Variables.Curva(Curva).Punto(Indice0), Context.Variables.Curva(Curva).Punto(Indice1), m, c, min, max
    
    bStatus = False
    If (Punto.Y < min) Then
        If (Curva = eInferiore) Then
            bStatus = True
        End If
    ElseIf (Punto.Y > max) Then
        If (Curva = eSuperiore) Then
            bStatus = True
        End If
    Else
        Y = m * Punto.X + c
        If (Curva = eSuperiore) Then
            If (Punto.Y > Y) Then
                bStatus = True
            End If
        Else
            If (Punto.Y < Y) Then
                bStatus = True
            End If
        End If
        
    End If
    If (bStatus = False) Then
        InternalCurveLimitePlusCheck = eOK_CurveLimite
    Else
        InternalCurveLimitePlusCheck = eERROR_CurveLimite
    End If
    Exit Function
err:
    InternalCurveLimitePlusCheck = eERROR_CurveLimite
    Debug.Print err.Description
    
End Function


' Questa funzione calcola i coefficienti c , m della retta ed i limiti inferiori / superiori della spezzata
' param Punto0 - Punto 0 usato per il calcolo
' param Punto1 - Punto 1 usato per il calcolo
' param m koefficiente angolare della retta
' param c costante della retta
' param min valore minimo
' param max valore massimo
Private Sub InternalCurveLimitePlusCalcolaKoefficientiCurva(Punto0 As PuntoCurveLimiteTag, Punto1 As PuntoCurveLimiteTag, ByRef m As Double, ByRef c As Double, min As Double, max As Double)
    On Error Resume Next
    m = (Punto0.Y - Punto1.Y) / (Punto0.X - Punto1.X)
    c = Punto0.Y - m * Punto0.X
    
    If (Punto0.Y < Punto1.Y) Then
        min = Punto0.Y
        max = Punto1.Y
    Else
        min = Punto1.Y
        max = Punto0.Y
    End If
End Sub




Public Function CurveLimitePlusCheck_2(ByRef Context As CurveLimitePlusTag, X As Double, Y As Double, ByRef min As Double, ByRef max As Double) As eCurveLimiteError
    
    Dim Punto As PuntoCurveLimiteTag
    Dim eReturn As eCurveLimiteError

    Dim bStatusSup As Boolean
    Dim bStatusInf As Boolean

On Error GoTo err:
    
    Punto.X = X
    Punto.Y = Y
    eReturn = eOK_CurveLimite
    
    If (InternalCurveLimitePlusCheck2(Context, Punto, eInferiore, bStatusSup, min) = eERROR_CurveLimite) Then
        eReturn = eERROR_CurveLimite
        Distributore.Text3 = "INF"
    
    ElseIf (InternalCurveLimitePlusCheck2(Context, Punto, eSuperiore, bStatusInf, max) = eERROR_CurveLimite) Then
        eReturn = eERROR_CurveLimite
        Distributore.Text3 = "SUP"
    
    Else
        If (bStatusSup = False And bStatusInf = False) Then
            CurveLimitePlusCheck_2 = eOK_CurveLimite
        End If
    End If
    
    CurveLimitePlusCheck_2 = eReturn
    Exit Function
    
err:
    'Debug.Print err.Description
    CurveLimitePlusCheck_2 = eERROR_CurveLimite
    
End Function




' QUesta funzione controlla se il punto Y,X è al di sopra della curva limite associata
' Param - Context puntatore al contesto del modulo
' Param - Punto Punto della curva da controllare
' Param - curva tipologia di curva ( inferiore o superiore )
' Param - bStatus - Stato di ritorno , TRUE il punto è ERROR , altrimento FALSE

Private Function InternalCurveLimitePlusCheck2(ByRef Context As CurveLimitePlusTag, Punto As PuntoCurveLimiteTag, Curva As eInfSup, ByRef bStatus As Boolean, ByRef p As Double) As eCurveLimiteError
    
    On Error GoTo err:
    
    Dim Indice0 As Integer
    Dim Indice1 As Integer
    Dim bFind As Boolean
    Dim bExit As Boolean
    'Dim bFind As Boolean
    Dim m As Double
    Dim c As Double
    Dim Y As Double
    
    
    Dim min As Double
    Dim max As Double
    Dim i As Integer
    bExit = False
    bFind = False
    
    'i = 1
    i = 0
    Do While (bExit = False)
        If (Punto.X >= Context.Variables.Curva(Curva).Punto(i).X And Punto.X < Context.Variables.Curva(Curva).Punto(i + 1).X) Then      ' controllo se il punto ricercato
            bExit = True
            bFind = True
        ElseIf (i >= UBound(Context.Variables.Curva(Curva).Punto) - 1) Then
            bExit = True
        Else
            i = i + 1
        End If
    Loop
    
    If (bFind = True) Then                  ' ho trovato il punto all'interno della curva
        Indice0 = i                         ' mi salvo il punto attuale
        Indice1 = i + 1                     ' mi salvo il punto successivo
    Else
        Indice0 = i                        ' mi salvo l'indice ultimo
        Indice1 = i + 1                       ' mi salvo l'indice ultimo
     
        InternalCurveLimitePlusCheck2 = eOK_CurveLimite
        
        InternalCurveLimitePlusCalcolaKoefficientiCurva Context.Variables.Curva(Curva).Punto(Indice0), Context.Variables.Curva(Curva).Punto(Indice1), m, c, min, max
    
        p = m * Punto.X + c
        
        Exit Function
    End If
    
    InternalCurveLimitePlusCalcolaKoefficientiCurva Context.Variables.Curva(Curva).Punto(Indice0), Context.Variables.Curva(Curva).Punto(Indice1), m, c, min, max
    
    p = m * Punto.X + c

    
    bStatus = False
    If (Punto.Y < min) Then
        If (Curva = eInferiore) Then
            bStatus = True
        End If
    ElseIf (Punto.Y > max) Then
        If (Curva = eSuperiore) Then
            bStatus = True
        End If
    Else
        Y = CLng(m * Punto.X + c)
        If (Curva = eSuperiore) Then
            If (Punto.Y > Y) Then
                bStatus = True
            End If
        Else
            If (Punto.Y < Y) Then
                bStatus = True
            End If
        End If
        
    End If
    If (bStatus = False) Then
        InternalCurveLimitePlusCheck2 = eOK_CurveLimite
    Else
        InternalCurveLimitePlusCheck2 = eERROR_CurveLimite
    End If
    Exit Function
err:
    InternalCurveLimitePlusCheck2 = eERROR_CurveLimite
    'Debug.Print err.Description
    
End Function
