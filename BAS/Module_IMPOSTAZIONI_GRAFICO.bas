Attribute VB_Name = "Module_IMPOSTAZIONI_GRAFICO"

'**************************************************************
'       M O D U L O :    Module_IMPOSTAZIONI_GRAFICO.bas
'**************************************************************

' Struttura dati di scambio con il file di settaggio
Public Type TagImpostazioniGrafico
    NomeVariabile As String         ' "NOME"
    IdScheda As Integer             ' "ID SCHEDA"
    TipoVariabile As String         ' "TIPO VARIABILE"
    AddrVariabile As Integer        ' "INDIRIZZO VARIABILE"
    MinVariabile As Integer         ' "MIN"
    MaxVariabile As Integer         ' "MAX"
    MajorTick As Integer            ' "MAJOR TICK"
    MinorTick As Integer            ' "MINOR TICK"
    LineColor As Long               ' "LINE COLOR"
End Type


Public Enum eErrorImpostaCWGraph
    eERROR_ImpostaCWGraph
    eOK_ImpostaCWGraph
End Enum


 
' Questa funzione carica da un file la configurazione del grafico "CWGraph"
' Param CurveLimitePlus - Contesto del modulo, da inizializzare all'interno del form principale
' Param Graph - classe del grafico da gestire
' Param FileName - Nome del file da cui accedere per caricare i valori delle curve limite
Public Function ImpostaCWGraphLoad(ByRef CWGrafico() As TagImpostazioniGrafico, ByVal NomeFile As String)

    On Error GoTo err:

    Dim eReturn As eErrorImpostaCWGraph

    Dim Canal As Long
    Dim count As Integer
    Dim LineaDati As String                 ' linea dei dati
    Dim Dati As Variant
    
'    NumeroDiVariabili = 0
    ReDim CWGrafico(0)  'inizializzazione (nessuna riga presente)
    count_righe_tot = 0
    count_righe_var = 0
    
    eReturn = eOK_ImpostaCWGraph

    'leggi file
    Canal = FreeFile
    If (Dir(NomeFile, vbNormal) <> "") Then
    
        Open NomeFile For Input As Canal
        
        Do While Not EOF(Canal)
            Line Input #Canal, LineaDati
            Dati = Split(LineaDati, ";")

            If (UBound(Dati) = 5) Then                  ' Controllo quanti campi ci sono attivi

                If (count_righe_tot >= 1) Then          ' salto la prima riga
                    CWGrafico(count_righe_var).NomeVariabile = Dati(0)
                    'CWGrafico(count_righe_var).IdScheda = Dati(1)
                    'CWGrafico(count_righe_var).TipoVariabile = Dati(2)
                    'CWGrafico(count_righe_var).AddrVariabile = Dati(3)
                    CWGrafico(count_righe_var).MinVariabile = Dati(1)
                    CWGrafico(count_righe_var).MaxVariabile = Dati(2)
                    CWGrafico(count_righe_var).MajorTick = Dati(3)
                    CWGrafico(count_righe_var).MinorTick = Dati(4)
                    CWGrafico(count_righe_var).LineColor = Dati(5)
                    
                    count_righe_var = count_righe_var + 1
                
                    ReDim Preserve CWGrafico(count_righe_var)
                End If
                
            End If '(UBound(Dati) = 1)
            count_righe_tot = count_righe_tot + 1
        
        Loop '(Not EOF(Canal))
    End If '(NomeFile <> "")

    Close Canal
    
    ImpostaCWGraphLoad = eReturn
    
    Exit Function

err:
    ImpostaCWGraphLoad = eERROR_ImpostaCWGraph
    Debug.Print err.Description

    Close Canal
    Exit Function


End Function 'END ImpostaCWGraphLoad()
  
  

