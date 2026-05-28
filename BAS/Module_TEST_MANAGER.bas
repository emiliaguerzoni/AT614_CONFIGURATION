Attribute VB_Name = "Module_TEST_MANAGER"
Public Enum eGenericTestState
    eTEST_OK = 2                    ' TEST CONCLUSO CORRETTAMETNE
    eTEST_BUSY = 1                  ' TEST IN ESECUZIONE MA NON BLOCCANTE
    eTEST_STOP = 0                  ' TEST FERMO
    eTEST_BUSY_NON_BLOCCANTE = 3    ' TEST IN ESECUZIONE MA NON BLOCCANTE
    eTEST_ERROR = -1                ' TEST ERRORE
End Enum

Public Type TagInfosSingleTest
    CWGraph1 As CWGraph
    CWGraph2 As CWGraph
    RichTextBox1 As RichTextBox
    Label1 As Label
    CWButtonTest As CWButton
    codicePRodotto As Label
End Type


Public Type TagInfosTest
    InfosTest() As TagInfosSingleTest
End Type

' struttura datiusata per definire il singolo test
Public Type TagSingoloTest
    Name As String                  ' Nome del test
    ID As String                    ' Identificativo del test da svolgere
    Parameter() As String           ' Parametri del test
End Type

Public Type TagSezioniTest
    Test() As TagSingoloTest
End Type



' Questa apre il file della lista dei test da applicare per il modulo montato sulla sezione
' TRUE errore
Public Function ApriFileTest(FileName As String, ByRef Sezione As TagSezioniTest) As Boolean
    
    Dim fileId As Long
    Dim LineaDati As String                 ' linea dei dati
    Dim dati As Variant                     ' Dati
    Dim i As Integer
    Dim bFirstTime As Boolean
    Dim bReturn As Boolean
    Dim NumeroDiTest As Integer
    bReturn = True
    bFirstTime = True
    Dim LastTestIndex As Long
    
    On Error GoTo err:
    fileId = FreeFile
    
    
    NumeroDiTest = 0
    ReDim Sezione.Test(NumeroDiTest)
    
    
    Open FileName For Input As #fileId
  
    Do While (EOF(fileId) = False)
        Line Input #fileId, LineaDati
        dati = Split(LineaDati, ";")                    ' Prelevo i dati
        If (bFirstTime = True) Then
            bFirstTime = False
        Else
            Sezione.Test(NumeroDiTest).Name = dati(0)
            Sezione.Test(NumeroDiTest).ID = dati(1)
            i = 0
            ReDim Sezione.Test(NumeroDiTest).Parameter(UBound(dati) - 2)            ' Ridimensiono il vettore dei parametri
            
            ' controllo se i l'indice del test è un numero oppure no
            If (IsNumeric(dati(2)) = True) Then
                Sezione.Test(NumeroDiTest).Parameter(0) = dati(2)
                LastTestIndex = CLng(dati(2))                                   ' Memorizzo il valore del test valido
            Else
                ' Il primo campo non è un numero , allora potrebbe essere un valore incrementale
                
                If (dati(2) = "i++") Then
                    
                    Sezione.Test(NumeroDiTest).Parameter(0) = CStr(LastTestIndex)
                    LastTestIndex = LastTestIndex + 1
                    
                ElseIf (dati(2) = "++i") Then
                    LastTestIndex = LastTestIndex + 1
                    Sezione.Test(NumeroDiTest).Parameter(0) = CStr(LastTestIndex)
                ElseIf (dati(2) = "i") Then
                    Sezione.Test(NumeroDiTest).Parameter(0) = CStr(LastTestIndex)
                End If
                
            End If
            
            i = 1
            
            
            Do While (i < UBound(dati) - 1)
                Sezione.Test(NumeroDiTest).Parameter(i) = dati(2 + i)
                i = i + 1
            Loop
            NumeroDiTest = NumeroDiTest + 1
            ReDim Preserve Sezione.Test(NumeroDiTest)
        End If
    Loop
    
    bReturn = False
    
    ApriFileTest = bReturn
    Close #fileId
    Exit Function
err:
    ApriFileTest = bReturn
    Close #fileId
End Function
    
' Questa funzione ritorna la info sul tipo di test ( bloccante TRUE o no FALSE )
Public Function GetTestInfos(ByVal IndiceSezione As Integer, ByVal IndiceTest As Integer) As Integer
    On Error GoTo err:
    If (Sezione(IndiceSezione).CWButtonEnable.Value = False) Then
        GetTestInfos = -1
    Else
        GetTestInfos = CInt(TestSezione(IndiceSezione).Test(IndiceTest).Parameter(0))
    End If
    Exit Function
err:
    GetTestInfos = -1
End Function
    
' Questa funzione gestisce
Public Function TaskSingoloTest(ByVal IndiceSezione As Integer, ByVal IndiceTest As Integer, ByVal bState As Boolean, ByVal bStop As Boolean) As Integer
    On Error Resume Next
    Dim bReturn As Integer
    bReturn = eTEST_ERROR
    If (IndiceTest >= 0) Then
    
        ' "ACQUISIZIONE_PRE"  =  Acquisizione X-Y (PRE)
    
       
        
        TEST_POST_PROCESSING TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                    POST_PROCESSING(IndiceSezione, IndiceTest), bState, bStop, bReturn
    
        ' "ACQUISIZIONE_PRE"  =  Acquisizione X-Y (PRE)
        TEST_ACQUISIZIONE_POST TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                    ACQUISIZIONE_POST(IndiceSezione, IndiceTest), bState, bStop, bReturn
            
    
        SEQUENZA_PLC TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                    SEQUENZA_PLC_DATA(IndiceSezione, IndiceTest), bState, bStop, bReturn
    
        TEST_SCRITTURA_PARAMETRI TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               SCRITTURA_PARAMETRI(IndiceSezione, IndiceTest), bState, bStop, bReturn
    
        TEST_CALIBRAZIONE TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               CALIBRAZIONE(IndiceSezione, IndiceTest), bState, bStop, bReturn
                                               
        TEST_ASSEGNA_NODEID TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               ASSEGNA_NODEID(IndiceSezione, IndiceTest), bState, bStop, bReturn
                                               
                                               
        TEST_CICLICA_EVO TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               CICLICA_EVO(IndiceSezione, IndiceTest), bState, bStop, bReturn
        
                                               
                                               
        TEST_CICLICA TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               CICLICA(IndiceSezione, IndiceTest), bState, bStop, bReturn
        
        TEST_SCRITTURA_SERIAL_NUMBER TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               SERIAL_NUMBER(IndiceSezione, IndiceTest), bState, bStop, bReturn
        
        
        TEST_ACQUISIZIONE_CAN TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                               ACQUISIZIONE_CAN(IndiceSezione, IndiceTest), bState, bStop, bReturn
                                               
        TEST_LETTURA_CFG_HW TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                                LETTURA_CFG_HW(IndiceSezione, IndiceTest), bState, bStop, bReturn

        
        TEST_ACQUISIZIONE_PRE_EVO TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                    ACQUISIZIONE_PRE_EVO(IndiceSezione, IndiceTest), bState, bStop, bReturn


        TEST_RISPOSTE_GRADINO TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                                RISPOSTE_GRADINO(IndiceSezione, IndiceTest), bState, bStop, bReturn
                                                
        ' "ACQUISIZIONE_PRE"  =  Acquisizione X-Y (PRE)
        TEST_ACQUISIZIONE_PRE TestSezione(IndiceSezione).Test(IndiceTest), IndiceSezione, InfoSezione(IndiceSezione).InfosTest(IndiceTest), _
                                    ACQUISIZIONE_PRE(IndiceSezione, IndiceTest), bState, bStop, bReturn
                                                
    End If
    If (bReturn = -1 And IndiceSezione = 0 And IndiceTest = 0) Then
        'Debug.Print bReturn
    End If
    TaskSingoloTest = bReturn
End Function

    
