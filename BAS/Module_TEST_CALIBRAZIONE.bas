Attribute VB_Name = "Module_TEST_CALIBRAZIONE"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 500
Private Const ACCENSIONE_TIMEOUT_MS = 1000
'eTEST_OK = 2
'eTEST_BUSY = 1
'eTEST_STOP = 0
'eTEST_ERROR = -1



' Stato macchina del modulo
Private Enum eMachineState
    eInit
    eSpegniModuli
    eSpegniModuliAttendi
    eAccendiModulo
    eCapture
    eAttesaCapture
    eSave
    eRun
    eEnd
    eError
End Enum


Private Enum eTestParameter
    eNomeFile = 1
    eSincro = 2
    Corsa_Retract_mm = 3                            ' Colonna del file di test della corsa massima in retract
    Corsa_Extend_mm = 4                             ' Colonna del file di test della corsa massima in extend
    Corsa_Tolleranza_mm = 5                         ' Tolleranza accettabile dal valore nominale della corsa per scatenare la procedura di calibrazione del punto
End Enum

' Enumeratore dei punti di calibrazione del sensore CE16
Public Enum CalibEvent
    CalibEvent_Retract = 0                          ' Punto di retract ( -x mm )
    CalibEvent_Extend = 1                           ' Punto in extend ( + z mm )
    CalibEvent_Max
End Enum

' Struttura dati usata per gestire la calibrazione del modulo CE16
Public Type TagTEST_CALIBRAZIONE_CE16
    eEvent As Boolean                               ' Flag di attivazione di inizio dell'evento di calibrazione del punto
    bDifuEvent As Boolean                           ' Difu flag dell'evento di attivazione del punto di calibrazione
    TimerCampionamento As Long                      ' Timer usato per gestire il campionamento delle letture del sensore grezzo
    CorsaTeorica As Double                            ' Corsa Teorica letta del trasduttore
    Tolleranza As Double                              ' Tolleranza della corsa teorica
    LettureCE16() As Long                           ' Letture del sensore grezzo
    LetturaCE16Minima As Long                       ' Valore minima della lettura del CE16 ( da scartare )
    LetturaCE16Massima As Long                      ' Valore massimo della lettura del CE16 ( da scartare )
    XpointCE16 As Long                              ' Valore del punto di calibrazione asse X
    bEnable As Boolean                              ' Flag di segnalazione dell'avvenuta calibrazione del punto ( sovrascrivere il file di calibrazione
End Type


Public Type TagTEST_CALIBRAZIONE
    bOldOn As Boolean               ' difu del comando di start del modulo
    Return As Integer                ' Stato del modulo
    State As Integer
    OldState As Integer
    NodeCaptured As Integer
    Timer As Long
    Timeout As Long
    
    Calib(CalibEvent_Max) As TagTEST_CALIBRAZIONE_CE16      ' Struttura dati usata per gestire la calibrazione del CE16
End Type

Private ContextA As TagTEST_CALIBRAZIONE

Public Sub INFO_CALIBRAZIONE(Infos As TagInfosSingleTest, Optional bClear As Boolean = False)
    If (Infos.CWGraph1.Visible <> False) Then
        Infos.CWGraph1.Visible = False
    End If
    
    If (Infos.RichTextBox1.Visible = False) Then
        Infos.RichTextBox1.Visible = True
    End If
    
    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
    End If
    
    If (bClear = True) Then
        
        If (Infos.RichTextBox1.Text <> "") Then
            Infos.RichTextBox1.Text = ""
        End If
        Infos.CWGraph1.ClearData
        Infos.CWGraph2.ClearData
    End If

End Sub


Private Sub TEST_CALIBRAZIONE_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_CALIBRAZIONE)
    Context1.State = eEnd
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    
    Infos.RichTextBox1.Text = "CALIBRAZIONE"
End Sub

Public Sub TEST_CALIBRAZIONE(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_CALIBRAZIONE, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
    'Set ContextA = Context1
    If (Test.ID <> "CALIBRAZIONE") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
    
        INFO_CALIBRAZIONE Infos
        
        If (bStop = True And bOn = True) Then
            TEST_CALIBRAZIONE_INIT_STRUCT Infos, Context1
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then          ' fronte di salita del test ( AVVIO )
            Init Test, Infos, IndiceSezione, Context1                                         ' init delle variabili del modulo
        End If
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        
        Task Test, Infos, IndiceSezione, Context1                                       ' Task del modulo
        
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        
        bReturn = Context1.Return
    End If
    
End Sub

' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    If (Test.Parameter(eSincro) = "") Then
        Infos.RichTextBox1.Text = "CALIBRAZIONE BLOCCANTE" + vbLf
        ContextA.Return = eTEST_BUSY
    Else
        Infos.RichTextBox1.Text = "CALIBRAZIONE NON BLOCCANTE" + vbLf
        ContextA.Return = eTEST_BUSY_NON_BLOCCANTE
    End If
        
    ContextA.State = eSpegniModuli
    ContextA.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    
    InitCalibrationPoints Test, Infos, IndiceSezione, ContextA
    
    Prepare_Textbox_Parametri Test, Infos
    
    Infos.RichTextBox1.BackColor = vbWhite
End Sub

' Funzione di inizializzazione della procedura di inizializzazione dei punti di calibrazione
Private Sub InitCalibrationPoints(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CALIBRAZIONE)
    Dim i As Integer
    
    ' Inizializzo la struttura dati di ogni punto di calibrazione
    Do While (i < CalibEvent_Max)
        ReDim Context.Calib(i).LettureCE16(0)          ' Pulisco il vettore delle lettura del CE16
        
        
        If (Test.Parameter(Corsa_Retract_mm + i) = "" Or Test.Parameter(Corsa_Tolleranza_mm) = "") Then
            Context.Calib(i).bEnable = False
        Else
            Context.Calib(i).bEnable = True
        End If
        
        
        Context.Calib(i).bDifuEvent = False
        Context.Calib(i).eEvent = False
        
        If (Context.Calib(i).bEnable = True) Then
            Context.Calib(i).CorsaTeorica = CDbl(Test.Parameter(Corsa_Retract_mm + i))
            Context.Calib(i).Tolleranza = CDbl(Test.Parameter(Corsa_Tolleranza_mm))

            Infos.RichTextBox1.Text = "INIT PT CALIB CE16." + CStr(i) + " - CORSA=" + CStr(Context.Calib(i).CorsaTeorica) + "; TOLLE= " + CStr(Context.Calib(i).Tolleranza) + vbLf + Infos.RichTextBox1.Text
        Else
            Infos.RichTextBox1.Text = "INIT PT CALIB CE16." + CStr(i) + " DISABLE" + vbLf + Infos.RichTextBox1.Text
        End If
        i = i + 1
    Loop
    
    
End Sub

' Questa funzione gestisce la calibrazione dei punti del CE16
Private Sub TaskCalibrationPoints(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CALIBRAZIONE)
    
    Dim i As Integer
    Dim j As Integer
    Dim Z As Integer
    
    Do While (i < CalibEvent_Max)
        If (Context.Calib(i).bEnable = True And CE16(IndiceSezione).Position <> "") Then
            If (CE16(IndiceSezione).Position > Context.Calib(i).CorsaTeorica - Context.Calib(i).Tolleranza And _
                 CE16(IndiceSezione).Position < Context.Calib(i).CorsaTeorica + Context.Calib(i).Tolleranza) Then
                
                Context.Calib(i).eEvent = True          ' Sono all'interno della finestra di attivazione della calibrazione del punto
            Else
                Context.Calib(i).eEvent = False         ' Sono uscito dalla finestra di attivazione della calibrazione del punto
                
            End If
            
            
            
            If (Context.Calib(i).eEvent = True And Context.Calib(i).bDifuEvent = False) Then
                Infos.RichTextBox1.Text = "EVENT FIRE PT CALIB CE16." + CStr(i) + ";CE16 POSITION=" + CStr(CE16(IndiceSezione).Position) + vbLf + Infos.RichTextBox1.Text
            End If
                
            
                
            ' Evento di uscita dalla zona di calibrazione
            If (Context.Calib(i).bDifuEvent = True And Context.Calib(i).eEvent = False) Then
                
                Infos.RichTextBox1.Text = "EVENT EXIT PT CALIB CE16." + CStr(i) + vbLf + Infos.RichTextBox1.Text
                
                If (UBound(Context.Calib(i).LettureCE16) >= 3) Then              ' Per considerare valido il punto devo avere almeno 4 letture
                    
                    Infos.RichTextBox1.Text = "EVENT VALID PT CALIB CE16." + CStr(i) + vbLf + Infos.RichTextBox1.Text
                    
                    Context.Calib(i).XpointCE16 = Context.Calib(i).LettureCE16(0)
                    Context.Calib(i).LetturaCE16Minima = Context.Calib(i).LettureCE16(0)        ' Prendo il primo valore come riferimento per il valore minimo
                    Context.Calib(i).LetturaCE16Massima = Context.Calib(i).LettureCE16(0)       ' Prendo il primo valore come riferimento per il valore massimo
                    Infos.RichTextBox1.Text = "PTS" + CStr(0) + "=" + CStr(Context.Calib(i).LettureCE16(j)) + vbLf + Infos.RichTextBox1.Text
                    j = 1
                    Do While (j < UBound(Context.Calib(i).LettureCE16))
                        
                        If (Context.Calib(i).LettureCE16(j) < Context.Calib(i).LetturaCE16Minima) Then
                            Context.Calib(i).LetturaCE16Minima = Context.Calib(i).LettureCE16(j)
                        End If
                        
                        If (Context.Calib(i).LettureCE16(j) > Context.Calib(i).LetturaCE16Massima) Then
                            Context.Calib(i).LetturaCE16Massima = Context.Calib(i).LettureCE16(j)
                        End If
                        Infos.RichTextBox1.Text = "PTS" + CStr(j) + "=" + CStr(Context.Calib(i).LettureCE16(j)) + vbLf + Infos.RichTextBox1.Text
                        Context.Calib(i).XpointCE16 = Context.Calib(i).XpointCE16 + Context.Calib(i).LettureCE16(j)
                        
                        j = j + 1
                    Loop
                    
                    
                    ' Eseguo la media dei punto scartando il massimo ed il minimo
                    Context.Calib(i).XpointCE16 = Context.Calib(i).XpointCE16 - Context.Calib(i).LetturaCE16Massima - Context.Calib(i).LetturaCE16Minima
                    Context.Calib(i).XpointCE16 = Context.Calib(i).XpointCE16 / (j - 2)
                    Infos.RichTextBox1.Text = "ACQ=" + CStr(j - 2) + vbLf + Infos.RichTextBox1.Text
                    Infos.RichTextBox1.Text = "MEDIA=" + CStr(Context.Calib(i).XpointCE16) + vbLf + Infos.RichTextBox1.Text
                    
                    If (i = CalibEvent_Extend) Then
                        
                        CalibCE16(IndiceSezione).Punto(0).X = Context.Calib(i).XpointCE16
                        CalibCE16(IndiceSezione).Punto(0).Y = Context.Calib(i).CorsaTeorica
                    
                    ElseIf (i = CalibEvent_Retract) Then
                        Infos.RichTextBox1.Text = "PTS RETRACT " + CStr(UBound(CalibCE16(IndiceSezione).Punto) - 1) + vbLf + Infos.RichTextBox1.Text
                        CalibCE16(IndiceSezione).Punto(UBound(CalibCE16(IndiceSezione).Punto) - 1).X = Context.Calib(i).XpointCE16
                        CalibCE16(IndiceSezione).Punto(UBound(CalibCE16(IndiceSezione).Punto) - 1).Y = Context.Calib(i).CorsaTeorica
                                    
                    End If
                    
                    'Infos.RichTextBox1.Text = "EVENT SAVE PT CALIB CE16." + CStr(i) + vbLf + Infos.RichTextBox1.Text
                    Z = 0
                    Do While (Z <= UBound(CalibCE16(IndiceSezione).Punto))
                        Infos.RichTextBox1.Text = "X=" + CStr(CalibCE16(IndiceSezione).Punto(Z).X) + ";" + _
                                                  "Y=" + CStr(CalibCE16(IndiceSezione).Punto(Z).Y) + ";" + vbLf + Infos.RichTextBox1.Text
                        
                        Z = Z + 1
                    Loop
                    'Infos.RichTextBox1.Text = "EVENT SAVE PT CALIB CE16." + CStr(i) + vbLf + Infos.RichTextBox1.Text
                    
                    
                    
                    If (SaveCalibration(CalibCE16(IndiceSezione)) = True) Then
                        Infos.RichTextBox1.Text = "EVENT ERROR !!!!!!!!!!!!!!!! " + vbLf + Infos.RichTextBox1.Text
                    Else
                        Infos.RichTextBox1.Text = "EVENT SAVED !!!!!!!!!!!!!!!! " + vbLf + Infos.RichTextBox1.Text
                    End If
                    
                Else
                    Infos.RichTextBox1.Text = "EVENT NOT VALID PT CALIB CE16." + CStr(i) + vbLf + Infos.RichTextBox1.Text
                End If
            
            End If
                
                
            If (Context.Calib(i).eEvent = True) Then
                
                ' Controllo se il timer legato al campionamento è scaduto
                If (Context.Calib(i).TimerCampionamento < GetTickCount) Then
                    'Infos.RichTextBox1.Text = "REDIM " + CStr(UBound(Context.Calib(i).LettureCE16)) + vbLf + Infos.RichTextBox1.Text
                    
                    Context.Calib(i).LettureCE16(UBound(Context.Calib(i).LettureCE16)) = CLng(CE16(IndiceSezione).PositionGrezza)     ' Mi memorizzo la posizione
                    ReDim Preserve Context.Calib(i).LettureCE16(UBound(Context.Calib(i).LettureCE16) + 1)                     ' Aumento la dimensione del vettore
                    Context.Calib(i).TimerCampionamento = GetTickCount + 100                                            ' Aggiorno il timer
                    
                End If
            Else
                Context.Calib(i).TimerCampionamento = GetTickCount + 100
                ReDim Preserve Context.Calib(i).LettureCE16(0)               ' Aumento la dimensione del vettore
            End If

        End If
        
        
        ' Memorizzo la difu dell'evento
        Context.Calib(i).bDifuEvent = Context.Calib(i).eEvent
            
        i = i + 1
    Loop
    
End Sub



' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
On Error GoTo err:
    Select Case ContextA.State
    Case eInit: Init Test, Infos, IndiceSezione, ContextA
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, ContextA
    Case eSpegniModuliAttendi: SpegniModuliAttendi Test, Infos, IndiceSezione, ContextA
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, ContextA
    Case eCapture: Capture Test, Infos, IndiceSezione, ContextA
    Case eRun: Run Test, Infos, IndiceSezione, ContextA
    Case eAttesaCapture: AttesaCapture Test, Infos, IndiceSezione, ContextA
    Case eEnd: ContextA.Return = eTEST_OK
    Case eError: ErrorTest Test, Infos, IndiceSezione, ContextA
    Case Else: ErrorTest Test, Infos, IndiceSezione, ContextA
    End Select
    
    
    TaskCalibrationPoints Test, Infos, IndiceSezione, ContextA
    
    
    ContextA.OldState = ContextA.State
    Exit Sub
err:
    ContextA.Return = eTEST_ERROR
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + err.Description + vbLf
End Sub

Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CALIBRAZIONE)
    Infos.RichTextBox1.BackColor = vbRed
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = Infos.RichTextBox1.Text + "ERRORE SUL TEST" + vbLf
    End If
    Context.Return = eTEST_ERROR

End Sub

Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    
    
    If (Test.Parameter(eSincro) = "") Then
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "SPENGO TUTTI I MODULI" + vbLf
        TurnOffAllModule
    Else
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "SPENGO IL MODULO " + CStr(IndiceSezione) + vbLf
        TurnOffModule IndiceSezione
    End If
    
    ContextA.State = eSpegniModuliAttendi
    
End Sub

Private Sub SpegniModuliAttendi(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    If (ContextA.Timer < GetTickCount) Then
        ContextA.State = eAccendiModulo
        ContextA.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    TurnOnModule IndiceSezione
    
    If (ContextA.Timer < GetTickCount) Then
        ContextA.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    Dim str As String
    If (Test.Parameter(eSincro) = "") Then
        
        Infos.Label1.Caption = ""
        NodeCaptured(IndiceSezione) = CatturaMltPlus(, str)
        Infos.Label1.Caption = str
        
    Else
        
        NodeCaptured(IndiceSezione) = IndiceSezione + 1
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "CATTURO MODULO " + CStr(NodeCaptured(IndiceSezione)) + vbLf
        
        Infos.Label1.Caption = ""
        CatturaMltPlus NodeCaptured(IndiceSezione), str
        Infos.Label1.Caption = str
        
    End If
    
    ContextA.State = eAttesaCapture
    ContextA.Timer = GetTickCount + 500
    
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "A U T O C A L I B R A Z I O N E" & " - " & "ID " & CStr(NodeCaptured(IndiceSezione)) + vbLf
    

    
End Sub

Private Sub AttesaCapture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    If (ContextA.Timer < GetTickCount) Then
        ContextA.Timeout = GetTickCount + 15000
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "SEND MESSAGE " + vbLf
        SendAutocalibrationMessage NodeCaptured(IndiceSezione)
        
        ContextA.State = eRun
    End If
    
    
    
End Sub
Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, ContextA As TagTEST_CALIBRAZIONE)
    Dim AutoCalib As Integer
    AutoCalib = Autocalibrazione(NodeCaptured(IndiceSezione), Infos.RichTextBox1, ContextA.Timeout)
    Select Case AutoCalib
    Case 1:
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        ContextA.State = eError
    Case 0:
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        ContextA.State = eEnd
    End Select
    
    
End Sub




'------------------
'Autocalibrazione()
'------------------
' return TRUE as ERROR
Public Function Autocalibrazione(ByVal NodeId As Integer, ByRef RichTextBox1 As RichTextBox, Timeout As Long) As Integer

'NOTA: passo alla funzione l'indice del canale selezionato al momento
Dim i As Integer
Dim bExit As Boolean
Dim bOk As Boolean
Dim dummy As Integer
Dim TempString As String

'label operazione
bExit = False
Dim cmpString As String
cmpString = "T18EF228" + CStr(NodeId) + "8" + "78"
Autocalibrazione = 0


                    
        
    'mando messaggio AUTOCALIBRAZIONE
    '--------------------------------
    If (Timeout < GetTickCount) Then
        bExit = True
    Else
       
            'attendo messaggio di risposta
        Do While (GetFrameFromRxBuffer(Calib1 + NodeId - 1, TempString) = 0)
            
            'risposta a "AUTOCALIBRAZIONE" ("T18EF228"|NodeId|DLC=8|"78")
            If (InStr(1, TempString, cmpString) > 0) Then
                
                'If (Mid$(TempString, 13, 2) = "02" ) then
                If (Mid$(TempString, 13, 2) = "02" And Left$(TempString, 8) = "T18EF228") And (Mid$(TempString, 11, 2) = "78") Then
                    bExit = True
                    bOk = True
                    'spia "CALIBRAZIONE SPOLA OK" accesa
                    Autocalibrazione = 0
                    RichTextBox1.Text = RichTextBox1.Text + "CALIBRAZIONE OK" + vbLf
                    'campo dati
                End If
                    
            'CALIBRAZIONE FALLITA
            ElseIf ((Left$(TempString, 8) = "T18EF228") And (Mid$(TempString, 11, 2) = "71")) Then
                bExit = True
                RichTextBox1.Text = RichTextBox1.Text + "STATO = " + Mid$(TempString, 13, 2) + "ERROR" + vbLf
                
                
                
                'MESSAGGIO OPERATORE "ERRORE IN CALIBRAZIONE"
                '"riprova"
                If (MsgBox("MESSAGGIO DAL MODULO: ERRORE IN CALIBRAZIONE", vbRetryCancel, "ERRORE") = vbRetry) Then
                    bExit = False
                '"annulla"
                Else
                    Autocalibrazione = 0
                    Exit Function
                End If
                    
            End If 'END 'risposta a "AUTOCALIBRAZIONE"
    
        Loop 'attendo messaggio di risposta

    End If

    If (bExit = False) Then
        Autocalibrazione = 2
        'MsgBox "TIMEOUT CALIBRAZIONE", vbCritical, "ERRORE"
    Else
        If (bOk = True) Then
            Autocalibrazione = 0
        Else
            Autocalibrazione = 1
        End If
    End If

End Function 'END Autocalibrazione()


'----------------------------
'SendAutocalibrationMessage()
'----------------------------
Private Sub SendAutocalibrationMessage(ByVal NodeId As Integer)
    
    Dim campo_dati_can As String
    Dim msg_can2send As String
    Dim i As Integer
    Dim dummy As Integer
    Dim Text(8) As String
    
    Dim NodeIdHex As String
    
    NodeIdHex = CStr(NodeId)
    
    'valorizzo "campo_dati_can"
    Text(0) = "78"
    Text(1) = "FF"
    Text(2) = "FF"
    Text(3) = "FF"
    Text(4) = "FF"
    Text(5) = "FF"
    Text(6) = "FF"
    Text(7) = "FF"
    
    campo_dati_can = ""
    For i = 0 To 7
        campo_dati_can = campo_dati_can & Text(i)
    Next i
    
    'msg di "AUTOCALIBRAZIONE"
    msg_can2send = "T" & "18EF8" & NodeIdHex & "22" & "8" & campo_dati_can & Chr(13)
    Output msg_can2send
    dummy = DoEvents()



End Sub 'END SendAutocalibrationMessage()

Private Function ConvertHs1(Message As String) As Integer
    Dim TempString As String
    
    TempString = Mid$(Message, 15, 2) + Mid$(Message, 13, 2)
    
    ConvertHs1 = CInt("&H" + TempString)
End Function



Private Function ConvertHs2(Message As String) As Integer
    Dim TempString As String
    
    TempString = Mid$(Message, 19, 2) + Mid$(Message, 17, 2)
    
    ConvertHs2 = CInt("&H" + TempString)
End Function



Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Dim FileTest As String
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFile                      = " + Test.Parameter(eNomeFile) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eSincro                      = " + Test.Parameter(eSincro) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Corsa_Retract_mm                      = " + Test.Parameter(Corsa_Retract_mm) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Corsa_Extend_mm                      = " + Test.Parameter(Corsa_Extend_mm) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Corsa_Tolleranza_mm                      = " + Test.Parameter(Corsa_Tolleranza_mm) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
        
    
End Sub


