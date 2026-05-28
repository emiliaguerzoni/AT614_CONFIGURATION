Attribute VB_Name = "Module_TEST_RISPOSTE_GRADINO"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 200
Private Const ACCENSIONE_TIMEOUT_MS = 200
Private Y(2) As Double
'eTEST_OK = 2
'eTEST_BUSY = 1
'eTEST_STOP = 0
'eTEST_ERROR = -1


Private Type TagRisposteGradinoPunto
    Time As Double
    Trigger As Double
    Variabile As Double
    State As Long
End Type

' Stato macchina del modulo
Private Enum eMachineState
    eIdle = 0
    eInit
    eSpegniModuli
    eAccendiModulo
    eCapture
    eAttesaTrigger
    eTriggering
    eTimeOut
    eControlloNumeroSequenza
    eDecisioneFinale
    eControlloFinestraTemporale
    esavefile
    eEnd
    eStart
    eError
End Enum


Private Enum eTestResult
    eTestResult_OK
    eTestResult_ERROR
End Enum


' Struttura dati utilizzata per identificare la variabile di controllo
Private Type TagVariabile
    Value As Double             ' Valore associato alla variabile di controllo
    bValidate As Boolean        ' Flag di validità della variabile di controllo
    TimeStamp As Long           ' Istante di ricezione della variabile di controllo
    TickCount As Long           ' Tick di sistema del PC
End Type



' Enumeratore delle colonne dei parametri utilizzati per gestire il test
Private Enum eTestParameter
    eGeneric_NomeFileCurvaComando = 1           ' Nome del file del comando
    eGeneric_Rampa                              ' Cartella dove salvare il file
    eGeneric_Command_SA                         ' Source address del messaggio AVC da inviare
    eGeneric_NumeroCicli                        ' Numero di test da effettuare
    
    
    eVariabile_Nome                             ' Nome della variabile di controllo da utilizzare
    eVariabile_ValoreTest                       ' Valore di fine controllo della variabile utilizzata
    eVariabile_Fronte                           ' Fronte di fine controllo della variabile utilizzata
    eVariabile_FinestraMassima_ms               ' Finestra massima di accettabilità
    
    ' Parametri relativi alla generazione del trigger sul comando per il calcolo del tempo trascorso
    Trigger_Value                              ' Valore del Trigger della variabile scelta di inizio controllo
    Trigger_Fronte                             ' Fronte del trigger ( salita / discesa ) di inizio controllo
    Trigger_Tipo                               ' Tipo della variabile scelta come ingresso di trigger
    
    eTest_Bloccante                            ' Indica se il test da effettuare è bloccante o meno
    
    eNumeroMassimoDiParametri
End Enum

' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagPuntiAcquisizioneRisposteGradino
    X As Double                         ' Asse X del file salvato
    Y As Double                         ' Prima variabile salvata nel file di log ( variabile di controllo )
    Z As Double                         ' Seconda varibile salvata nel file di log ( variabile di supervisione )
    MLT_Temperature As Double           ' Temperatura del modulo MLT salvata nel file di log
    Pressure As Double                  ' Pressione letta dalla scheda elettronica salvata nel file di log
End Type

Public Type TagTEST_RISPOSTE_GRADINO
    bOldOn As Boolean                               ' difu del comando di start del modulo
    Return As Integer                               ' Stato del modulo
    State As Integer
    MLT_OldState As Integer                         ' Old dello stato del multidrom
    NomeFile As String                              ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean                           ' Flag segnalazione primo avvio del programma
    NodeCaptured As Integer
    Timer As Long
    CURVE_LIMITE As CurveLimitePlusTag              ' Struttura dati delle curve limite
    CURVE_COMANDO As TagChannel                     ' File di configurazione delle curve comando CAN in funzione di V teorico
    CURVE_XY As TagChannelRampa                     ' File di configurazione della rampa usata per effettuare l'acquisizione
    DeltaTime() As Long                                 ' Valore del delta di acquisizione
    Tempo_sec_start As Double                       ' Istante di avvio della sequenza espresso in secondi
    NumeroPuntiAcquisiti As Long                    ' numero di punti acquisiti
    CommandInMilliVolts As Double                   ' Comando in millivolt
    OldTest(eNumeroMassimoDiParametri) As String
    Trigger_Old As TagVariabile                     ' Old della variabile usata come trigger
    Variabile_Old As TagVariabile                   ' Old della variabile usata come variabile di controllo
    NumeroTestEffettuati As Long                    ' Numero di test effettuati
    NumeroFallimenti As Long                        ' Numero di fallimenti del test
    NumeroTestOk As Long                            ' Numero di test effettuati correttamente
    
    
    MLT_CommandState As MltState                    ' Stato del segnale di comando AVC generato dal modulo sw
    MLT_CommandFlow As Integer                         ' Flow del segnale di comando AVC generato dal modulo sw
    
    StartEvent As TagVariabile
    EndEvent As TagVariabile
    
    
    AVC_Flow As TagVariabile                        ' Variabile interna usata per triggerare le variabili generate dal software
    AVC_State As TagVariabile                       ' Variabile interna usata per triggerare le variabili generate dal software
    
    
    Punto() As TagRisposteGradinoPunto
    PathSalvaPunti As String
    
End Type


Private Context As TagTEST_RISPOSTE_GRADINO




' Funzione di inizializzazione dei settaggi del grafico. Carica le curve limite
Private Sub GraphSettings(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
On Error GoTo err:
    Dim bUpdate As Boolean
    Dim i As Integer
    Dim par As CurveLimitePlusParametersTag
    

    ' inizializzazione delle curve limite
    par.Indice(eInferiore) = 1
    par.Indice(eSuperiore) = 2

    CurveLimitePlusInit Context.CURVE_LIMITE, par


    If (Context.OldTest(eGeneric_NomeFileCurvaComando) <> Test.Parameter(eGeneric_NomeFileCurvaComando)) Then
        ' Inizializzazione dei comandi da applicare
        If (Update_CURVECOMANDO(Context.CURVE_COMANDO, Settings.FolderFileCurveComando + "\" + Test.Parameter(eGeneric_NomeFileCurvaComando)) = True) Then
            MsgBox "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI COMANDO CAN ", vbCritical
            
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eGeneric_NomeFileCurvaComando) = Test.Parameter(eGeneric_NomeFileCurvaComando)
    End If


    If (Context.OldTest(eGeneric_Rampa) <> Test.Parameter(eGeneric_Rampa)) Then

         ' Inizializzazione dei file della rampa da applicare
        If (Update_CURVE_XY(Context.CURVE_XY, Settings.FolderRampeXY + "\" + Test.Parameter(eGeneric_Rampa)) = True) Then
            MsgBox "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI RAMPE", vbCritical
            
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eGeneric_Rampa) = Test.Parameter(eGeneric_Rampa)
    End If

    Infos.RichTextBox1.BackColor = vbWhite

    
Exit Sub
err:
Debug.Print err.Description
End Sub




Private Sub INFO_RISPOSTE_GRADINO(Infos As TagInfosSingleTest)
    If (Infos.CWGraph1.Visible <> False) Then
        Infos.CWGraph1.Visible = False
    End If
    
    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
    End If
    
    
    If (Infos.RichTextBox1.Visible = False) Then
        Infos.RichTextBox1.Visible = True
    End If
    
End Sub



Private Sub TEST_RISPOSTE_GRADINO_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_RISPOSTE_GRADINO)
    Context1.State = eEnd
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = "RISPOSTE GRADINO"
    End If
End Sub


Public Sub TEST_RISPOSTE_GRADINO(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_RISPOSTE_GRADINO, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "RISPOSTE_GRADINO") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        INFO_RISPOSTE_GRADINO Infos
        
        If (bStop = True And bOn = True) Then
        
        TEST_RISPOSTE_GRADINO_INIT_STRUCT Infos, Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then          ' fronte di salita del test ( AVVIO )
            Init Test, Infos, IndiceSezione, Context1                                         ' init delle variabili del modulo
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
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
On Error GoTo err:
    
    Context.Return = eTEST_BUSY
        
    ReDim Context.Punto(0)
    
    Context.State = eSpegniModuli
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.bFirstTime = True
    
    GraphSettings Test, Infos, IndiceSezione, Context
    
    Context.NumeroPuntiAcquisiti = 0
    ReDim Context.DeltaTime(0)
    Infos.RichTextBox1.Text = ""
    Infos.RichTextBox1.Font = "Courier"

    Infos.RichTextBox1.Text = vbCrLf + vbCrLf + "LISTA PARAMETRI" + vbCrLf _
                                              + "---------------" + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "File Curva Comando = " + Test.Parameter(eGeneric_NomeFileCurvaComando) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "File Rampa         = " + Test.Parameter(eGeneric_Rampa) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "CAN Source address =" + Test.Parameter(eGeneric_Command_SA) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Numero Cicli       =" + Test.Parameter(eGeneric_NumeroCicli) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Variabile Tipo     =" + Test.Parameter(eVariabile_Nome) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Variabile Valore   =" + Test.Parameter(eVariabile_ValoreTest) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Variabile Fronte   =" + Test.Parameter(eVariabile_Fronte) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Variabile Finestra =" + Test.Parameter(eVariabile_FinestraMassima_ms) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Trigger Tipo       =" + Test.Parameter(Trigger_Tipo) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Trigger Value      =" + Test.Parameter(Trigger_Value) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "Trigger Fronte     =" + Test.Parameter(Trigger_Fronte) + vbCrLf
    
    Context.NumeroTestEffettuati = 0
    Context.NumeroFallimenti = 0
    Context.NumeroTestOk = 0
      
    Context.PathSalvaPunti = ""
    Exit Sub
err:
Debug.Print err.Description
End Sub

' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Select Case Context.State
    Case eIdle:
    Case eInit: Init Test, Infos, IndiceSezione, Context
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, Context
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, Context
    Case eCapture: Capture Test, Infos, IndiceSezione, Context
    Case eAttesaTrigger:
    
        ' Stato di attesa del trigger
        AttesaTrigger Test, Infos, IndiceSezione, Context
        GenerazioneComando Test, Infos, IndiceSezione, Context
        
    Case eTriggering:
        ' Avvio del trigger !!!!!
        Triggering Test, Infos, IndiceSezione, Context
        GenerazioneComando Test, Infos, IndiceSezione, Context
        
    Case eTimeOut: Timeout Test, Infos, IndiceSezione, Context
    Case eControlloNumeroSequenza: ControlloNumeroSequenza Test, Infos, IndiceSezione, Context
    Case eDecisioneFinale: DecisioneFinale Test, Infos, IndiceSezione, Context
    Case eStart: Start Test, Infos, IndiceSezione, Context
    Case eControlloFinestraTemporale: ControlloFinestraTemporale Test, Infos, IndiceSezione, Context
    Case esavefile: SaveFile Test, Infos, IndiceSezione, Context
    Case eEnd: Fine Test, Infos, IndiceSezione, Context
    Case eError: ErrorTest Test, Infos, IndiceSezione, Context
    Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
End Sub


Private Function Formatx(str As String, spacing As String, dimension As Integer)

    Formatx = str
    Do While (Len(Formatx) < dimension)
        Formatx = spacing + Formatx
    Loop
End Function

Private Sub ControlloFinestraTemporale(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    
    Dim FinestraTemporale As Long
    
    FinestraTemporale = CLng(Test.Parameter(eVariabile_FinestraMassima_ms))
    
    
    If (Context.EndEvent.TimeStamp <> -1 And Context.StartEvent.TimeStamp <> -1) Then
        Context.DeltaTime(Context.NumeroTestEffettuati) = Context.EndEvent.TimeStamp - Context.StartEvent.TimeStamp
    Else
        Context.DeltaTime(Context.NumeroTestEffettuati) = Context.EndEvent.TickCount - Context.StartEvent.TickCount
    End If
    
    If (Context.DeltaTime(Context.NumeroTestEffettuati) > FinestraTemporale) Then
        Context.NumeroFallimenti = Context.NumeroFallimenti + 1
        Infos.RichTextBox1.Text = " DELTA=" + Formatx(CStr(Context.DeltaTime(Context.NumeroTestEffettuati)), " ", 4) + " ms " + "ERR" + Infos.RichTextBox1.Text
    ElseIf (Context.DeltaTime(Context.NumeroTestEffettuati) = 0) Then
        Infos.RichTextBox1.Text = " DELTA=" + Formatx(CStr(Context.DeltaTime(Context.NumeroTestEffettuati)), " ", 4) + " ms " + "DELETE" + Infos.RichTextBox1.Text
    Else
        Context.NumeroTestOk = Context.NumeroTestOk + 1
        Infos.RichTextBox1.Text = " DELTA=" + Formatx(CStr(Context.DeltaTime(Context.NumeroTestEffettuati)), " ", 4) + " ms " + "OK" + Infos.RichTextBox1.Text
    End If
    
    Context.State = eControlloNumeroSequenza
    
End Sub


Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Infos.RichTextBox1.BackColor = vbRed
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = "- ERRORE SUL TEST" + vbLf + Infos.RichTextBox1.Text
    End If
    Context.Return = eTEST_ERROR

End Sub


Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    'TurnOffAllModule
    TurnOffModule IndiceSezione
    If (Context.Timer < GetTickCount) Then
        Context.State = eAccendiModulo
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)

    NodeCaptured(IndiceSezione) = IndiceSezione + 1
    
    Context.State = eStart
End Sub

Private Sub SaveFile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    On Error Resume Next
    
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    
    Context.State = eEnd
End Sub

Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)

    If (Infos.RichTextBox1.BackColor = vbRed) Then
        
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
End Sub

Private Sub Info_showTriggeringEvent(Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    
    'Infos.RichTextBox1.Text = " DT=" + Formatx(CStr(Context.EndEvent.Value), " ", 4) + Infos.RichTextBox1.Text
'    Infos.RichTextBox1.Text = "TICK =" + CStr(Context.EndEvent.TickCount) + vbCrLf + Infos.RichTextBox1.Text
'    Infos.RichTextBox1.Text = "TIME =" + CStr(Context.EndEvent.TimeStamp) + vbCrLf + Infos.RichTextBox1.Text
    'Infos.RichTextBox1.Text = CStr(Context.NumeroTestEffettuati) + "- END TRIGGER VALUE=" + vbCrLf + Infos.RichTextBox1.Text
End Sub

Private Sub Info_showTriggerEvent(Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)

   
    'Infos.RichTextBox1.Text = " TR=" + Formatx(CStr(Context.StartEvent.Value), " ", 4) + Infos.RichTextBox1.Text
    'Infos.RichTextBox1.Text = "TICK =" + CStr(Context.StartEvent.TickCount) + vbCrLf + Infos.RichTextBox1.Text
    'Infos.RichTextBox1.Text = "TIME =" + CStr(Context.StartEvent.TimeStamp) + vbCrLf + Infos.RichTextBox1.Text
    'Infos.RichTextBox1.Text = CStr(Context.NumeroTestEffettuati) + "- START TRIGGER -------" + vbCrLf + Infos.RichTextBox1.Text
End Sub

 ' Sono arrivato alla fine della scansione della curva di comando Devo segnalare l'errore
Private Sub Timeout(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    
    Infos.RichTextBox1.Text = CStr(Context.NumeroTestEffettuati) + "- TIMEOUT " + Infos.RichTextBox1.Text
    
    Context.NumeroFallimenti = Context.NumeroFallimenti + 1
    
    Context.State = eControlloNumeroSequenza
End Sub

 ' Sono arrivato alla fine della scansione della curva di comando Devo segnalare l'errore
Private Sub ControlloNumeroSequenza(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Dim NumeroDiTest As Long
    Dim str As String
    
    str = CStr(Context.NumeroTestEffettuati)
    Do While (Len(str) < 2)
        str = "0" + str
    Loop
    
    str = str + "-"
    Infos.RichTextBox1.Text = vbCrLf + str + Infos.RichTextBox1.Text
    
'    Infos.RichTextBox1.Text = "- CONTROLLO SEQUENZA -------" + vbCrLf + Infos.RichTextBox1.Text
    
    
    ' Prelevo il numero di test da effettuare
    NumeroDiTest = CLng(Test.Parameter(eGeneric_NumeroCicli))
    
    Context.NumeroTestEffettuati = Context.NumeroTestEffettuati + 1
    ReDim Preserve Context.DeltaTime(Context.NumeroTestEffettuati)
    
    If (Context.NumeroTestEffettuati < NumeroDiTest) Then
        Context.State = eStart
    Else
        Context.State = eDecisioneFinale
    End If
    
End Sub


Private Sub CalcoloValoreMedio(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Dim calc As Double
    Dim i As Long
    
    Do While (i < Context.NumeroTestEffettuati)
        calc = calc + Context.DeltaTime(i)
        i = i + 1
    Loop
    
    calc = calc / Context.NumeroTestEffettuati
    Infos.RichTextBox1.Text = "------- VALORE MEDIO -------" + vbCrLf + _
                              "VALOR MEDIO     =" + Format(calc, "0.0") + vbCrLf + _
                              Infos.RichTextBox1.Text

    
    
End Sub



' Sono arrivato alla fine della scansione della curva di comando Devo segnalare l'errore
Private Sub DecisioneFinale(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Dim NumeroDiTest As Long
    
    CalcoloValoreMedio Test, Infos, IndiceSezione, Context
    
     SalvaPunti Test, Infos, Context, IndiceSezione
     
    Infos.RichTextBox1.Text = "------- FILE DATI -------" + vbCrLf + _
                              Context.PathSalvaPunti + vbCrLf + _
                              Infos.RichTextBox1.Text
    
   
    
    
    Infos.RichTextBox1.Text = "------- DECISIONE FINALE -------" + vbCrLf + _
                              "TEST EFFETTUATI =" + CStr(Context.NumeroTestEffettuati) + vbCrLf + _
                              "TEST FALLITI    =" + CStr(Context.NumeroFallimenti) + vbCrLf + _
                              "TEST OK         =" + CStr(Context.NumeroTestOk) + vbCrLf + _
                              Infos.RichTextBox1.Text
    
    
     
    
    If (Context.NumeroFallimenti > 0) Then
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eError
    Else
        Context.State = esavefile
    End If
    
    
    If (Context.NumeroFallimenti > Context.NumeroTestOk) Then
        
    Else
        
    End If
        
End Sub




 
' Questo stato stato viene attivato successivamente all'attivazione del trigger
Private Sub Triggering(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    
    Dim VariabileDiControllo As TagVariabile
    Dim LivelloVariabileDiControllo As Double
    Dim Trigger As TagVariabile
    

    ' Prelevo la variabile di controllo
    VariabileDiControllo = Ottieni_Variabile(Test.Parameter(eVariabile_Nome), IndiceSezione, Context)
    Trigger = Ottieni_Variabile(Test.Parameter(Trigger_Tipo), IndiceSezione, Context)
    
    
    Context.Punto(UBound(Context.Punto)).Time = GetTickCount / 1000
    Context.Punto(UBound(Context.Punto)).State = 0
    Context.Punto(UBound(Context.Punto)).Trigger = Trigger.Value
    Context.Punto(UBound(Context.Punto)).Variabile = VariabileDiControllo.Value
    ReDim Preserve Context.Punto(UBound(Context.Punto) + 1)

    
    
    
    
    
        
    LivelloVariabileDiControllo = CDbl(Test.Parameter(eVariabile_ValoreTest))           ' Prelevo il livello da controllare
    
    
    If (VariabileDiControllo.bValidate = True) Then
        ' Controllo il fronte
        
        If (Test.Parameter(eVariabile_Fronte) = "SALITA") Then
            
            If (Context.Variabile_Old.Value < LivelloVariabileDiControllo And _
                 VariabileDiControllo.Value >= LivelloVariabileDiControllo) Then
                 
                 Context.EndEvent = VariabileDiControllo
                 
                 Info_showTriggeringEvent Infos, IndiceSezione, Context
                 
                 Context.State = eControlloFinestraTemporale
                 
            End If
            
        ElseIf (Test.Parameter(eVariabile_Fronte) = "UGUALE") Then
            
            If (VariabileDiControllo.Value = LivelloVariabileDiControllo) Then
            
                 Context.EndEvent = VariabileDiControllo
                 
                 Info_showTriggeringEvent Infos, IndiceSezione, Context
                 
                 Context.State = eControlloFinestraTemporale
            End If
        ElseIf (Test.Parameter(eVariabile_Fronte) = "DISCESA") Then
            
            
            If (Context.Variabile_Old.Value > LivelloVariabileDiControllo And _
                 VariabileDiControllo.Value <= LivelloVariabileDiControllo) Then
                 
                 Context.EndEvent = VariabileDiControllo
                 
                 Info_showTriggeringEvent Infos, IndiceSezione, Context
                 
                 Context.State = eControlloFinestraTemporale
            End If
            
            
        End If
        
    
        'Salvo solo ed esclusivamente se la variabile ricevuta è valida
        Context.Variabile_Old = VariabileDiControllo
    End If
    
 
End Sub


Private Sub AttesaTrigger(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Dim Tempo_sec As Double
    
    
    Dim bPunti_Y_Validi As Boolean
    Dim Percentual As Integer
    Dim LivelloTrigger As Double
    
    
    
    Dim VariabileDiControllo As TagVariabile
    Dim Trigger As TagVariabile
    

    On Error Resume Next

    
    ' Prelievo la viariabile di controllo in base alla configurazione del parametro
    VariabileDiControllo = Ottieni_Variabile(Test.Parameter(eVariabile_Nome), IndiceSezione, Context)
    Trigger = Ottieni_Variabile(Test.Parameter(Trigger_Tipo), IndiceSezione, Context)
    
    Context.Punto(UBound(Context.Punto)).Time = GetTickCount / 1000
    Context.Punto(UBound(Context.Punto)).State = 0
    Context.Punto(UBound(Context.Punto)).Trigger = Trigger.Value
    Context.Punto(UBound(Context.Punto)).Variabile = VariabileDiControllo.Value
    ReDim Preserve Context.Punto(UBound(Context.Punto) + 1)
    
    
    
    LivelloTrigger = CDbl(Test.Parameter(Trigger_Value))
    
    ' Eseguo il test solo se le variabili sono valide
    If (Context.Trigger_Old.bValidate = True And Trigger.bValidate = True) Then
    
        ' Rimango in attesa dell'attivazione del trigger !!!!!!!
        If (Test.Parameter(Trigger_Fronte) = "SALITA") Then
            
            ' Controllo se il trigger è partito !!!!
            If (Context.Trigger_Old.Value < LivelloTrigger And _
                Trigger.Value >= LivelloTrigger) Then
                
                Context.StartEvent = Trigger            ' Memorizzo la variabile dell'inizio dell'evento
                Context.State = eTriggering             ' Cambio stato
                
                ' Visualizzo le info dell'evento
                Info_showTriggerEvent Infos, IndiceSezione, Context
                
                
            End If
        
        ElseIf (Test.Parameter(Trigger_Fronte) = "UGUALE") Then
             
             If (Trigger.Value = LivelloTrigger) Then
             
                Context.StartEvent = Trigger            ' Memorizzo la variabile dell'inizio dell'evento
                Context.State = eTriggering             ' Cambio stato
                
                ' Visualizzo le info dell'evento
                Info_showTriggerEvent Infos, IndiceSezione, Context
            End If
            
        ElseIf (Test.Parameter(Trigger_Fronte) = "DISCESA") Then
    
            ' Controllo se il trigger è partito !!!!
            If (Context.Trigger_Old.Value > LivelloTrigger And _
                Trigger.Value <= LivelloTrigger) Then
                
                Context.StartEvent = Trigger            ' Memorizzo la variabile dell'inizio dell'evento
                Context.State = eTriggering             ' Cambio stato
                
                ' Visualizzo le info dell'evento
                Info_showTriggerEvent Infos, IndiceSezione, Context
                            
            End If
            
        End If
    End If
    
    'Salvo solo ed esclusivamente se la variabile ricevuta è valida
    If (Trigger.bValidate = True) Then
        If (Abs(Context.Trigger_Old.TimeStamp - Trigger.TimeStamp) > 70) Then
            Debug.Print (CStr(Context.Trigger_Old.TimeStamp - Trigger.TimeStamp))
        End If
        Context.Trigger_Old = Trigger
    End If
    
    'Salvo solo ed esclusivamente se la variabile ricevuta è valida
    If (VariabileDiControllo.bValidate = True) Then
        Context.Variabile_Old = VariabileDiControllo
    End If
                    
  
End Sub


Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    Context.State = eAttesaTrigger
    Context.Tempo_sec_start = GetTickCount / 1000
    
   ' Infos.RichTextBox1.Text = "------- START -------" + vbCrLf + Infos.RichTextBox1.Text
    
    
End Sub

Private Sub Send_AVC_MESSAGE(ByRef State As MltState, ByRef Percentual As Integer, ByVal IndiceSezione As Integer, ByRef SourceAddress As String)
    Dim strState As String
    Dim strPercentual As String
    Dim str As String
    'trasformo in Hex
    strState = Hex(State)
    strPercentual = Hex(Percentual)
    
    'devo sempre avere 2 cifre Hex
    Do While (Len(strPercentual) < 2)
        strPercentual = "0" + strPercentual
    Loop
    
          'T18FE3    NodeId di CH1    default = 22           byte0          byte1    byte2 = stato     'byte 3-8
    str = "T18FE3" & NodeCaptured(IndiceSezione) & SourceAddress & "8" & strPercentual & "FF" & "F" & strState & "FFFFFFFFFF" & vbCr
    Output str

End Sub



Private Function CreaDirectory(Infos As TagInfosSingleTest, ByVal StartDirectory As String) As String
    On Error GoTo err:
    Dim PathDirectory  As String
        
        CreaDirectory = StartDirectory + "\" + Infos.codicePRodotto + "\"
        
        MkDir (CreaDirectory)
'        'If (DirExists(CreaDirectory) = False) Then
'        If (Dir(CreaDirectory) = "") Then
'            MkDir (CreaDirectory)
'        End If
'
'        If (Dir(CreaDirectory) = "") Then
'            CreaDirectory = ""
'            MsgBox "ERRORE NON SI RIESCE A CREARE LA CARTELLA DEL CODICE DEL PRODOTTO", vbCritical
'        End If
        Exit Function
err:
'        CreaDirectory = ""
'        Debug.Print err.Description
'        MsgBox "ERRORE NON SI RIESCE A CREARE LA CARTELLA DEL CODICE DEL PRODOTTO", vbCritical
End Function



Public Function DirExists(ByVal Path As String) As Boolean
    On Error Resume Next
    'Legge l'attributo e si assicura che si tratti di una directory
    FileExists = GetAttr(Path) And vbDirectory
    
    
    'Se avviene un errore la Function restituisce False
End Function



Private Function TestRispostaGradino(ByVal State As MltState, ByVal Percentage As Integer, Time As Double, VariabileControllo As TagVariabile) As eTestResult

End Function


' Ritorno la variabile di controllo ed la sua validità in base ai parametri imposti del test
Private Function Ottieni_Variabile(ByRef StringaVariabileDiControllo As String, ByVal IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO) As TagVariabile

    Dim TempVariabile As TagVariabile

    'TempVariabile

    Select Case StringaVariabileDiControllo
    
    Case "MLT - STATE"                           ' La variabile di controllo è la posizione del sensore
    
    
        TempVariabile.bValidate = MLT(IndiceSezione).bActive
        TempVariabile.Value = MLT(IndiceSezione).State
        TempVariabile.TimeStamp = MLT(IndiceSezione).TimeStamp
        TempVariabile.TickCount = MLT(IndiceSezione).TickCount
    
    Case "MLT - POSITION"                           ' La variabile di controllo è la posizione del sensore
    
        ' Controllo la validità dei messaggi ricevuti dal modulo
        TempVariabile.bValidate = MLT(IndiceSezione).bActive
        TempVariabile.Value = MLT(IndiceSezione).Position
        TempVariabile.TimeStamp = MLT(IndiceSezione).TimeStamp
        TempVariabile.TickCount = MLT(IndiceSezione).TickCount
    
    Case "MLT - SETPOINT"                           ' La variabile di controllo è la posizione del sensore
    
        ' Controllo la validità dei messaggi ricevuti dal modulo
        TempVariabile.bValidate = MLT(IndiceSezione).bActive
        TempVariabile.Value = MLT(IndiceSezione).Setpoint
        TempVariabile.TimeStamp = MLT(IndiceSezione).TimeStamp
        TempVariabile.TickCount = MLT(IndiceSezione).TickCount
        
    Case "CE16 - POSITION"                          ' La variabile di controllo è la posizione del sensore CE16
        TempVariabile.bValidate = CE16(IndiceSezione).bActive
        TempVariabile.Value = CE16(IndiceSezione).Position
        TempVariabile.TimeStamp = CE16(IndiceSezione).TimeStamp
        TempVariabile.TickCount = CE16(IndiceSezione).TickCount
        
    Case "AVC - STATE"
        
        TempVariabile = Context.AVC_State
        
'        TempVariabile.bValidate = True
'        TempVariabile.Value = Context.MLT_CommandState
'        TempVariabile.TimeStamp = -1
'        TempVariabile.TickCount = GetTickCount
    
    Case "AVC - FLOW"
        TempVariabile = Context.AVC_Flow
        ' Controllo la validità dei messaggi ricevuti dal modulo
'        TempVariabile.bValidate = True
'        TempVariabile.Value = Context.MLT_CommandFlow
'        TempVariabile.TimeStamp = -1                    ' Non valido
'        TempVariabile.TickCount = GetTickCount
    
    Case "CMD - MILLIVOLT"
    
        ' Controllo la validità dei messaggi ricevuti dal modulo
        TempVariabile.bValidate = True
        TempVariabile.Value = Context.CommandInMilliVolts
        TempVariabile.TimeStamp = -1                    ' Non valido
        TempVariabile.TickCount = GetTickCount
        
    Case "DP1 - DISTRVALPOS"
        ' Controllo la validità dei messaggi ricevuti dal modulo
        TempVariabile.bValidate = DP1(IndiceSezione).bActive
        TempVariabile.Value = DP1(IndiceSezione).DistrValPos
        TempVariabile.TimeStamp = DP1(IndiceSezione).TimeStamp
        TempVariabile.TickCount = DP1(IndiceSezione).TickCount
        
    Case Else
        
        
    End Select
    Ottieni_Variabile = TempVariabile
End Function


' Funzione di generazione del comando
Private Sub GenerazioneComando(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_RISPOSTE_GRADINO)
    
    Dim Tempo_sec As Double
    Dim TickDiSistema As Long
    
    Dim Percentual As Integer
    Dim State As MltState
    
   ' If (Context.Timer < GetTickCount) Then
        
        TickDiSistema = GetTickCount
        
        Tempo_sec = TickDiSistema / 1000 - Context.Tempo_sec_start
        
        Context.Timer = TickDiSistema + 10
        
        ' Prelevo il comando dalla rampa applicata
        If (Get_CURVE_XY_VALUE(Context.CURVE_XY, Tempo_sec, Context.CommandInMilliVolts) = eERROR_Rampe) Then
                        
            Context.State = eTimeOut            ' ho raggiunto la fine della rampa da applicare... se non è stato generato nessun trigger vado in errore di timeout
        Else
            If (Get_CURVECOMANDO_VALUE(Context.CURVE_COMANDO, Context.CommandInMilliVolts / 1000, Context.MLT_CommandState, Context.MLT_CommandFlow) = eError_CURVE_COMANDO_CAN) Then
                SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
                Context.State = eError
            Else
            
            End If
        End If
        
                
        Send_AVC_MESSAGE Context.MLT_CommandState, Context.MLT_CommandFlow, IndiceSezione, Test.Parameter(eGeneric_Command_SA)      ' Invio il messaggio CAN
        
        
        Context.AVC_Flow.bValidate = True
        Context.AVC_Flow.TickCount = TimeStamp 'TickDiSistema
        Context.AVC_Flow.TimeStamp = -1
        Context.AVC_Flow.Value = Context.MLT_CommandFlow
        
        Context.AVC_State.bValidate = True
        Context.AVC_State.TickCount = TimeStamp ' TickDiSistema
        Context.AVC_State.TimeStamp = -1
        Context.AVC_State.Value = Context.MLT_CommandState
   ' End If
    
                
End Sub


Private Sub SalvaPunti(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_RISPOSTE_GRADINO, IndiceSezione As Integer)

    Dim bSaveLastAcq As Boolean
On Error GoTo err:
    Dim PathDirectory As String
    Dim i As Long
    Dim ora As String
    Dim oggi As String
    Dim Intestazioni() As String
    
    
    ora = Format(Time, ("HH.MM.SS"))
    ora = Replace(ora, ".", "")

    oggi = Format(Date, ("yy/mm/dd"))
    oggi = Replace(oggi, "/", "")
    
    
    PathDirectory = CreaDirectory(Infos, Settings.FolderGraphSaved)
    
    If (PathDirectory = "") Then
        PathDirectory = Settings.FolderGraphSaved + "\"
    End If
    
    If (Infos.Label1 <> "") Then
        Context.PathSalvaPunti = PathDirectory + Infos.Label1.Caption + "_" + Test.Name + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
    Else
        Context.PathSalvaPunti = PathDirectory + Test.Name + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
    End If
    
    ScriviCSV Context.PathSalvaPunti, Context.Punto
    
    
    Exit Sub
err:

End Sub



Function ScriviCSV(FilePath As String, Dati() As TagRisposteGradinoPunto)
    Dim fileNum As Integer
    Dim i As Integer
    Dim Riga As String
    
    ' Apre il file per la scrittura
    fileNum = FreeFile
    Open FilePath For Output As fileNum
    
    ' Scrive l'intestazione nel file CSV
    Print #fileNum, "TIME;TRIGGER;VARIABLES;STATE"
    
    
    Riga = ""
    ' Scrive i dati nel file CSV
    For i = LBound(Dati) To UBound(Dati) - 1
        Riga = Format(Dati(i).Time, "0.000") + ";" + _
               Format(Dati(i).Trigger, "0.0") + ";" + _
               Format(Dati(i).Variabile, "0.0") + ";" + _
               Format(Dati(i).State, "0.0")
        Print #fileNum, Riga
    Next i
    
    ' Chiude il file
    Close fileNum
End Function




