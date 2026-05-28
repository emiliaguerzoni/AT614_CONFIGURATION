Attribute VB_Name = "Module_TEST_CICLICA_CAN"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 200
Private Const ACCENSIONE_TIMEOUT_MS = 200
'eTEST_OK = 2
'eTEST_BUSY = 1
'eTEST_STOP = 0
'eTEST_ERROR = -1


'STATI DELL' MLT
Private Enum MltState
    Neutral = 0
    Extend = 1
    Retract = 2
    Float = 3
    Alarm = 4
End Enum

' Stato macchina del modulo
Private Enum eMachineState
    eInit = 0
    eSpegniModuli = 1
    eAccendiModulo = 2
    eCapture = 3
    eRun = 4
    eEnd = 5
    eError = 6
    eNeutro = 7
    eFloat = 8
    eMaxE = 9
End Enum


Private Enum eTestParameter
    eTimerCiclica = 1                       ' Tempo totale della ciclica
    eTimerCiclicaSingoloCiclo = 2           ' Tempo della singola forma d'onda della ciclica
    eSourceAddress = 3                      ' Source address del messaggio AVC da trasmettere
    eRiferimentoNeutro = 4                  ' Riferimento del comando NEUTRO ( se minore a 32 è riferito al CE16 altrimenti al position del MLT )
    eRiferimentoMaxR = 5                    ' Riferimento del comando MAX R ( se minore a 32 è riferito al CE16 altrimenti al position del MLT )
    eRiferimentoMaxE = 6                    ' Riferimento del comando MAX E ( se minore a 32 è riferito al CE16 altrimenti al position del MLT )
    eTolleranza = 7                         ' Valore della finestra di accettabilità da sommare e sottrarre al valore di riferimento per confermare la posizione del modulo
End Enum


Public Type TagTEST_CICLICA
    bOldOn As Boolean               ' difu del comando di start del modulo
    Return As Integer                ' Stato del modulo
    State As Integer                    ' Stato macchina del modulo software
    OldState As Integer                 ' Old stato macchina del modulo software
    Timer As Long
    ErrorePosizioni As Integer          ' Numero di errori di posizione della spola
    NumeroCicli As Long                 ' Numero di cicli effettuati
    OldErrorePosizioni As Integer       ' Old numero errori di posizione spola
    TimerCiclica As Long
    TimerSendMesssage As Long
    bFirstCicloMaxE As Boolean          ' Flag di segnalazione primo ciclo MaxE
    bFirstCicloMaxR As Boolean          ' Falg di segnalazione primo ciclo MaxR
    SensoreMLT_FirstMaxE As Long         ' Valore Iniziale del sensore LMT lato Extend
    SensoreMLT_FirstMAXR As Long         ' Valore Iniziale del sensore MLT lato Retract
    DeltaMLT_MaxE As Long               ' Differenza letta del sensore di al variare della ciclica
    DeltaMLT_MaxR As Long               ' Differenza letta del sensore di al variare della ciclica
    SensoreMLT_MaxE As Long             ' Valore Iniziale del sensore LMT lato Extend
    SensoreMLT_MAXR As Long             ' Valore Iniziale del sensore MLT lato Retract
    
    NodeCaptured As Integer
End Type

Private Context As TagTEST_CICLICA

Public Sub INFO_TEST_CICLICA(Infos As TagInfosSingleTest)
    If (Infos.CWGraph1.Visible <> False) Then
        Infos.CWGraph1.Visible = False
    End If
    
    If (Infos.RichTextBox1.Visible = False) Then
        Infos.RichTextBox1.Visible = True
    End If

    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
    End If


End Sub



Private Sub TEST_CICLICA_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_CICLICA)
    Context1.State = eEnd
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = "CICLICA"
    End If
    
End Sub


Public Sub TEST_CICLICA(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_CICLICA, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "TEST_CICLICA_CAN") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
        bReturn = Context1.Return
        INFO_TEST_CICLICA Infos
        
        If (bStop = True And bOn = True) Then
            
            TEST_CICLICA_INIT_STRUCT Infos, Context1
            
        ElseIf (bOn = True And Context1.bOldOn = False) Then          ' fronte di salita del test ( AVVIO )
            Init Test, Infos, IndiceSezione, Context1                                         ' init delle variabili del modulo
            Infos.RichTextBox1 = Infos.RichTextBox1 + "START" + vbLf
            
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            Infos.RichTextBox1 = Infos.RichTextBox1 + "STOP" + vbLf
            
            Context1.State = eError
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        End If
        
        Task Test, Infos, IndiceSezione, Context1                                       ' Task del modulo
        
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
            
    
        'Context1.bOldOn = bOn                     ' memorizzo il fronte
        
    End If
    
End Sub

' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA)
    Dim i As Integer

    Context.Return = eTEST_BUSY_NON_BLOCCANTE
    
    Context.OldState = 0
    Context.State = eSpegniModuli
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    
    Do While (i < MAX_SEZIONI)
        MLTCANRelais(i) = True
        i = i + 1
    Loop
    Infos.RichTextBox1.Text = ""
    Context.bFirstCicloMaxE = True
    Context.bFirstCicloMaxR = True
    Context.SensoreMLT_MaxE = 0
    Context.SensoreMLT_MAXR = 0
    Context.SensoreMLT_FirstMaxE = 0
    Context.SensoreMLT_FirstMAXR = 0
    Context.NumeroCicli = 0
    Context.ErrorePosizioni = 0                                     ' Resetto gli errori di posizione
    Context.OldErrorePosizioni = 0
    Context.TimerSendMesssage = GetTickCount
    Context.Timer = GetTickCount
    
    
    Infos.RichTextBox1.BackColor = vbWhite
    
    Prepare_Textbox_Parametri Test, Infos
    
End Sub

' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA)
    Select Case Context.State
    Case eInit: Init Test, Infos, IndiceSezione, Context
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, Context
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, Context
    Case eCapture: Capture Test, Infos, IndiceSezione, Context
    Case eRun: Run Test, Infos, IndiceSezione, Context
    Case eNeutro: NeutroTest Test, Infos, IndiceSezione, Context
    Case eFloat:  FloatTest Test, Infos, IndiceSezione, Context
    Case eMaxE:   MaxeTest Test, Infos, IndiceSezione, Context
    Case eEnd: Fine Test, Infos, IndiceSezione, Context
    Case eError: ErrorTest Test, Infos, IndiceSezione, Context
    Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
    If (Context.OldState <> Context.State) Then
        'Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "STATE = " + CStr(Context.State) + vbLf
    End If
    
    Context.OldState = Context.State
End Sub

Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA)
    If (Infos.RichTextBox1.BackColor <> vbRed) Then
        Context.Return = eTEST_OK
    Else
        Context.Return = eTEST_ERROR
    End If
End Sub


Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA)
    
    Infos.RichTextBox1.BackColor = vbRed
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = "ERRORE SUL TEST" + vbLf + Infos.RichTextBox1.Text
    End If
    Context.Return = eTEST_ERROR
End Sub

Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA)
    TurnOffModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eAccendiModulo
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA)
    If (NodeCaptured(IndiceSezione) = -1) Then
        Context.State = eError
        AggiungiRiga Settings.FolderLOG, Infos, "ERROR Capture"
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        'NodeCaptured(IndiceSezione) = CatturaMltPlus()
        'Context.NodeCaptured = NodeCaptured(IndiceSezione)
        'ExitCatturaMlt Context.NodeCaptured
    Else
        'Infos.RichTextBox1 = Infos.RichTextBox1 + "NODO RIFERIMENTO= " + CStr(NodeCaptured(IndiceSezione)) + vbLf
        'Infos.RichTextBox1 = Infos.RichTextBox1 + "SOURCE ADDRESS= " + Test.Parameter(eSourceAddress) + vbLf
        
    End If
    
    NodeCaptured(IndiceSezione) = IndiceSezione + 1
    Context.NodeCaptured = NodeCaptured(IndiceSezione)
    
    Infos.RichTextBox1.Text = "ID = " + CStr(NodeCaptured(IndiceSezione)) + vbLf + Infos.RichTextBox1.Text
    
    Context.State = eRun
    
End Sub

Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA)
    Context.Timer = GetTickCount + CDbl(Test.Parameter(eTimerCiclicaSingoloCiclo)) * 1000
    Context.TimerCiclica = GetTickCount + CDbl(Test.Parameter(eTimerCiclica)) * 1000


    Infos.RichTextBox1.Text = "7 - TOLLERANZA = " + Test.Parameter(eTolleranza) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "6 - RIFERIMENTO MAX E = " + Test.Parameter(eRiferimentoMaxE) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "5 - RIFERIMENTO MAX R = " + Test.Parameter(eRiferimentoMaxR) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "4 - RIFERIMENTO NEUTRO = " + Test.Parameter(eRiferimentoNeutro) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "3 - SOURCE ADDRESS = " + Test.Parameter(eSourceAddress) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "2 - TEMPO CICLO = " + Test.Parameter(eTimerCiclicaSingoloCiclo) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "1 - TEMPO CICLICA = " + Test.Parameter(eTimerCiclica) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "  - TEST_CICLICA_CAN"
    Context.State = eNeutro
End Sub

Private Sub FloatTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA)
    Dim CE16Riferimento As Double
    Dim Tolleranza As Double

    If (Context.TimerSendMesssage < GetTickCount) Then
        Context.TimerSendMesssage = GetTickCount + 100
        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Retract, 250, Test.Parameter(eSourceAddress)
    End If
    If (Context.Timer < GetTickCount) Then
    
        ' Controllo la reale posizione della spola
        CE16Riferimento = CDbl(Test.Parameter(eRiferimentoMaxR))
        Tolleranza = CDbl(Test.Parameter(eTolleranza))
        
        
        
                ' Se il riferimento è inferiore ai 32 allora è espresso in mm e riferito al CE16
        If (CE16Riferimento <= 32) Then
            If (CE16(IndiceSezione).Position < CE16Riferimento + Tolleranza And CE16(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
                ' TUTTO OK
            Else
                Infos.RichTextBox1.Text = " ERRORE RETRACT " + Format(CStr(CE16(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
            End If
        Else
            If (MLT(IndiceSezione).Position < CE16Riferimento + Tolleranza And MLT(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
                ' TUTTO OK
            Else
                Infos.RichTextBox1.Text = " ERRORE RETRACT " + Format(CStr(MLT(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
            End If
        
        End If
                
                
                
        
        If (Context.bFirstCicloMaxR = True) Then
            Context.bFirstCicloMaxR = False
            Context.SensoreMLT_FirstMAXR = MLT(IndiceSezione).Position
            Context.SensoreMLT_MAXR = Context.SensoreMLT_FirstMAXR
            Infos.RichTextBox1.Text = " PRIMA LETTURA RETRACT " + vbLf + Infos.RichTextBox1.Text
        Else
            Context.SensoreMLT_MAXR = MLT(IndiceSezione).Position
        End If
        
        
        Context.State = eMaxE
        Context.Timer = GetTickCount + CDbl(Test.Parameter(eTimerCiclicaSingoloCiclo)) * 1000
    End If
    

        
End Sub

Private Sub NeutroTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA)
    Dim CE16Riferimento As Double
    Dim Tolleranza As Double

    If (Context.TimerSendMesssage < GetTickCount) Then
        Context.TimerSendMesssage = GetTickCount + 100
        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Neutral, 0, Test.Parameter(eSourceAddress)
    End If
    
    
    If (Context.Timer < GetTickCount) Then
        ' Controllo la reale posizione della spola
        CE16Riferimento = CDbl(Test.Parameter(eRiferimentoNeutro))
        Tolleranza = CDbl(Test.Parameter(eTolleranza))
        
        
        
        ' Se il riferimento è inferiore ai 32 allora è espresso in mm e riferito al CE16
        If (CE16Riferimento <= 32) Then
            If (CE16(IndiceSezione).Position < CE16Riferimento + Tolleranza And CE16(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
                ' TUTTO OK
            Else
                Infos.RichTextBox1.Text = " ERRORE NEUTRO " + Format(CStr(CE16(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
            End If
        Else
            
            ' Il riferimento è il modulo MLT
        
            If (MLT(IndiceSezione).Position < CE16Riferimento + Tolleranza And MLT(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
                ' TUTTO OK
            Else
                Infos.RichTextBox1.Text = " ERRORE NEUTRO " + Format(CStr(MLT(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
            End If
        End If
        
        
        
        ' Gestione del colore del richtextbox per l'errore
        If (Context.ErrorePosizioni > 0) Then
            If (Context.ErrorePosizioni = 1) Then
                Infos.RichTextBox1.BackColor = &H80FF&
            Else
                ' Ho letto un errore dalla ciclica
                If (Context.ErrorePosizioni = Context.OldErrorePosizioni) Then
                    ' Se l'errore = all'old vuol dire che durante questa sequenza non ho letto nessun errore dal modulo = COLORE ARANCIONE
                    'Infos.RichTextBox1.BackColor = &H80FF&
                    'Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "MODULO SBLOCCATO - CONTEGGIO ERRORE= " + CStr(Context.ErrorePosizioni) + vbLf
                Else
                    Infos.RichTextBox1.Text = "MODULO BLOCCATO " + CStr(Context.ErrorePosizioni) + vbLf + Infos.RichTextBox1.Text
                    Infos.RichTextBox1.BackColor = vbRed
                End If
            End If
        Else
            Infos.RichTextBox1.BackColor = vbWhite
        End If
        
        
        Context.DeltaMLT_MaxE = GetAbsVariazione(Context.SensoreMLT_FirstMaxE, Context.SensoreMLT_MaxE)
        Context.DeltaMLT_MaxR = GetAbsVariazione(Context.SensoreMLT_FirstMAXR, Context.SensoreMLT_MAXR)
     
'        If (Test.Parameter(eDeltaMassimoSensore) <> "") Then
'            If (Context.DeltaMLT_MaxE > CLng(Test.Parameter(eDeltaMassimoSensore))) Then
'                Infos.RichTextBox1.Text = "VARIAZIONE DEL MAGNETE ECCESSIVA LATO EXTEND " + vbLf + Infos.RichTextBox1.Text
'                Infos.RichTextBox1.BackColor = vbRed
'
'
'            ElseIf (Context.DeltaMLT_MaxE > CLng(Test.Parameter(eDeltaMassimoSensore))) Then
'                Infos.RichTextBox1.Text = "VARIAZIONE DEL MAGNETE ECCESSIVA LATO RETRACT " + vbLf + Infos.RichTextBox1.Text
'                Infos.RichTextBox1.BackColor = vbRed
'            End If
'
'        End If
        Infos.RichTextBox1.Text = "CICLO = " + CStr(Context.NumeroCicli) + "; DELTA E = " + CStr(Context.DeltaMLT_MaxE) + " ; DELTA R = " + CStr(Context.DeltaMLT_MaxR) + vbLf + Infos.RichTextBox1.Text
        
        Context.NumeroCicli = Context.NumeroCicli + 1
        
        
        ' Aggiorno la difu
        Context.OldErrorePosizioni = Context.ErrorePosizioni
        
        Context.State = eFloat
        Context.Timer = GetTickCount + CDbl(Test.Parameter(eTimerCiclicaSingoloCiclo)) * 1000
    End If
    
    If (Context.TimerCiclica < GetTickCount) Then
        Context.State = eEnd
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    End If
        
End Sub



Private Sub MaxeTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA)
    Dim CE16Riferimento As Double
    Dim Tolleranza As Double
    
    If (Context.TimerSendMesssage < GetTickCount) Then
        Context.TimerSendMesssage = GetTickCount + 100
        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Extend, 250, Test.Parameter(eSourceAddress)
    End If
    If (Context.Timer < GetTickCount) Then
        ' Controllo la reale posizione della spola
        CE16Riferimento = CDbl(Test.Parameter(eRiferimentoMaxE))
        Tolleranza = CDbl(Test.Parameter(eTolleranza))
        
        
        
        If (CE16Riferimento <= 32) Then
            If (CE16(IndiceSezione).Position < CE16Riferimento + Tolleranza And CE16(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
                ' TUTTO OK
            Else
                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " ERRORE EXTEND " + Format(CStr(CE16(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf
                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
            End If
        Else
            If (MLT(IndiceSezione).Position < CE16Riferimento + Tolleranza And MLT(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
                ' TUTTO OK
            Else
                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " ERRORE EXTEND " + Format(CStr(MLT(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf
                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
            End If
        End If
    
    
    
    
    
        If (Context.bFirstCicloMaxE = True) Then
            Context.bFirstCicloMaxE = False
            Context.SensoreMLT_FirstMaxE = MLT(IndiceSezione).Position
            Context.SensoreMLT_MaxE = Context.SensoreMLT_FirstMaxE
            Infos.RichTextBox1.Text = " PRIMA LETTURA EXTEND " + vbLf + Infos.RichTextBox1.Text
        Else
            Context.SensoreMLT_MaxE = MLT(IndiceSezione).Position
        End If
        
        Context.State = eNeutro
        Context.Timer = GetTickCount + CDbl(Test.Parameter(eTimerCiclicaSingoloCiclo)) * 1000
    End If
    

        
End Sub
'======================================
'FormattaCanAvcMessageForCanUsbDriver()
'
'COMPOSIZIONE DEL MESSAGGIO DI COMANDO DEL MODULO
'------------------------------------------------
Private Sub FormattaCanAvcMessageForCanUsbDriver(ByVal NodeId As Integer, ByVal State As MltState, ByVal Percentual As Integer, SourceAddSelez As String)
    
    On Error GoTo err:


    Dim strPercentual As String
    Dim strState As String
    Dim str As String
    
    'trasformo in Hex
    strState = Hex(State)
    strPercentual = Hex(Percentual)
    
    'devo sempre avere 2 cifre Hex
    Do While (Len(strPercentual) < 2)
        strPercentual = "0" + strPercentual
    Loop
    
          'T18FE3    NodeId del CH      default = 22           byte0          byte1    byte2 = stato     'byte 3-8
    'strVerificaCalibrazione = "T18FE3" & NodeIdSelez(Chn) & SourceAddSelez & "8" & strPercentual & "FF" & "F" & strState & "FFFFFFFFFF" & vbCr
    strVerificaCalibrazione = "T18FE3" & CStr(NodeId) & SourceAddSelez & "8" & strPercentual & "FF" & "F" & strState & "FFFFFFFFFF" & vbCr
    Output strVerificaCalibrazione

err:
End Sub 'END FormattaCanAvcMessageForCanUsbDriver()


Private Function GetAbsVariazione(ByVal PrimaLettura As Long, ByVal SecondaLettura As Long) As Long
    If (PrimaLettura > SecondaLettura) Then
        GetAbsVariazione = PrimaLettura - SecondaLettura
    Else
        GetAbsVariazione = SecondaLettura - PrimaLettura
    End If
End Function




Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Dim FileTest As String
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eTimerCiclica                    = " + Test.Parameter(eTimerCiclica) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eTimerCiclicaSingoloCiclo        = " + Test.Parameter(eTimerCiclicaSingoloCiclo) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eSourceAddress                   = " + Test.Parameter(eSourceAddress) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eRiferimentoNeutro               = " + Test.Parameter(eRiferimentoNeutro) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eRiferimentoMaxR                 = " + Test.Parameter(eRiferimentoMaxR) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eRiferimentoMaxE                 = " + Test.Parameter(eRiferimentoMaxE) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eTolleranza                      = " + Test.Parameter(eTolleranza) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
        
       
End Sub

