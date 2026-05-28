Attribute VB_Name = "Module_TEST_CICLICA_CAN_EVO"
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
    ePresetValue = 10
    eSave
    ePresetValue_wait
End Enum


Private Enum eTestParameter
    eNomeFile = 1                       ' Tempo totale della ciclica
End Enum

' struttura dati delle informazioni necessarie per la configurazione del file di uscita
Private Type TagTuplaFiles
    IdScheda As Integer
    TipoVariabile As String
    AddrVariabile As Integer
    NOME As String
End Type


Private Type TagPresetValue
     Tupla As TagTuplaFiles
     Value As Double
     Delay As Long
End Type

Public Type TagTEST_CICLICA_EVO
    bOldOn As Boolean               ' difu del comando di start del modulo
    Return As Integer                ' Stato del modulo
    State As Integer                    ' Stato macchina del modulo software
    OldState As Integer                 ' Old stato macchina del modulo software
    Timer As Long
    Timer_Timeout As Long
    ErrorePosizioni_N As Integer          ' Numero di errori di posizione della spola
    ErrorePosizioni_E As Integer          ' Numero di errori di posizione della spola
    ErrorePosizioni_R As Integer          ' Numero di errori di posizione della spola
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
    
    PresetValue() As TagPresetValue                 ' Variabili da settare prima di attivare la sequenza
    PresetIndex As Long
    
    StartTimer As Long
    
    
    Par_TimerCiclica As String
    Par_TimerCiclicaSingoloCiclo As String
    Par_SourceAddress As String
    Par_RiferimentoNeutro As String
    Par_RiferimentoMaxR As String
    Par_RiferimentoMaxE As String
    Par_Tolleranza As String
    Par_Timeout As String
    
    
    
    
    
    
    
    NodeCaptured As Integer
End Type

Private Context As TagTEST_CICLICA_EVO

Public Sub INFO_TEST_CICLICA_EVO(Infos As TagInfosSingleTest)
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










Private Sub Open_input_file(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)

    Dim FileTest As String
    Dim Righe As Variant
    Dim i As Long
    Dim NOME As String
    Dim Valore As String
    
    FileTest = FileText(Get_Nome_Assoluto_File(Test.Parameter(eNomeFile)))
    Righe = Split(FileTest, vbLf)
        
    ReDim Context.PresetValue(0)
    
    Context.Par_Timeout = 0
    
    i = 0
    Do While (i <= UBound(Righe))
    
        Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
        Get_Coppia_Variabile_Valore Righe(i), NOME, Valore

        Select Case NOME
        Case "TIMER_CICLICA": Context.Par_TimerCiclica = Valore
        Case "TIMER_CICLICA_SINGOLO_CICLO": Context.Par_TimerCiclicaSingoloCiclo = Valore
        
        Case "SOURCE_ADDRESS": Context.Par_SourceAddress = Valore
        Case "RIFERIMENTO_NEUTRO": Context.Par_RiferimentoNeutro = Valore
        Case "RIFERIMENTO_MAXR": Context.Par_RiferimentoMaxR = Valore
        Case "RIFERIMENTO_MAXE": Context.Par_RiferimentoMaxE = Valore
        Case "TOLLERANZA": Context.Par_Tolleranza = Valore
        Case "TIMEOUT": Context.Par_Timeout = Valore
        
        Case "PRESET_VARIABLE":
            UpdatePresetValue Context.PresetValue(UBound(Context.PresetValue)), Valore, IndiceSezione
            ReDim Preserve Context.PresetValue(UBound(Context.PresetValue) + 1)
            
        
        End Select
        i = i + 1

    Loop
        
        

End Sub














Private Sub TEST_CICLICA_EVO_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_CICLICA_EVO)
    Context1.State = eEnd
    
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = "CICLICA EVO"
    End If
End Sub


Public Sub TEST_CICLICA_EVO(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_CICLICA_EVO, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "TEST_CICLICA_CAN_EVO") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
        bReturn = Context1.Return
        INFO_TEST_CICLICA_EVO Infos
        
        If (bStop = True And bOn = True) Then
            
            TEST_CICLICA_EVO_INIT_STRUCT Infos, Context1
            
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            
            Infos.RichTextBox1 = Infos.RichTextBox1 + "STOP" + vbLf
            
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context1.State = eError
        ElseIf (bOn = True And Context1.bOldOn = False) Then          ' fronte di salita del test ( AVVIO )
        
            Context1.StartTimer = GetTickCountEvo
            Init Test, Infos, IndiceSezione, Context1                                         ' init delle variabili del modulo
            Infos.RichTextBox1 = Infos.RichTextBox1 + "START" + vbLf
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
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    Dim i As Integer


    Open_input_file Test, Infos, IndiceSezione, Context
    
    Context.Return = eTEST_BUSY_NON_BLOCCANTE
    
    Context.OldState = 0
    Context.State = eSpegniModuli
    Context.Timer = GetTickCountEvo + SPEGMINENTO_TIMEOUT_MS
    
    Do While (i < MAX_SEZIONI)
        MLTCANRelais(i) = True
        i = i + 1
    Loop
    'Infos.RichTextBox1.Text = ""
    Context.bFirstCicloMaxE = True
    Context.bFirstCicloMaxR = True
    Context.SensoreMLT_MaxE = 0
    Context.SensoreMLT_MAXR = 0
    Context.SensoreMLT_FirstMaxE = 0
    Context.SensoreMLT_FirstMAXR = 0
    Context.NumeroCicli = 0
    Context.ErrorePosizioni_N = 0                                     ' Resetto gli errori di posizione
    Context.ErrorePosizioni_E = 0                                     ' Resetto gli errori di posizione
    Context.ErrorePosizioni_R = 0                                     ' Resetto gli errori di posizione
    Context.OldErrorePosizioni = 0
    Infos.RichTextBox1.BackColor = vbWhite
    
    Prepare_Textbox_Parametri Test, Infos
    
End Sub

' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    Select Case Context.State
    Case eInit: Init Test, Infos, IndiceSezione, Context
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, Context
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, Context
    Case eCapture: Capture Test, Infos, IndiceSezione, Context
    Case ePresetValue: PresetValue Test, Infos, IndiceSezione, Context
    Case ePresetValue_wait: PresetValue_wait Test, Infos, IndiceSezione, Context
    Case eRun: Run Test, Infos, IndiceSezione, Context
    Case eNeutro: NeutroTest Test, Infos, IndiceSezione, Context
    Case eFloat:  FloatTest Test, Infos, IndiceSezione, Context
    Case eMaxE:   MaxeTest Test, Infos, IndiceSezione, Context
    Case eEnd: Fine Test, Infos, IndiceSezione, Context
    Case eSave: Save Test, Infos, IndiceSezione, Context
    Case eError: ErrorTest Test, Infos, IndiceSezione, Context
    Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
    If (Context.OldState <> Context.State) Then
        'Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "STATE = " + CStr(Context.State) + vbLf
    End If
    
    Context.OldState = Context.State
End Sub

Private Sub Save(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
End Sub

Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    If (Infos.RichTextBox1.BackColor <> vbRed) Then
        Context.Return = eTEST_OK
    Else
        Context.Return = eTEST_ERROR
    End If
End Sub


Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    
    If (Infos.RichTextBox1.BackColor <> vbRed) Then
        Infos.RichTextBox1.BackColor = vbRed
    End If
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = "ERRORE SUL TEST" + vbLf + Infos.RichTextBox1.Text
    End If
    Context.Return = eTEST_ERROR
End Sub

Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    TurnOffModule IndiceSezione
    
    If (Context.Timer < GetTickCountEvo) Then
        Context.State = eAccendiModulo
        Context.Timer = GetTickCountEvo + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCountEvo) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    
    Dim i As Long
    
    If (NodeCaptured(IndiceSezione) = -1) Then
    
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eError
    Else
        
    End If
    
    NodeCaptured(IndiceSezione) = IndiceSezione + 1
    Context.NodeCaptured = NodeCaptured(IndiceSezione)
    
    Infos.RichTextBox1.Text = "ID = " + CStr(NodeCaptured(IndiceSezione)) + vbLf + Infos.RichTextBox1.Text
    
    

    Context.Timer = GetTickCountEvo
    Context.State = ePresetValue_wait
    Context.PresetIndex = 0
    Context.Timer = GetTickCountEvo + Context.PresetValue(Context.PresetIndex).Delay
    
End Sub


Private Sub PresetValue(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    
    
    Infos.RichTextBox1.Text = CStr(Context.PresetValue(Context.PresetIndex).Tupla.IdScheda) + "." + Context.PresetValue(Context.PresetIndex).Tupla.TipoVariabile + "." + CStr(Context.PresetValue(Context.PresetIndex).Tupla.AddrVariabile) + " = " + CStr(Context.PresetValue(Context.PresetIndex).Value) + vbLf + Infos.RichTextBox1.Text
    SetCalibratedValue Context.PresetValue(Context.PresetIndex).Tupla.IdScheda, Context.PresetValue(Context.PresetIndex).Tupla.TipoVariabile, Context.PresetValue(Context.PresetIndex).Tupla.AddrVariabile, Context.PresetValue(Context.PresetIndex).Value
    
    Context.PresetIndex = Context.PresetIndex + 1
    
    If (Context.PresetIndex > UBound(Context.PresetValue)) Then
        Context.State = eRun
    Else
        Context.State = ePresetValue_wait
        Context.Timer = GetTickCountEvo + Context.PresetValue(Context.PresetIndex).Delay
    End If
    
    
    'Context.Timer = GetTickCountEvo + Context.PresetValue(Context.PresetIndex).Delay
    

    
    
        
        
    'Context.State = ePresetValue_wait

    


End Sub


Private Sub PresetValue_wait(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    
    
    If (Context.Timer < GetTickCountEvo) Then
        Context.State = ePresetValue
'        If (Context.PresetIndex > UBound(Context.PresetValue)) Then
'            Context.State = eStart
'        Else
'            Context.State = ePresetValue
'        End If
    End If
        
        
    


End Sub




Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
On Error GoTo err:
    Context.Timer = GetTickCountEvo + CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
    Context.TimerSendMesssage = 0
    Context.TimerCiclica = GetTickCountEvo + CDbl(Context.Par_TimerCiclica) * 1000
    Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout) * 1000

    Infos.RichTextBox1.Text = "7 - TOLLERANZA = " + Context.Par_Tolleranza + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "6 - RIFERIMENTO MAX E = " + Context.Par_RiferimentoMaxE + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "5 - RIFERIMENTO MAX R = " + Context.Par_RiferimentoMaxR + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "4 - RIFERIMENTO NEUTRO = " + Context.Par_RiferimentoNeutro + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "3 - SOURCE ADDRESS = " + Context.Par_SourceAddress + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "2 - TEMPO CICLO = " + Context.Par_TimerCiclicaSingoloCiclo + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "1 - TEMPO CICLICA = " + Context.Par_TimerCiclica + vbLf + Infos.RichTextBox1.Text
    Context.State = eNeutro
    Exit Sub
err:
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-----------------------------" + vbLf + "ERROR Run =" + err.Description
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eError
End Sub




Private Sub UpdatePresetValue(ByRef PresetValue As TagPresetValue, Valore As String, IndiceSezione As Integer)
    Dim Ritardo As Variant
    Dim Dati As Variant
    Dim tupa As Variant
    'Valore = Rimuovi_Commento_Da_Stringa(Valore, "//")      ' Tolgo i commenti
    Valore = Replace(Valore, "(", "")
    Valore = Replace(Valore, ")", "")
    Valore = Replace(Valore, " ", "")
    
    
    PresetValue.Delay = CLng(GetDelayValue(PresetValue, Valore, IndiceSezione))
    
    
    Dati = Split(Valore, "<-")
    Dati(0) = Trim(Dati(0))
    Dati(1) = Trim(Dati(1))

    tupa = Split(Dati(0), ".")
    PresetValue.Tupla.IdScheda = CInt(tupa(0))
    
    If (PresetValue.Tupla.IdScheda = -1) Then
        PresetValue.Tupla.IdScheda = IndiceSezione + 1
    End If
    
    
    PresetValue.Tupla.TipoVariabile = tupa(1)
    PresetValue.Tupla.AddrVariabile = CInt(tupa(2))
    
    If (PresetValue.Tupla.AddrVariabile = -1) Then
        PresetValue.Tupla.AddrVariabile = IndiceSezione
    End If
    
    
    'PresetValue.Delay = CLng(GetDelayValue(PresetValue, Valore, IndiceSezione))
    PresetValue.Value = CDbl(Dati(1))
    

End Sub




Private Function GetDelayValue(ByRef PresetValue As TagPresetValue, Valore As String, IndiceSezione As Integer) As String
    On Error GoTo err:
    Dim s As String
    s = Valore
    Dim leftBracket As Integer
    Dim rightBracket As Integer
    leftBracket = InStr(s, "[")
    rightBracket = InStr(s, "]")
    Dim result As String
    result = Mid(s, leftBracket + 1, rightBracket - leftBracket - 1)
    
    Valore = Replace(Valore, "[" + result + "]", "")
    GetDelayValue = result
    Exit Function
err:
    GetDelayValue = "0"
    
End Function

















Private Sub FloatTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    
    
    
  
    Dim CE16Riferimento As Double
    Dim Tolleranza As Double
    Dim Position As Double
    Dim bOk As Boolean
    Dim bExit As Boolean
    
    ' Trasmissione del messaggio ciclico
    If (Context.TimerSendMesssage < GetTickCountEvo) Then
        Context.TimerSendMesssage = GetTickCountEvo + 50
        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Retract, 250, Context.Par_SourceAddress
    End If
    
    ' Tempo minimo di attesa
    If (Context.Timer < GetTickCountEvo) Then
        ' Controllo la reale posizione della spola
        CE16Riferimento = CDbl(Context.Par_RiferimentoMaxR)
        Tolleranza = CDbl(Context.Par_Tolleranza)
        
        If (CE16Riferimento <= 32) Then
            If (CE16(IndiceSezione).Position = "") Then
                bExit = True
                Infos.RichTextBox1.Text = " ERRORE MODULO CE16 NON PRESENTE" + vbCrLf + Infos.RichTextBox1.Text
            Else
                Position = CDbl(CE16(IndiceSezione).Position)
            End If
        Else
            If (MLT(IndiceSezione).Position = "") Then
                bExit = True
                Infos.RichTextBox1.Text = " ERRORE MODULO MLT NON PRESENTE" + vbCrLf + Infos.RichTextBox1.Text
            Else
                Position = CDbl(MLT(IndiceSezione).Position)
            End If
        End If
        
        
        ' controllo se non ho avuto nessun errore
        If (bExit = False) Then
            
            If (Position < CE16Riferimento + Tolleranza And Position > CE16Riferimento - Tolleranza) Then
                bOk = True
            End If
            
        End If
        
    End If
        
    
    If (bOk = True) Then
        
        bExit = True
        Context.DeltaMLT_MaxE = GetAbsVariazione(Context.SensoreMLT_FirstMaxE, Context.SensoreMLT_MaxE)
        Context.DeltaMLT_MaxR = GetAbsVariazione(Context.SensoreMLT_FirstMAXR, Context.SensoreMLT_MAXR)

    
    Else
        ' Controllo se sono andato in timeout!!!!
        If (Context.Timer_Timeout < GetTickCountEvo And Context.Timer < GetTickCountEvo) Then
            bExit = True
        End If
    End If
    
        
    
    If (bExit = True) Then
        If (bOk = False) Then
            Context.ErrorePosizioni_R = Context.ErrorePosizioni_R + 1
        End If
        
        
        'Infos.RichTextBox1.Text = "CICLO = " + CStr(Context.NumeroCicli) + vbLf + Infos.RichTextBox1.Text  '+ "; DELTA E = " + CStr(Context.DeltaMLT_MaxE) + " ; DELTA R = " + CStr(Context.DeltaMLT_MaxR) + vbLf + Infos.RichTextBox1.Text
        
        
        
        
        ' Aggiorno la difu
        'Context.OldErrorePosizioni = Context.ErrorePosizioni
        
        Context.State = eMaxE
        Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout) * 1000
        Context.TimerSendMesssage = 0
        Context.Timer = GetTickCountEvo + CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
    Else
    
    End If
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'    Dim CE16Riferimento As Double
'    Dim Tolleranza As Double
'
'    If (Context.TimerSendMesssage < GetTickCountEvo) Then
'        Context.TimerSendMesssage = GetTickCountEvo + 50
'        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Retract, 250, Context.Par_SourceAddress
'    End If
'    If (Context.Timer < GetTickCountEvo) Then
'
'        ' Controllo la reale posizione della spola
'        CE16Riferimento = CDbl(Context.Par_RiferimentoMaxR)
'        Tolleranza = CDbl(Context.Par_Tolleranza)
'
'
'
'                ' Se il riferimento è inferiore ai 32 allora è espresso in mm e riferito al CE16
'        If (CE16Riferimento <= 32) Then
'            If (CE16(IndiceSezione).Position < CE16Riferimento + Tolleranza And CE16(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
'                ' TUTTO OK
'            Else
'                Infos.RichTextBox1.Text = " ERRORE RETRACT " + Format(CStr(CE16(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'            End If
'        Else
'            If (MLT(IndiceSezione).Position = "") Then
'                Infos.RichTextBox1.Text = " ERRORE MODULO CAN NON PRESENTE"
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'
'            ElseIf (MLT(IndiceSezione).Position < CE16Riferimento + Tolleranza And MLT(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
'                ' TUTTO OK
'            Else
'                Infos.RichTextBox1.Text = " ERRORE RETRACT " + Format(CStr(MLT(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'            End If
'
'        End If
'
'
'
'
'        If (Context.bFirstCicloMaxR = True) Then
'            Context.bFirstCicloMaxR = False
'            Context.SensoreMLT_FirstMAXR = MLT(IndiceSezione).Position
'            Context.SensoreMLT_MAXR = Context.SensoreMLT_FirstMAXR
'            Infos.RichTextBox1.Text = " PRIMA LETTURA RETRACT " + CStr(MLT(IndiceSezione).Position) + vbLf + Infos.RichTextBox1.Text
'        Else
'            If (MLT(IndiceSezione).Position <> "") Then
'                Context.SensoreMLT_MAXR = MLT(IndiceSezione).Position
'            End If
'        End If
'
'
'        Context.State = eMaxE
'        Context.TimerSendMesssage = 0
'        Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout)
'        Context.Timer = GetTickCountEvo + 2 * CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
'    End If
'

        
End Sub

Private Sub NeutroTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    Dim CE16Riferimento As Double
    Dim Tolleranza As Double
    Dim Position As Double
    Dim bOk As Boolean
    Dim bExit As Boolean
    Dim strx As String
    Dim Valuex As Double
    
    
    ' Trasmissione del messaggio ciclico
    If (Context.TimerSendMesssage < GetTickCountEvo) Then
        Context.TimerSendMesssage = GetTickCountEvo + 50
        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Neutral, 0, Context.Par_SourceAddress
    End If
    
    ' Tempo minimo di attesa
    If (Context.Timer < GetTickCountEvo) Then
        ' Controllo la reale posizione della spola
        CE16Riferimento = CDbl(Context.Par_RiferimentoNeutro)
        Tolleranza = CDbl(Context.Par_Tolleranza)
        
        If (CE16Riferimento <= 32) Then
            If (CE16(IndiceSezione).Position = "") Then
                bExit = True
                Infos.RichTextBox1.Text = " ERRORE MODULO CE16 NON PRESENTE " + vbCrLf + Infos.RichTextBox1.Text
            Else
                Position = CDbl(CE16(IndiceSezione).Position)
            End If
        Else
            If (MLT(IndiceSezione).Position = "") Then
                bExit = True
                Infos.RichTextBox1.Text = " ERRORE MODULO MLT NON PRESENTE " + vbCrLf + Infos.RichTextBox1.Text
            Else
                Position = CDbl(MLT(IndiceSezione).Position)
            End If
        End If
        
        
        ' controllo se non ho avuto nessun errore
        If (bExit = False) Then
            
            If (Position < CE16Riferimento + Tolleranza And Position > CE16Riferimento - Tolleranza) Then
                bOk = True
            End If
            
        End If
        
    End If
        
    
    If (bOk = True) Then
        
        bExit = True
        Context.DeltaMLT_MaxE = GetAbsVariazione(Context.SensoreMLT_FirstMaxE, Context.SensoreMLT_MaxE)
        Context.DeltaMLT_MaxR = GetAbsVariazione(Context.SensoreMLT_FirstMAXR, Context.SensoreMLT_MAXR)

    
    Else
        ' Controllo se sono andato in timeout!!!!
        If (Context.Timer_Timeout < GetTickCountEvo And Context.Timer < GetTickCountEvo) Then
            bExit = True
        End If
    End If
    
        
    
    If (bExit = True) Then
        If (bOk = False) Then
            Context.ErrorePosizioni_N = Context.ErrorePosizioni_N + 1
            
            
        End If
        
        If (Context.ErrorePosizioni_N = 0 And Context.ErrorePosizioni_E = 0 And Context.ErrorePosizioni_R = 0) Then
            Infos.RichTextBox1.BackColor = vbWhite
        ElseIf (Context.ErrorePosizioni_N <= 1 And Context.ErrorePosizioni_R <= 1 And Context.ErrorePosizioni_E <= 1) Then
            Infos.RichTextBox1.BackColor = &H80FF&
        Else
            Infos.RichTextBox1.BackColor = vbRed
        End If
        
        Valuex = CDbl(Context.TimerCiclica - GetTickCountEvo) / 1000
        strx = CStr(Valuex)
        strx = Format(strx, "0000.0")
        
        
        Infos.RichTextBox1.Text = strx + "sec " + " CICLO=" + CStr(Context.NumeroCicli) + _
        " ERROR N=" + CStr(Context.ErrorePosizioni_N) + _
        " ERROR E=" + CStr(Context.ErrorePosizioni_E) + _
        " ERROR R=" + CStr(Context.ErrorePosizioni_R) + _
        vbLf
        
        Context.NumeroCicli = Context.NumeroCicli + 1
        
        
        ' Aggiorno la difu
        'Context.OldErrorePosizioni = Context.ErrorePosizioni
        
        Context.State = eFloat
        Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout) * 1000
        Context.TimerSendMesssage = 0
        Context.Timer = GetTickCountEvo + CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
    Else
        If (Context.TimerCiclica < GetTickCountEvo) Then
            Context.State = eSave
        End If
    
    End If
        
        
        
        
        
        
        
        
        
        
        
        
        
'
'        ' Se il riferimento è inferiore ai 32 allora è espresso in mm e riferito al CE16
'        If (CE16Riferimento <= 32) Then
'            If (CE16(IndiceSezione).Position = "") Then
'
'                Infos.RichTextBox1.Text = " ERRORE MODULO CE16 NON PRESENTE"
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'                bExit = True
'
'            ElseIf (CE16(IndiceSezione).Position < CE16Riferimento + Tolleranza And CE16(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
'                ' TUTTO OK
'                bOk = True
'            Else
'                If (Context.Timer_Timeout < GetTickCountEvo) Then
'                    Infos.RichTextBox1.Text = " ERRORE NEUTRO " + Format(CStr(CE16(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
'                    Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'                    bExit = True
'                End If
'            End If
'        Else
'
'            ' Il riferimento è il modulo MLT
'            If (MLT(IndiceSezione).Position = "") Then
'                Infos.RichTextBox1.Text = " ERRORE MODULO CAN NON PRESENTE"
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'                bExit = True
'
'            ElseIf (MLT(IndiceSezione).Position < CE16Riferimento + Tolleranza And MLT(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
'                ' TUTTO OK
'                bOk = True
'            Else
'                If (Context.Timer_Timeout < GetTickCountEvo) Then
'                    Infos.RichTextBox1.Text = " ERRORE NEUTRO " + Format(CStr(MLT(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf + Infos.RichTextBox1.Text
'                    Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'                    bExit = True
'                End If
'            End If
'        End If
'
'
'
'        ' Gestione del colore del richtextbox per l'errore
'        If (Context.ErrorePosizioni > 0) Then
'            If (Context.ErrorePosizioni = 1) Then
'                Infos.RichTextBox1.BackColor = &H80FF&
'            Else
'                ' Ho letto un errore dalla ciclica
'                If (Context.ErrorePosizioni = Context.OldErrorePosizioni) Then
'                    ' Se l'errore = all'old vuol dire che durante questa sequenza non ho letto nessun errore dal modulo = COLORE ARANCIONE
'                    'Infos.RichTextBox1.BackColor = &H80FF&
'                    'Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "MODULO SBLOCCATO - CONTEGGIO ERRORE= " + CStr(Context.ErrorePosizioni) + vbLf
'                Else
'                    Infos.RichTextBox1.Text = "MODULO BLOCCATO " + CStr(Context.ErrorePosizioni) + vbLf + Infos.RichTextBox1.Text
'                    Infos.RichTextBox1.BackColor = vbRed
'                End If
'            End If
'        Else
'            Infos.RichTextBox1.BackColor = vbWhite
'        End If
'
'
'        Context.DeltaMLT_MaxE = GetAbsVariazione(Context.SensoreMLT_FirstMaxE, Context.SensoreMLT_MaxE)
'        Context.DeltaMLT_MaxR = GetAbsVariazione(Context.SensoreMLT_FirstMAXR, Context.SensoreMLT_MAXR)
'
'
'
'
'        Infos.RichTextBox1.Text = "CICLO = " + CStr(Context.NumeroCicli) + "; DELTA E = " + CStr(Context.DeltaMLT_MaxE) + " ; DELTA R = " + CStr(Context.DeltaMLT_MaxR) + vbLf + Infos.RichTextBox1.Text
'
'        Context.NumeroCicli = Context.NumeroCicli + 1
'
'
'        ' Aggiorno la difu
'        Context.OldErrorePosizioni = Context.ErrorePosizioni
'
'        Context.State = eFloat
'        Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout)
'        Context.TimerSendMesssage = 0
'        Context.Timer = GetTickCountEvo + CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
'    End If
'
'    If (Context.TimerCiclica < GetTickCountEvo) Then
'        Context.State = eEnd
'    End If
        
End Sub



Private Sub MaxeTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_CICLICA_EVO)
    
    
    Dim CE16Riferimento As Double
    Dim Tolleranza As Double
    Dim Position As Double
    Dim Time As Long
    
    
    Dim bOk As Boolean
    Dim bExit As Boolean
    
    
    Time = GetTickCountEvo
    
    ' Trasmissione del messaggio ciclico
    If (Context.TimerSendMesssage < GetTickCountEvo) Then
        Context.TimerSendMesssage = GetTickCountEvo + 50
        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Extend, 250, Context.Par_SourceAddress
    End If
    
    ' Tempo minimo di attesa
    If (Context.Timer < GetTickCountEvo) Then
        ' Controllo la reale posizione della spola
        CE16Riferimento = CDbl(Context.Par_RiferimentoMaxE)
        Tolleranza = CDbl(Context.Par_Tolleranza)
        
        If (CE16Riferimento <= 32) Then
            If (CE16(IndiceSezione).Position = "") Then
                bExit = True
                Infos.RichTextBox1.Text = " ERRORE MODULO CE16 NON PRESENTE" + vbCrLf + Infos.RichTextBox1.Text
            Else
                Position = CDbl(CE16(IndiceSezione).Position)
            End If
        Else
            If (MLT(IndiceSezione).Position = "") Then
                bExit = True
                Infos.RichTextBox1.Text = " ERRORE MODULO MLT NON PRESENTE" + vbCrLf + Infos.RichTextBox1.Text
            Else
                Position = CDbl(MLT(IndiceSezione).Position)
            End If
        End If
        
        
        ' controllo se non ho avuto nessun errore
        If (bExit = False) Then
            
            If (Position < CE16Riferimento + Tolleranza And Position > CE16Riferimento - Tolleranza) Then
                bOk = True
            End If
            
        End If
        
    End If
        
    
    If (bOk = True) Then
        
        bExit = True
        Context.DeltaMLT_MaxE = GetAbsVariazione(Context.SensoreMLT_FirstMaxE, Context.SensoreMLT_MaxE)
        Context.DeltaMLT_MaxR = GetAbsVariazione(Context.SensoreMLT_FirstMAXR, Context.SensoreMLT_MAXR)

    
    Else
        ' Controllo se sono andato in timeout!!!!
        If (Context.Timer_Timeout < Time And Context.Timer < Time) Then
            bExit = True
        End If
    End If
    
        
    
    If (bExit = True) Then
        If (bOk = False) Then
            Context.ErrorePosizioni_E = Context.ErrorePosizioni_E + 1
            'Infos.RichTextBox1.Text = "MODULO BLOCCATO " + CStr(Context.ErrorePosizioni) + vbLf + Infos.RichTextBox1.Text
        End If
    
        
        
        
        
        
        Context.State = eNeutro
        Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout) * 1000
        Context.TimerSendMesssage = 0
        Context.Timer = GetTickCountEvo + CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
    Else
        If (Context.TimerCiclica < GetTickCountEvo) Then
            Context.State = eSave
        End If
    
    End If
            
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
'
'
'
'
'
'
'    Dim CE16Riferimento As Double
'    Dim Tolleranza As Double
'
'    If (Context.TimerSendMesssage < GetTickCountEvo) Then
'        Context.TimerSendMesssage = GetTickCountEvo + 50
'        FormattaCanAvcMessageForCanUsbDriver Context.NodeCaptured, Extend, 250, Context.Par_SourceAddress
'
'    ElseIf (Context.Timer < GetTickCountEvo) Then
'        ' Controllo la reale posizione della spola
'        CE16Riferimento = CDbl(Context.Par_RiferimentoMaxE)
'        Tolleranza = CDbl(Context.Par_Tolleranza)
'
'
'
'        If (CE16Riferimento <= 32) Then
'            If (CE16(IndiceSezione).Position = "") Then
'                Infos.RichTextBox1.Text = " ERRORE CS16 MODULO NON PRESENTE"
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'
'            ElseIf (CE16(IndiceSezione).Position < CE16Riferimento + Tolleranza And CE16(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
'                ' TUTTO OK
'            Else
'                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " ERRORE EXTEND " + Format(CStr(CE16(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'            End If
'        Else
'            If (MLT(IndiceSezione).Position = "") Then
'
'                Infos.RichTextBox1.Text = " ERRORE MODULO CAN NON PRESENTE"
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'
'            ElseIf (MLT(IndiceSezione).Position < CE16Riferimento + Tolleranza And MLT(IndiceSezione).Position > CE16Riferimento - Tolleranza) Then
'                ' TUTTO OK
'            Else
'                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " ERRORE EXTEND " + Format(CStr(MLT(IndiceSezione).Position), "0.##") + " - " + Format(CStr(CE16Riferimento), "0.##") + CStr(Tolleranza) + vbLf
'                Context.ErrorePosizioni = Context.ErrorePosizioni + 1
'            End If
'        End If
'
'
'
'
'
'        If (Context.bFirstCicloMaxE = True) Then
'            Context.bFirstCicloMaxE = False
'
'            If (MLT(IndiceSezione).Position <> "") Then
'                Context.SensoreMLT_FirstMaxE = MLT(IndiceSezione).Position
'            End If
'
'            Context.SensoreMLT_MaxE = Context.SensoreMLT_FirstMaxE
'            Infos.RichTextBox1.Text = " PRIMA LETTURA EXTEND " + CStr(MLT(IndiceSezione).Position) + vbLf + Infos.RichTextBox1.Text
'        Else
'
'            If (MLT(IndiceSezione).Position <> "") Then
'                Context.SensoreMLT_MaxE = MLT(IndiceSezione).Position
'            End If
'
'        End If
'
'        Context.State = eNeutro
'        Context.TimerSendMesssage = 0
'        Context.Timer_Timeout = GetTickCountEvo + CDbl(Context.Par_Timeout)
'        Context.Timer = GetTickCountEvo + CDbl(Context.Par_TimerCiclicaSingoloCiclo) * 1000
'    End If
'

        
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





Private Function FileText(FileName As String) As String
    On Error GoTo err:
    Dim handle As Integer
    handle = FreeFile
    
    Open FileName For Input As #handle
    
    FileText = Input$(LOF(handle), handle)
    
    
    Close #handle
    Exit Function
err:
    MsgBox "File di configurazione delle variabili da acquisire errato ", vbCritical
End Function


Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Dim FileTest As String
    FileTest = FileText(Get_Nome_Assoluto_File(Test.Parameter(eNomeFile)))
    
        Infos.RichTextBox1.Text = "-------------------------------" + vbCrLf + _
                                  "LISTA PARAMETRI" + vbCrLf + _
                                  "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFile                    = " + Test.Parameter(eNomeFile) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + FileTest + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
        
       
    

End Sub


