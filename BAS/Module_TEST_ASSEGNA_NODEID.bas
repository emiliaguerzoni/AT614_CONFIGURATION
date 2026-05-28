Attribute VB_Name = "Module_TEST_ASSEGNA_NODEID"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 300
Private Const ACCENSIONE_TIMEOUT_MS = 350
'eTEST_OK = 2
'eTEST_BUSY = 1
'eTEST_STOP = 0
'eTEST_ERROR = -1



' Stato macchina del modulo
Private Enum eMachineState
    eInit
    eInitRepeat
    eSpegniModuli
    eAccendiModulo
    eCapture
    eRun
    eSave
    eSaveError
    eEnd
    eError
    eSpegniModuli2
    eAccendiModulo2
    eCapture2
End Enum


Private Enum eTestParameter
    eNomeFile = 1
End Enum


Public Type TagTEST_ASSEGNA_NODEID
    bOldOn As Boolean               ' difu del comando di start del modulo
    Return As Integer                ' Stato del modulo
    State As Integer
    Timer As Long
    StartTimer As Long
    Repeat As Long
    ErrorString As String
    
End Type

Private Context As TagTEST_ASSEGNA_NODEID

Public Sub INFO_ASSEGNA_NODEID(Infos As TagInfosSingleTest)
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

Private Sub TEST_ASSEGNA_NODEID_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_ASSEGNA_NODEID)
    Context1.State = eEnd
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = ""
    End If
End Sub


Public Sub TEST_ASSEGNA_NODEID(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_ASSEGNA_NODEID, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
    If (bStop = True) Then
        Debug.Print "SUKS"
    End If
    If (Test.ID <> "ASSEGNA_NODEID") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
        If (bStop = True And bOn = True) Then
        
            TEST_ASSEGNA_NODEID_INIT_STRUCT Infos, Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            
            Context1.ErrorString = "STOP DA PARTE DELL'UTENTE"
            Context1.State = eSaveError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then          ' fronte di salita del test ( AVVIO )
            Context1.StartTimer = GetTickCountEvo
            Init Test, Infos, IndiceSezione, Context1                                         ' init delle variabili del modulo
        End If
        
        INFO_ASSEGNA_NODEID Infos
        
        Task Test, Infos, IndiceSezione, Context1                                       ' Task del modulo
'        If (bOn = False) Then
'            Debug.Print "ERR"
'        End If
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        
        bReturn = Context1.Return
    End If
    
End Sub

' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Context.Return = eTEST_BUSY
    Context.State = eSpegniModuli
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.Repeat = 0
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = "ASSEGNA NODE ID"
    End If
    
    If (Infos.RichTextBox1.BackColor <> vbWhite) Then
        Infos.RichTextBox1.BackColor = vbWhite
    End If
    
    Prepare_Textbox_Parametri Test, Infos
    
    Infos.RichTextBox1.Text = CStr(CStr(Now)) & "-" + "INIT" + vbLf + Infos.RichTextBox1.Text
    
End Sub

' Questa funzione inizializza la memoria del frame
Private Sub InitRepeat(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Context.Return = eTEST_BUSY
    Context.State = eSpegniModuli
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.Repeat = Context.Repeat + 1
    
    If (Infos.RichTextBox1.BackColor <> vbWhite) Then
        Infos.RichTextBox1.BackColor = vbWhite
    End If
    
    
    If (Context.Repeat >= 5) Then
        Context.ErrorString = "NUMERO MASSIMO DI TENTATIVI RAGGIUNTO"
        Context.State = eSaveError
    Else
        Context.State = eSpegniModuli
    End If
    
End Sub




' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Select Case Context.State
    Case eInit: Init Test, Infos, IndiceSezione, Context
    Case eInitRepeat: InitRepeat Test, Infos, IndiceSezione, Context
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, Context
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, Context
    Case eCapture: Capture Test, Infos, IndiceSezione, Context
    Case eRun: Run Test, Infos, IndiceSezione, Context
    Case eSave: Save Test, Infos, IndiceSezione, Context
    Case eSaveError: SaveError Test, Infos, IndiceSezione, Context
    Case eSpegniModuli2: SpegniModuli2 Test, Infos, IndiceSezione, Context
    Case eAccendiModulo2: AccendiModulo2 Test, Infos, IndiceSezione, Context
    Case eCapture2: Capture2 Test, Infos, IndiceSezione, Context
    Case eEnd: Context.Return = eTEST_OK
    Case eError: ErrorTest Test, Infos, IndiceSezione, Context
    Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
End Sub


Private Sub SaveError(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Context.State = eError
End Sub


Private Sub Save(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
End Sub

Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    
    Infos.RichTextBox1.BackColor = vbRed
    
    If (Context.Return <> eTEST_ERROR) Then
        
        Infos.RichTextBox1 = " ------------------------------------ " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = Context.ErrorString + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = " ------------------------------------ " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = " ! ! ! ! ! ERRORE SUL TEST ! ! ! ! ! " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = " ------------------------------------ " + vbLf + Infos.RichTextBox1.Text
        
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    End If
    Context.Return = eTEST_ERROR

End Sub

Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    TurnOffAllModule
    
    If (Context.Timer < GetTickCount) Then
        Infos.RichTextBox1.Text = CStr(CStr(Now)) & "-" + "SPENGO I MODULI" + vbLf + Infos.RichTextBox1.Text
        Context.State = eAccendiModulo
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub



Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Dim str As String
'    If (NodeCaptured(IndiceSezione) = -1) Then
'        NodeCaptured(IndiceSezione) = CatturaMltPlus()
'    Else
'        NodeCaptured(IndiceSezione) = NodeCaptured(IndiceSezione)
'    End If
    
    NodeCaptured(IndiceSezione) = CatturaMltPlus(, str)
    Infos.Label1.Caption = str
    
    
    
    
    If (NodeCaptured(IndiceSezione) = -1) Then
        Infos.RichTextBox1.Text = CStr(Now) & "-" + "CATTURO I MODULI FAILURE [" + CStr(Context.Repeat) + "]" + vbLf + Infos.RichTextBox1.Text
        
        Context.State = eInitRepeat
    Else
        Infos.RichTextBox1.Text = CStr(Now) & "-" + "CATTURO I MODULI OK [" + CStr(Context.Repeat) + "]" + vbLf + Infos.RichTextBox1.Text
        Context.State = eRun
    End If
    
End Sub

Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Dim NodeId As Integer
    If (Test.Parameter(1) = "") Then
        NodeId = IndiceSezione + 1
    Else
        NodeId = CInt(Test.Parameter(1))
    End If
    If (AssegnaNodeId(NodeCaptured(IndiceSezione), NodeId, Infos.RichTextBox1) = True) Then
        
        Context.ErrorString = "LA FUNZIONE AssegnaNodeId() HA RITORNATO UN ERRORE"
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " TEST TIME = " + CStr(GetTickCountEvo - Context.StartTimer) + " (ms)" + vbCrLf
        Context.State = eSaveError
    Else
        If (StoreNodeId(NodeCaptured(IndiceSezione), Infos.RichTextBox1) = True) Then
            
            Context.ErrorString = "LA FUNZIONE StoreNodeId() HA RITORNATO UN ERRORE"
            Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " TEST TIME = " + CStr(GetTickCountEvo - Context.StartTimer) + " (ms)" + vbCrLf
            Context.State = eSaveError
        Else
            Context.State = eSpegniModuli2
            Infos.RichTextBox1.Text = CStr(Now) & "-" + "SPENGO I MODULI" + vbLf + Infos.RichTextBox1.Text
            Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
        End If
    End If
End Sub



Private Sub AccendiModulo2(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture2
    End If
    
End Sub


Private Sub SpegniModuli2(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    TurnOffAllModule
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eAccendiModulo2
        Infos.RichTextBox1.Text = CStr(Now) & "-" + "ACCENDO I MODULI" + vbLf + Infos.RichTextBox1.Text
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub Capture2(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ASSEGNA_NODEID)
    Dim str As String
    Infos.Label1.Caption = ""
    NodeCaptured(IndiceSezione) = CatturaMltPlus(, str)
    
    If (NodeCaptured(IndiceSezione) <> -1) Then
    
        Infos.Label1.Caption = str
    
        ExitCatturaMlt NodeCaptured(IndiceSezione)
    
        Context.State = eEnd
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " TEST TIME = " + CStr(GetTickCountEvo - Context.StartTimer) + " (ms)" + vbCrLf
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Else
    
        Context.State = eSpegniModuli2
    End If
End Sub

'---------------
'AssegnaNodeId()
'---------------
Public Function AssegnaNodeId(ByVal NodeId As Integer, ByVal NewNodeId As Integer, ByRef RichTextBox1 As RichTextBox) As Boolean
    
    'NOTA: passo alla funzione l'indice del canale selezionato al momento
    Dim CanaleSelezionato As Integer '(da 0 a 5)
    Dim i As Integer
    Dim bExit As Boolean
    Dim dummy As Integer
    Dim count_NodeId As Integer
    Dim NodeIdHex As String
    Dim TempString As String
    
    Dim Timeout As Double
    
    Dim indice_tentativo As Integer
    

    'label operazione
    RichTextBox1.Text = "C A M B I O  N O D E  I D" & " - " & "ID OLD =" & CStr(NodeId) & " - ID NEW =" & CStr(NewNodeId) + vbLf + RichTextBox1.Text
    
    CanaleSelezionato = Index
    
    indice_tentativo = 0
    bExit = False
    
    
    Do While (indice_tentativo <= 3 And bExit = False)
        indice_tentativo = indice_tentativo + 1
        Timeout = GetTickCount + 2000
    
        'mando messaggio CHANGE NODE ID
        '------------------------------
        SendChangeIdMessage NodeId, NewNodeId
        Attesa_Plus (0)
    
        Do While (GetTickCount < Timeout And bExit = False)
            

        
            Attesa_Plus (0)
            DoEvents
        
            Do While (GetFrameFromRxBuffer(eCaptureTask, TempString) = 0 And bExit = False)
        
                'risposta a "CHANGE NODE ID" ("T18EF228"|vecchioNodeId|"75""13"|nuovoNodeId|"FF FF FF FF FF FF")
                If ((Left$(TempString, 8) = "T18EF228") And (Mid$(TempString, 11, 2) = "75")) Then
                    bExit = True
                    FlushRxBuffer eCaptureTask
        
                   ' BackColoMltCaptured(CanaleSelezionato) = vbYellow
        
               End If 'END 'risposta a "CHANGE NODE ID"
        
            Loop 'attendo messaggio di risposta
        Loop 'END (GetTickCount < Timeout And bExit = False)
    
    Loop 'END (indice_tentativo <= 3 And bExit = False)
    If (bExit = True) Then
        AssegnaNodeId = False
    Else
        AssegnaNodeId = True
    End If
End Function 'END AssegnaNodeId()


'---------------------
'SendChangeIdMessage()
'---------------------
Private Sub SendChangeIdMessage(ByVal NodeId As Integer, ByVal NewNodeId As Integer)

    Dim campo_dati_can As String
    Dim msg_can2send As String
    Dim i As Integer
    Dim dummy As Integer
    Dim Text(8) As String
    Dim NodeIdHex As String
    
    NodeIdHex = "0" & CStr(NewNodeId)
    
    'valorizzo "campo_dati_can"
    Text(0) = "75"
    Text(1) = "13"
    Text(2) = NodeIdHex
    Text(3) = "FF"
    Text(4) = "FF"
    Text(5) = "FF"
    Text(6) = "FF"
    Text(7) = "FF"
    
    campo_dati_can = ""
    For i = 0 To 7
        campo_dati_can = campo_dati_can & Text(i)
    Next i
    
    'msg di "WRITE NODE ID"                                                     'DLC.Text
    msg_can2send = "T" & "18EF8" & CStr(NodeId) & "22" & "8" & campo_dati_can & Chr(13)
    
    
    Output msg_can2send
    dummy = DoEvents()

End Sub 'END SendChangeIdMessage()

'------------------------
'SendStoreNodeIdMessage()
'------------------------
Private Sub SendStoreNodeIdMessage(ByVal NodeId As Integer)

Dim campo_dati_can As String
Dim msg_can2send As String
Dim i As Integer
Dim dummy As Integer
Dim Text(8) As String
Dim NodeIdHex As String

'NodeIdHex = frmSetting.cmbNodeId(Index).Text
NodeIdHex = CStr(NodeId)

'valorizzo "campo_dati_can"
Text(0) = "76"
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

'msg di "STORE NODE ID"
msg_can2send = "T" & "18EF8" & NodeIdHex & "22" & "8" & campo_dati_can & Chr(13)
'msg_can2send = "T" & "18EF8" & cmbNodeId(Index) & "22" & "8" & campo_dati_can & Chr(13)


Output msg_can2send
dummy = DoEvents()



End Sub 'END SendStoreNodeIdMessage()

'-------------
'StoreNodeId()
'-------------
Public Function StoreNodeId(ByVal NodeId As Integer, ByRef RichTextBox1 As RichTextBox) As Boolean

    'NOTA: passo alla funzione l'indice del canale selezionato al momento
    Dim CanaleSelezionato As Integer '(da 0 a 5)
    Dim i As Integer
    Dim bExit As Boolean
    Dim dummy As Integer
    Dim NodeIdHex As String
    Dim TempString As String
    
    Dim Timeout As Double
    
    Dim indice_tentativo As Integer
    
    
    'se è stato premuto lo "STOP" esci
    If (bStopStartUpModuli = True) Then
        Exit Function
    End If
    
    'label operazione
    RichTextBox1.Text = "STORE NODE ID" & " - " & "ID " & CStr(NodeId) + vbLf + RichTextBox1.Text
    

    
    indice_tentativo = 0
    bExit = False
    
    
    Do While (indice_tentativo <= 3 And bExit = False)
        indice_tentativo = indice_tentativo + 1
        Timeout = GetTickCount + 2000
    
    
        'mando messaggio STORE NODE ID
        '-----------------------------
        SendStoreNodeIdMessage (NodeId)
        Attesa_Plus (0)
    
        Do While (GetTickCount < Timeout And bExit = False)
            

            
            Attesa_Plus (0)
            DoEvents
            
            'attendo messaggio di risposta
            Do While (GetFrameFromRxBuffer(eCaptureTask, TempString) = 0)
        
                'risposta a "STORE NODE ID" ("T18EF228"|NodeId|DLC=8|"76")
                If ((Left$(TempString, 8) = "T18EF228") And (Mid$(TempString, 11, 2) = "76")) Then
                    bExit = True
           
                End If 'END 'risposta a "STORE NODE ID"
           
            Loop 'attendo messaggio di risposta
        Loop 'END (GetTickCount < Timeout And bExit = False)
    
    Loop 'END (indice_tentativo <= 3 And bExit = False)

    If (bExit = True) Then
        StoreNodeId = False
    Else
        StoreNodeId = True
        RichTextBox1.Text = "ERROR " + vbLf + RichTextBox1.Text
    End If
End Function 'END StoreNodeId()


Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Dim FileTest As String
    Infos.RichTextBox1.Text = "-------------------------------" + vbCrLf + _
                              "LISTA PARAMETRI" + vbCrLf + _
                             "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFile                      = " + Test.Parameter(eNomeFile) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
    
End Sub


