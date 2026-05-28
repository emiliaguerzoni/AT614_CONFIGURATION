Attribute VB_Name = "Module_TEST_SCRITTURA_SERIAL_NUMBER"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 200
Private Const ACCENSIONE_TIMEOUT_MS = 200
'eTEST_OK = 2
'eTEST_BUSY = 1
'eTEST_STOP = 0
'eTEST_ERROR = -1



' Stato macchina del modulo
Private Enum eMachineState
    eInit = 0
    eSpegniModuli = 1
    eAccendiModulo = 2
    eCapture = 3
    eRun = 4
    eEnd = 5
    eError = 6
    eInitRepeat = 7
End Enum


Private Enum eTestParameter
    eNomeFile = 1
    eSincro = 2
End Enum


Public Type TagTEST_SCRITTURA_SERUAL_NUMBER
    bOldOn As Boolean               ' difu del comando di start del modulo
    Return As Integer                ' Stato del modulo
    State As Integer
    OldState As Integer
    NodeCaptured As Integer
    Timer As Long
    Repeat As Long
End Type

Private Context As TagTEST_SCRITTURA_SERUAL_NUMBER

Public Sub INFO_SCRITTURA_SERIAL_NUMBER(Infos As TagInfosSingleTest)
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


Private Sub TEST_SCRITTURA_SERIAL_NUMBER_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_SCRITTURA_SERUAL_NUMBER)
    Context1.State = eEnd
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    Infos.RichTextBox1.Text = "SCRITTURA SERIAL NUMBER"
End Sub



Public Sub TEST_SCRITTURA_SERIAL_NUMBER(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_SCRITTURA_SERUAL_NUMBER, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "SERIAL_NUMBER") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
        bNewState = True
    Else
        
        INFO_SCRITTURA_SERIAL_NUMBER Infos
        
        
        If (bStop = True And bOn = True) Then
        
            TEST_SCRITTURA_SERIAL_NUMBER_INIT_STRUCT Infos, Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context1.State = eError
            bNewState = True
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
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)

    If (Test.Parameter(eSincro) = "") Then
        Context.Return = eTEST_BUSY
    Else
        Context.Return = eTEST_BUSY_NON_BLOCCANTE
    End If
    
    Context.State = eSpegniModuli
    bNewState = True
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.Repeat = 0
    
    Infos.RichTextBox1.Text = ""
    Prepare_Textbox_Parametri Test, Infos
    
    Infos.RichTextBox1.BackColor = vbWhite
End Sub


' Questa funzione inizializza la memoria del frame
Private Sub InitRepeat(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)

    If (Test.Parameter(eSincro) = "") Then
        Context.Return = eTEST_BUSY
    Else
        Context.Return = eTEST_BUSY_NON_BLOCCANTE
    End If
    
    Context.State = eSpegniModuli
    bNewState = True
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    
    Infos.RichTextBox1.BackColor = vbWhite
    
    Context.Repeat = Context.Repeat + 1
    If (Context.Repeat > 5) Then
        Context.State = eError
    End If
End Sub

' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)
    Select Case Context.State
    Case eInit: Init Test, Infos, IndiceSezione, Context
    Case eInitRepeat: InitRepeat Test, Infos, IndiceSezione, Context
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, Context
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, Context
    Case eCapture: Capture Test, Infos, IndiceSezione, Context
    Case eRun: Run Test, Infos, IndiceSezione, Context
    Case eEnd: Context.Return = eTEST_OK
    Case eError: ErrorTest Test, Infos, IndiceSezione, Context
    Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
    If (Context.OldState <> Context.State) Then
        
    End If
    Context.OldState = Context.State
    
End Sub

Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)
    Infos.RichTextBox1.BackColor = vbRed
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = Infos.RichTextBox1.Text + CStr(Now) + "- ERRORE SUL TEST" + vbLf
    End If
    Context.Return = eTEST_ERROR

End Sub


Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)
    
    If (Test.Parameter(eSincro) <> "") Then
        TurnOffModule IndiceSezione
        If (Context.State <> Context.OldState) Then
            Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + CStr(Now) + " SPENGO NODO ID " + CStr(NodeCaptured(IndiceSezione)) + vbLf
    End If
    Else
        If (Context.State <> Context.OldState) Then
            Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + CStr(Now) + " SPENGO TUTTI MODULI" + vbLf
        End If
        TurnOffAllModule
    End If
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eAccendiModulo
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)

    Dim str As String
    If (Test.Parameter(eSincro) = "") Then
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + CStr(Now) + " CATTURO TUTTI I NODE ID" + vbLf
        
        Infos.Label1.Caption = ""
        NodeCaptured(IndiceSezione) = CatturaMltPlus(, str)
        Infos.Label1.Caption = str
        
        
        If (NodeCaptured(IndiceSezione) = -1) Then
            Context.State = eInitRepeat
        Else
            Context.State = eRun
        End If
    Else
        
        NodeCaptured(IndiceSezione) = IndiceSezione + 1
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + CStr(Now) + " CATTURATO NODE ID " + CStr(NodeCaptured(IndiceSezione)) + vbLf
        
        Infos.Label1.Caption = ""
        
        If (CatturaMltPlus(NodeCaptured(IndiceSezione), str) = -1) Then
            Context.State = eInitRepeat
        Else
            Context.State = eRun
        End If
        
        Infos.Label1.Caption = str
    End If
    
    
End Sub

Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_SCRITTURA_SERUAL_NUMBER)
    
    Dim thisDate As Date
    Dim Serial As String
    Dim str As String
    
    If (Aggiorna_SERIAL_NUMBER(Immissione_SerialNumber(Infos.Label1.Caption), NodeCaptured(IndiceSezione), Infos.RichTextBox1) = True) Then
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eEnd
    Else
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eError
    End If
    
        
    Infos.Label1.Caption = ""
    ExitCatturaMlt CatturaMltPlus(, str)
    Infos.Label1.Caption = str
    
End Sub



Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    
    Infos.RichTextBox1.Text = "-------------------------------" + vbCrLf + _
                              "LISTA PARAMETRI" + vbCrLf + _
                            "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFile                  = " + Test.Parameter(eNomeFile) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eSincro                    = " + Test.Parameter(eSincro) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
        
        
End Sub


