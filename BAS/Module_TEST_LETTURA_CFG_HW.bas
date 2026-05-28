Attribute VB_Name = "Module_TEST_LETTURA_CFG_HW"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 300
Private Const ACCENSIONE_TIMEOUT_MS = 500


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

' Sotto stato macchina dello stato eRun
Private Enum eMachineRunState
    eRunState_PARAM_REL_FW              ' Lettura della release dell'hardware
    eRunState_PARAM_REL_HW              ' Lettura del codice dell'hardware
    eRunState_ERROR
    
    eRunState_FINE
End Enum

Private Enum eTestParameter
    eSincro = 1                         ' Flag di comando test bloccante o meno
    eFirmwareVersion                    ' Stringa della versione del firmware
    eReleaseHardware                    ' Stringa del release hardware del prodotto
    eJumpSpegniModuli
End Enum

' Struttura dati del modulo
Public Type TagTEST_LETTURA_CFG_HW
    bOldOn As Boolean               ' difu del comando di start del modulo
    Return As Integer                ' Stato del modulo
    State As Integer
    OldState As Integer
    NodeCaptured As Integer
    Timer As Long
    StartTime As Long
    Repeat As Long
    
    RunState As Integer             ' Stato macchina interno del modulo usato per effettuare le varie richieste al modulo
End Type

' Area dati del modulo
Private Context As TagTEST_LETTURA_CFG_HW

Public Sub INFO_TEST_LETTURA_CFG_HW(Infos As TagInfosSingleTest)
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

Private Sub TEST_LETTURA_CFG_HW_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_LETTURA_CFG_HW)
    Context1.State = eEnd
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = CStr(Now) & "-" + "TEST LETTURA CFG" + vbLf
    End If
End Sub

Public Sub TEST_LETTURA_CFG_HW(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_LETTURA_CFG_HW, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "LETTURA_CFG_HW") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
        bNewState = True
    Else
        
        INFO_TEST_LETTURA_CFG_HW Infos
        
        If (bStop = True And bOn = True) Then
        
            TEST_LETTURA_CFG_HW_INIT_STRUCT Infos, Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
            Context1.State = eError
            bNewState = True
        ElseIf (bOn = True And Context1.bOldOn = False) Then          ' fronte di salita del test ( AVVIO )
            Context1.StartTime = GetTickCountEvo
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




Private Sub InitRepeat(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)
    If (Test.Parameter(eSincro) = "") Then
        Context.Return = eTEST_BUSY
    Else
        Context.Return = eTEST_BUSY_NON_BLOCCANTE
    End If
    
    
    If (Test.Parameter(eJumpSpegniModuli) = "") Then
        Context.State = eSpegniModuli
    Else
        If (GetModuleState(IndiceSezione) = True) Then
            Context.State = eCapture
        Else
            Context.State = eAccendiModulo
        End If
    End If
    bNewState = True
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.RunState = eRunState_PARAM_REL_FW
        
    Context.Repeat = Context.Repeat + 1
    
    If (Context.Repeat >= 5) Then
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eError
    End If
        
    Infos.RichTextBox1.BackColor = vbWhite
    
End Sub



' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)

    If (Test.Parameter(eSincro) = "") Then
        Context.Return = eTEST_BUSY
    Else
        Context.Return = eTEST_BUSY_NON_BLOCCANTE
    End If
    
    
    If (Test.Parameter(eJumpSpegniModuli) = "") Then
        Context.State = eSpegniModuli
    Else
        If (GetModuleState(IndiceSezione) = True) Then
            Context.State = eCapture
        Else
            Context.State = eAccendiModulo
        End If
    End If
    bNewState = True
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.RunState = eRunState_PARAM_REL_FW
    Context.Repeat = 0
        
    Infos.RichTextBox1.BackColor = vbWhite
    
    
    Prepare_Textbox_Parametri Test, Infos
    
End Sub

' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)
On Error GoTo err:
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
        'Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "STATE = " + CStr(Context.State) + vbLf
    End If
    Context.OldState = Context.State
    Exit Sub

err:
    Infos.RichTextBox1.Text = CStr(Now) + "-" + err.Description + vbLf + Infos.RichTextBox1.Text
    Context.State = eError
End Sub

Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)
    Infos.RichTextBox1.BackColor = vbRed
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = Infos.RichTextBox1.Text + "ERRORE SUL TEST" + vbLf
    End If
    Context.Return = eTEST_ERROR

End Sub


Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)
    
    If (Test.Parameter(eSincro) <> "") Then
        TurnOffModule IndiceSezione
        If (Context.State <> Context.OldState) Then
            Infos.RichTextBox1.Text = CStr(Now) & "-" + " SPENGO NODO ID " + CStr(NodeCaptured(IndiceSezione)) + vbLf + Infos.RichTextBox1.Text
    End If
    Else
        If (Context.State <> Context.OldState) Then
            Infos.RichTextBox1.Text = CStr(Now) & "-" + " SPENGO TUTTI MODULI" + vbLf + Infos.RichTextBox1.Text
        End If
        TurnOffAllModule
    End If
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eAccendiModulo
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)

    Dim str As String
    If (Test.Parameter(eSincro) = "") Then
        Infos.RichTextBox1.Text = CStr(Now) & "-" + " CATTURO TUTTI I NODE ID" + vbLf + Infos.RichTextBox1.Text
        
        Infos.Label1.Caption = ""
        NodeCaptured(IndiceSezione) = CatturaMltPlus(, str)
        Infos.Label1.Caption = str
    Else
        
        NodeCaptured(IndiceSezione) = IndiceSezione + 1
        Infos.RichTextBox1.Text = CStr(Now) & "-" + " CATTURO NODE ID " + CStr(NodeCaptured(IndiceSezione)) + vbLf + Infos.RichTextBox1.Text
        
        Infos.Label1.Caption = ""
        CatturaMltPlus NodeCaptured(IndiceSezione), str
        Infos.Label1.Caption = str
    End If
    
    Context.State = eRun
End Sub

Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_LETTURA_CFG_HW)
    Dim TempStr As String
    Dim thisDate As Date
    Dim Serial As String
    Dim str As String
    Dim ReleaseHardware As String
    Select Case Context.RunState
    Case eRunState_PARAM_REL_FW
        
        Infos.RichTextBox1.Text = CStr(Now) & "-" + " EXPECT REL FW = " + Test.Parameter(eFirmwareVersion) + vbLf + Infos.RichTextBox1.Text
        TempStr = CStr(Now) + "- READ REL FW = "
        
        ' Leggo la release del firmware
        If (ReadFd5_PARAM_REL_FW(NodeCaptured(IndiceSezione), ReleaseHardware) = True) Then
                                    
             
            If (ReleaseHardware <> Test.Parameter(eFirmwareVersion)) Then
                
                Context.RunState = eRunState_ERROR              ' Errore devo fermare la sequenza
                
                TempStr = TempStr + ReleaseHardware + " - ERROR" + vbLf
            Else
                TempStr = TempStr + ReleaseHardware + " - OK" + vbLf
                Context.RunState = eRunState_PARAM_REL_HW
            End If
        Else
            TempStr = TempStr + "NO ANSWAER - ERROR" + vbLf
            
            Context.State = eInitRepeat
        End If
        
        Infos.RichTextBox1.Text = TempStr + Infos.RichTextBox1.Text
        
    Case eRunState_PARAM_REL_HW:
        
        Infos.RichTextBox1.Text = CStr(Now) & "-" + " EXPECT REL HW = " + Test.Parameter(3) + vbLf + Infos.RichTextBox1.Text
        TempStr = CStr(Now) & "-" + " READ REL HW = "
        
        ' Leggo la release del firmware
        If (ReadFd5_PARAM_REL_HW(NodeCaptured(IndiceSezione), ReleaseHardware) = True) Then
                        
            TempStr = TempStr + ReleaseHardware
            
            If (ReleaseHardware <> Test.Parameter(3)) Then
                TempStr = TempStr + " - ERROR" + vbLf
                Context.RunState = eRunState_ERROR              ' Errore devo fermare la sequenza
            Else
                TempStr = TempStr + " - OK" + vbLf
                Context.RunState = eRunState_FINE
            End If
        Else
            TempStr = TempStr + "NO ANSWEAR" + " - HW OK" + vbLf
            Context.State = eInitRepeat
        End If
    
        Infos.RichTextBox1 = TempStr + Infos.RichTextBox1.Text
        
    ' Stato di errore interno
    Case eRunState_ERROR:
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + CStr(Now) & "-" + " TEST TIME = " + CStr(GetTickCountEvo - Context.StartTime) + " (ms)" + vbCrLf
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eError
        Infos.Label1.Caption = ""
        ExitCatturaMlt CatturaMltPlus(, str)
        Infos.Label1.Caption = str
    
    Case eRunState_FINE:
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + CStr(Now) & "-" + " TEST TIME = " + CStr(GetTickCountEvo - Context.StartTime) + " (ms)" + vbCrLf
        Context.State = eEnd
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        
        Infos.Label1.Caption = ""
        ExitCatturaMlt CatturaMltPlus(, str)
        Infos.Label1.Caption = str

    Case Else
            ExitCatturaMlt CatturaMltPlus(, str)
    End Select

    
End Sub




Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eSincro                    = " + Test.Parameter(eSincro) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eFirmwareVersion           = " + Test.Parameter(eFirmwareVersion) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eReleaseHardware           = " + Test.Parameter(eReleaseHardware) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eJumpSpegniModuli          = " + Test.Parameter(eJumpSpegniModuli) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
        

End Sub


