Attribute VB_Name = "Module_TEST_ACQUISIZIONE_CAN"
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Const SPEGMINENTO_TIMEOUT_MS = 200
Private Const ACCENSIONE_TIMEOUT_MS = 200
Private Y(2) As Double
'eTEST_OK = 2
'eTEST_BUSY = 1
'eTEST_STOP = 0
'eTEST_ERROR = -1



' PARAMETRI
' 1 - Nome File della curva di comando
' 2 - Nome File della rampa applicata ( in V vs Tempo )
' 3 - Source Address da utilizzare
' 4 - Nome File del limite inferiore
' 5 - Nome File del limite superiore


' Stato macchina del modulo
Private Enum eMachineState
    eInit
    eSpegniModuli
    eAccendiModulo
    eCapture
    eRun
    esavefile
    eEnd
    eStart
    eError
End Enum


Private Enum eTestParameter
    eNomeFileCurvaComando = 1                           ' Nome del file della curva di comando
    eFolderXY_AcqFile = 2                               ' Nome del file dell'ultima acquisizione
    eSourceAddressComando = 3                           ' SOurce address del messaggio CAN AVC da utilizzare come comando
    eAcqZeroValue = 4                                   ' Flag di attivazione dell'attesa della validità del messaggio di controllo
    eAcqDeltaValue = 5                                  ' Non Usato
    NomeFileCurvaLimiteInferiore = 6
    NomeFileCurvaLimiteSuperiore = 7
End Enum

' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagPuntiAcquisizione
    X As Double
    Y As Double
    Z As Double
    Mlt_Error As Long
    MLT_Temperature As Double
    MMS2218_Temperature As Double
    Pressure As Double
    TimeStamp As Double
    Error As Long
End Type

Public Type TagTEST_ACQUISIZIONE_CAN
    ErrorString As String
    bOldOn As Boolean                       ' difu del comando di start del modulo
    Return As Integer                       ' Stato del modulo
    State As Integer
    NomeFile As String                      ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean                   ' Flag segnalazione primo avvio del programma
    NodeCaptured As Integer
    Timer As Long
    CURVE_LIMITE As CurveLimitePlusTag       ' Struttura dati delle curve limite
    CURVE_COMANDO As TagChannel             ' File di configurazione delle curve comando CAN in funzione di V teorico
    CURVE_XY As TagChannelRampa             ' File di configurazione della rampa usata per effettuare l'acquisizione
    Punto() As TagPuntiAcquisizione         ' Punti della acquisizione
    Tempo_sec_start As Double               ' Istante di avvio della sequenza espresso in secondi
    NumeroPuntiAcquisiti As Long            ' numero di punti acquisiti
    OldCommandInMilliVolts As Double
    OldTest(NomeFileCurvaLimiteSuperiore) As String
End Type


Private Context As TagTEST_ACQUISIZIONE_CAN


Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text = "-------------------------------" + vbCrLf + _
                              "LISTA PARAMETRI" + vbCrLf + _
                              "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "File Curva Comando           = " + Test.Parameter(eNomeFileCurvaComando) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "XY File                      = " + Test.Parameter(eFolderXY_AcqFile) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eSourceAddressComando        = " + Test.Parameter(eSourceAddressComando) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eAcqZeroValue                = " + Test.Parameter(eAcqZeroValue) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eAcqDeltaValue               = " + Test.Parameter(eAcqDeltaValue) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eAcqDeltaValue               = " + Test.Parameter(eAcqDeltaValue) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileCurvaLimiteInferiore = " + Test.Parameter(NomeFileCurvaLimiteInferiore) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileCurvaLimiteSuperiore = " + Test.Parameter(NomeFileCurvaLimiteSuperiore) + vbCrLf

End Sub



Public Sub Graph_Points(ByRef CWGraph1 As CWGraph, ByRef Context As TagTEST_ACQUISIZIONE_CAN, ByRef ulPointIndex As Long, ByVal ClearData As Boolean)
    Dim i As Long
    Dim TempY As Double
    Dim TempX As Double
    Dim TempZ As Double
    Dim DeltaError As Double
    
    If (ClearData = True) Then
        CWGraph1.Plots(3).ClearData
    End If
    CWGraph1.Plots(3).Visible = False
    i = ulPointIndex
    Do While (i < UBound(Context.Punto))
        If (i < 2) Then
            CWGraph1.Plots(3).ChartXvsY Context.Punto(i).X, Context.Punto(i).Y
        Else
            TempY = (Context.Punto(i - 1).Y) '+ Context.Punto(i).Y) / 2
            TempX = Context.Punto(i).X
            TempZ = (Context.Punto(i - 1).Z)
            'TempY = (Context.Punto(i - 4).Y + Context.Punto(i - 3).Y + Context.Punto(i - 2).Y + Context.Punto(i - 1).Y + Context.Punto(i).Y) / 2
            'TempX = (Context.Punto(i - 4).X + Context.Punto(i - 3).X + Context.Punto(i - 2).X + Context.Punto(i - 1).X + Context.Punto(i).X) / 2
            
            If (TempY > Context.Punto(i).Y) Then
                DeltaError = TempY - Context.Punto(i).Y
            Else
                DeltaError = Context.Punto(i).Y - TempY
            End If
            
            If (DeltaError > 0.5) Then
                TempY = Context.Punto(i).Y
                TempX = Context.Punto(i).X
            End If
            
            CWGraph1.Plots(3).ChartXvsY Context.Punto(i).X, TempY
            CWGraph1.Plots(4).ChartXvsY Context.Punto(i).X, TempZ
            
        End If
        
        
        i = i + 1
    Loop
    
    ulPointIndex = UBound(Context.Punto)
    CWGraph1.Plots(3).Visible = True
    CWGraph1.Plots(4).Visible = True
End Sub

' Funzione di inizializzazione dei settaggi del grafico. Carica le curve limite
Private Sub GraphSettings(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
On Error GoTo err:
    Dim bUpdate As Boolean
    Dim i As Integer
    Dim par As CurveLimitePlusParametersTag
    


    ' inizializzazione delle curve limite
    par.Indice(eInferiore) = 1
    par.Indice(eSuperiore) = 2

    CurveLimitePlusInit Context.CURVE_LIMITE, par


    If (Context.OldTest(eNomeFileCurvaComando) <> Test.Parameter(eNomeFileCurvaComando)) Then
        ' Inizializzazione dei comandi da applicare
        
        If (Update_CURVECOMANDO(Context.CURVE_COMANDO, Settings.FolderFileCurveComando + "\" + Test.Parameter(eNomeFileCurvaComando)) = True) Then
            
            Context.ErrorString = "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI COMANDO CAN "
            Prepare_MsgBox Infos, Context.ErrorString, vbCrtical
            Context.State = eError
        
        End If

        bUpdate = True
        Context.OldTest(eNomeFileCurvaComando) = Test.Parameter(eNomeFileCurvaComando)
    End If


    If (Context.OldTest(eFolderXY_AcqFile) <> Test.Parameter(eFolderXY_AcqFile)) Then

         ' Inizializzazione dei file della rampa da applicare
        If (Update_CURVE_XY(Context.CURVE_XY, Settings.FolderRampeXY + "\" + Test.Parameter(eFolderXY_AcqFile)) = True) Then
        
            Context.ErrorString = "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI RAMPE"
            Prepare_MsgBox Infos, Context.ErrorString, vbCritical
        
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eFolderXY_AcqFile) = Test.Parameter(eFolderXY_AcqFile)
    End If

    
    If (Context.OldTest(NomeFileCurvaLimiteInferiore) <> Test.Parameter(NomeFileCurvaLimiteInferiore) Or _
        Context.OldTest(NomeFileCurvaLimiteSuperiore) <> Test.Parameter(NomeFileCurvaLimiteSuperiore)) Then
    
        Infos.CWGraph1.ClearData
    End If


    If (Context.OldTest(NomeFileCurvaLimiteInferiore) <> Test.Parameter(NomeFileCurvaLimiteInferiore)) Then
    
        If (CurveLimitePlusLoad(Context.CURVE_LIMITE, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteInferiore), eInferiore) = eERROR_CurveLimite) Then
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CURVE LIMITE INFERIORE", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph1, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteInferiore), eInferiore
        End If


        Context.OldTest(NomeFileCurvaLimiteInferiore) = Test.Parameter(NomeFileCurvaLimiteInferiore)
        bUpdate = True
    End If



    If (Context.OldTest(NomeFileCurvaLimiteSuperiore) <> Test.Parameter(NomeFileCurvaLimiteSuperiore)) Then

        If (CurveLimitePlusLoad(Context.CURVE_LIMITE, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteSuperiore), eSuperiore) = eERROR_CurveLimite) Then
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CURVE LIMITE SUPERIORE", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph1, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteSuperiore), eSuperiore
        End If


        Context.OldTest(NomeFileCurvaLimiteSuperiore) = Test.Parameter(NomeFileCurvaLimiteSuperiore)
        bUpdate = True
    End If


    'If (bUpdate = True) Then
        Infos.CWGraph1.PlotAreaColor = vbWhite
        Infos.CWGraph1.Plots(3).PointColor = vbBlue
        Infos.CWGraph1.Plots(3).LineColor = vbBlue

        
    'End If


    
    
    
    
    
    
    
    
    
    
    

Exit Sub
err:
Debug.Print err.Description
End Sub




Private Sub INFO_ACQUISIZIONE_CAN(Infos As TagInfosSingleTest)
    If (Infos.CWGraph1.Visible <> True) Then
        Infos.CWGraph1.Visible = True
    End If
    
    If (Infos.RichTextBox1.Visible = True) Then
        Infos.RichTextBox1.Visible = False
    End If
    
    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
    End If
    
End Sub

Private Sub TEST_ACQUISIZIONE_CAN_INIT_STRUCT(Infos As TagInfosSingleTest, Context1 As TagTEST_ACQUISIZIONE_CAN)
    Context1.State = eEnd
    
    Infos.CWGraph1.ClearData
    Infos.CWGraph2.ClearData
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = ""
    End If
    
End Sub

Public Sub TEST_ACQUISIZIONE_CAN(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_ACQUISIZIONE_CAN, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "ACQUISIZIONE_CAN") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        INFO_ACQUISIZIONE_CAN Infos
        
       'GraphSettings Test, Infos, IndiceSezione, Context
        If (bStop = True And bOn = True) Then
        
            TEST_ACQUISIZIONE_CAN_INIT_STRUCT Infos, Context1


        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            
            Context.ErrorString = "STOP DA PARTE DELL'UTENTE"
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
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    Context.Return = eTEST_BUSY_NON_BLOCCANTE
    'Context.State = eStart
    Context.State = eSpegniModuli
    Context.Timer = GetTickCount + SPEGMINENTO_TIMEOUT_MS
    Context.bFirstTime = True
    
    GraphSettings Test, Infos, IndiceSezione, Context
    
    Context.NumeroPuntiAcquisiti = 0
    
    
    Prepare_Textbox_Parametri Test, Infos
    
    
    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
End Sub

' Task del modulo
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    Select Case Context.State
    Case eInit: Init Test, Infos, IndiceSezione, Context
    Case eSpegniModuli: SpegniModuli Test, Infos, IndiceSezione, Context
    Case eAccendiModulo: AccendiModulo Test, Infos, IndiceSezione, Context
    Case eCapture: Capture Test, Infos, IndiceSezione, Context
    Case eRun: Run Test, Infos, IndiceSezione, Context
    Case eStart: Start Test, Infos, IndiceSezione, Context
    Case esavefile: SaveFile Test, Infos, IndiceSezione, Context
    Case eEnd: Fine Test, Infos, IndiceSezione, Context
    Case eError: ErrorTest Test, Infos, IndiceSezione, Context
    Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
End Sub

Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
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


Private Sub SpegniModuli(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    'TurnOffAllModule
    TurnOffModule IndiceSezione
    If (Context.Timer < GetTickCount) Then
        Context.State = eAccendiModulo
        Context.Timer = GetTickCount + ACCENSIONE_TIMEOUT_MS
    End If
End Sub

Private Sub AccendiModulo(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    TurnOnModule IndiceSezione
    
    If (Context.Timer < GetTickCount) Then
        Context.State = eCapture
    End If
    
End Sub

Private Sub Capture(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    Dim str As String
'    If (NodeCaptured(IndiceSezione) = -1) Then
'        NodeCaptured(IndiceSezione) = CatturaMltPlus()
'    Else
'        NodeCaptured(IndiceSezione) = NodeCaptured(IndiceSezione)
'    End If

    'NodeCaptured(IndiceSezione) = CatturaMltPlus()
    
    NodeCaptured(IndiceSezione) = IndiceSezione + 1
    
    'Infos.Label1.Caption = ""
    'ExitCatturaMlt CatturaMltPlus(NodeCaptured(IndiceSezione), str)
    'Infos.Label1.Caption = str

    
    Context.State = eStart
End Sub

Private Sub SaveFile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    SalvaPunti Test, Infos, Context, IndiceSezione
    
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    
    Context.State = eEnd
End Sub

Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)

    If (Infos.CWGraph1.PlotAreaColor = vbRed) Then
        
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
End Sub


Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    Dim Tempo_sec As Double
    Dim CommandInMilliVolts As Double
    Dim State As MltState
    Dim bPunti_Y_Validi As Boolean
    Dim Percentual As Integer
    On Error Resume Next
    
    
    If (Context.Timer < GetTickCount) Then
        
        Tempo_sec = GetTickCount / 1000 - Context.Tempo_sec_start
        Context.Timer = GetTickCount + 50
        
        
        ' Prelevo il comando dalla rampa applicata
        If (Get_CURVE_XY_VALUE(Context.CURVE_XY, Tempo_sec, CommandInMilliVolts) = eERROR_Rampe) Then
            Context.State = esavefile
        Else
            If (Get_CURVECOMANDO_VALUE(Context.CURVE_COMANDO, CommandInMilliVolts / 1000, State, Percentual) = eError_CURVE_COMANDO_CAN) Then
                Context.ErrorString = "ERROR Curva Comando Value"
                Context.State = eError
            Else
            
            End If
        End If
        
        
        If (Context.State <> esavefile) Then
        
            Context.Punto(Context.NumeroPuntiAcquisiti).TimeStamp = Tempo_sec
            Context.Punto(Context.NumeroPuntiAcquisiti).Mlt_Error = MLT(IndiceSezione).Fault
            
            If (Test.Parameter(eAcqZeroValue) = "1") Then
                If (MLT(IndiceSezione).bActive = True) Then
                    bPunti_Y_Validi = True
                End If
                
                Context.Punto(Context.NumeroPuntiAcquisiti).Y = MLT(IndiceSezione).Position
                Context.Punto(Context.NumeroPuntiAcquisiti).Z = CE16(IndiceSezione).Position
                
                                
            Else
                If (CE16(IndiceSezione).bActive = True) Then
                    bPunti_Y_Validi = True
                End If
            
                Context.Punto(Context.NumeroPuntiAcquisiti).Y = CE16(IndiceSezione).Position
                Context.Punto(Context.NumeroPuntiAcquisiti).Z = MLT(IndiceSezione).Position
            
            End If
            
            
            If (Context.bFirstTime = True) Then
                Context.bFirstTime = False
                Context.Punto(Context.NumeroPuntiAcquisiti).X = CommandInMilliVolts
                
            End If
            
            If (CurveLimitePlusCheck(Context.CURVE_LIMITE, _
                                          Context.Punto(Context.NumeroPuntiAcquisiti).X, _
                                          Context.Punto(Context.NumeroPuntiAcquisiti).Y) = eERROR_CurveLimite) Then
                If (Infos.CWGraph1.PlotAreaColor <> vbRed) Then
                    Distributore.Text2 = Context.Punto(Context.NumeroPuntiAcquisiti).X
                    Distributore.Text1 = Context.Punto(Context.NumeroPuntiAcquisiti).Y
                    Infos.CWGraph1.PlotAreaColor = vbRed
                End If
            End If

            Infos.CWGraph1.Plots(3).ChartXvsY Context.Punto(Context.NumeroPuntiAcquisiti).X, Context.Punto(Context.NumeroPuntiAcquisiti).Y
            'Infos.CWGraph1.Plots(4).ChartXvsY Context.Punto(Context.NumeroPuntiAcquisiti).x, Context.Punto(Context.NumeroPuntiAcquisiti).Z
            

            
            Context.Punto(Context.NumeroPuntiAcquisiti).MLT_Temperature = CE16(IndiceSezione).Temperature
            Context.Punto(Context.NumeroPuntiAcquisiti).Pressure = GetCalibratedValue(ePressione)
            Context.Punto(Context.NumeroPuntiAcquisiti).MMS2218_Temperature = GetCalibratedValue(eTemperatura)
            
            
            
            
            If (bPunti_Y_Validi = True) Then
                Context.NumeroPuntiAcquisiti = Context.NumeroPuntiAcquisiti + 1
                ReDim Preserve Context.Punto(Context.NumeroPuntiAcquisiti)
            End If
            
            
            
            Context.Punto(Context.NumeroPuntiAcquisiti).X = CommandInMilliVolts
            
            Context.Punto(Context.NumeroPuntiAcquisiti).Error = Send_AVC_MESSAGE(State, Percentual, IndiceSezione, Test.Parameter(eSourceAddressComando))
            
        End If
        
        Context.OldCommandInMilliVolts = CommandInMilliVolts
    End If
End Sub


Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_CAN)
    Context.State = eRun
    Context.Tempo_sec_start = GetTickCount / 1000
    
    Infos.CWGraph1.Plots(3).ClearData
End Sub

Private Function Send_AVC_MESSAGE(ByRef State As MltState, ByRef Percentual As Integer, ByVal IndiceSezione As Integer, ByRef SourceAddress As String) As Long
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
    'Output str
    
    Send_AVC_MESSAGE = Output_Test(str)
End Function


Private Sub SalvaPunti(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_ACQUISIZIONE_CAN, IndiceSezione As Integer)
    Dim FileName As String
    Dim bSaveLastAcq As Boolean
On Error GoTo err:
    Dim PathDirectory As String
    Dim i As Long
    Dim ora As String
    Dim oggi As String
    
    ora = Format(Time, ("HH.MM.SS"))
    ora = Replace(ora, ".", "")

    oggi = Format(Date, ("yy/mm/dd"))
    oggi = Replace(oggi, "/", "")
    
    
    PathDirectory = CreaDirectory(Infos, Settings.FolderGraphSaved)
    
    If (PathDirectory = "") Then
        PathDirectory = Settings.FolderGraphSaved + "\"
    End If
    
    If (Infos.Label1 <> "") Then
        FileName = PathDirectory + Test.Name + "_" + Infos.Label1.Caption + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
    Else
        FileName = PathDirectory + Test.Name + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
    End If
    
    On Error Resume Next
    Close #22
    
    Open FileName For Append As #22
    Print #22, "COMMAND;POSITION_CE16_(mm);POSITION_MLT_(pts);TIMESTAMP;ERROR_LAWICELL;ERROR_MLT"
    Do While (i < UBound(Context.Punto))
        Print #22, CStr(Context.Punto(i).X) + ";" + CStr(Context.Punto(i).Y) + ";" + CStr(Context.Punto(i).Z) + ";" + CStr(Context.Punto(i).TimeStamp) + ";" + CStr(Context.Punto(i).Error) + ";" + CStr(Context.Punto(i).Mlt_Error)
        i = i + 1
    Loop

    Close #22
    
    ' Voglio salvalre l'ultima acq
    bSaveLastAcq = True
    
    If (Infos.Label1.Caption = "") Then
        If (MsgBox("Attenzione non è stata associata nessun codice al modulo, voler procedere al salvataggio della curva ?", vbYesNo) = vbYes) Then
            Infos.Label1.Caption = Immissione_SerialNumber()
        Else
            ' Non voglio salvare l'ultima acq
            bSaveLastAcq = False
        End If
    End If
    
    If (bSaveLastAcq = True) Then
        
        PathDirectory = CreaDirectory(Infos, Settings.FolderGraphSaved_LastACQ)
        
        If (PathDirectory = "") Then
            PathDirectory = Settings.FolderGraphSaved_LastACQ + "\"
        End If
        
        
        ' Apro il file di destinazione dove salvare l'ultima curva utile effettuata dal modulo
        FileName = PathDirectory + Infos.Label1.Caption + ".csv"
    
        ' Sovrascrivo sempre l'ultimo
        Open FileName For Output As #22
        i = 0
        Print #22, "COMMAND;POSITION_CE16_(mm);POSITION_MLT_(pts);TIMESTAMP;ERROR;ERROR_MLT"
        Do While (i < UBound(Context.Punto))
            Print #22, CStr(Context.Punto(i).X) + ";" + CStr(Context.Punto(i).Y) + ";" + CStr(Context.Punto(i).Z) + ";" + CStr(Context.Punto(i).TimeStamp) + ";" + CStr(Context.Punto(i).Error) + ";" + CStr(Context.Punto(i).Mlt_Error)
            i = i + 1
        Loop
    
        Close #22
    End If
Exit Sub
err:
End Sub



Private Function CreaDirectory(Infos As TagInfosSingleTest, ByVal StartDirectory As String) As String
    On Error GoTo err:
    Dim PathDirectory  As String
        
        CreaDirectory = StartDirectory + "\" + Infos.codicePRodotto.Caption + "\"
        
        MkDir (CreaDirectory)


        Exit Function
err:

End Function



Public Function DirExists(ByVal Path As String) As Boolean
    On Error Resume Next
    'Legge l'attributo e si assicura che si tratti di una directory
    FileExists = GetAttr(Path) And vbDirectory
    
    
    'Se avviene un errore la Function restituisce False
End Function

