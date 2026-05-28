Attribute VB_Name = "Module_TEST_ACQUISIZIONE_PRE"

'**************************************************************
'       M O D U L O :    Module_TEST_ACQUISIZIONE_PRE.bas
'**************************************************************

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SPEGNINENTO_TIMEOUT_MS = 200
Private Const ACCENSIONE_TIMEOUT_MS = 200

Private Y(2) As Double

Const ASSE_X = 1
Const ASSE_Y1 = 2
Const ASSE_Y2 = 3

Const PLOT_Y1 = 5
Const PLOT_Y2 = 6


' PARAMETRI
' 1 - Nome File della rampa applicata ( Tempo -> V )
' 2 - Nome File della curva di comando ( V -> mA )
' 3 - Lista delle Variabili da graficare


'---------------------------
' Stato macchina del modulo
'---------------------------
Private Enum eMachineState
    eInit = 0
    eStart = 1
    eRun = 2
    esavefile = 3
    eEnd = 4
    eError = 5
End Enum


'----------------
' PARAMETRI TEST
'----------------
Private Enum eTestParameter

    eNomeFileRampaXY = 1            ' Nome del file della RAMPA XY
    eNomeFileCurvaComando = 2       ' Nome del file della CURVA DI COMANDO
    
    eScheda_Indirizzo_X1 = 3         ' Source address della MMS6252 per la var di comando
    eVar_Tipo_X1 = 4                 ' Tipo della Variabile indipendente di comando
    eVar_Indirizzo_X1 = 5            ' Variabile indipendente di comando

    eScheda_Indirizzo_X2 = 6         ' Source address della MMS6252 per la var di comando
    eVar_Tipo_X2 = 7                 ' Tipo della Variabile indipendente di comando
    eVar_Indirizzo_X2 = 8            ' Variabile indipendente di comando



    eScheda_IndirizzoY1 = 9         ' Source address della MMS6252 per la var di comando
    eVar_Tipo_Y1 = 10                ' Tipo della Variabile dipendente da acquisire
    eVar_Indirizzo_Y1 = 11           ' Variabile dipendente da acquisire

    eScheda_IndirizzoY2 = 12         ' Source address della MMS6252 per la var di comando
    eVar_Tipo_Y2 = 13               ' Tipo della Variabile dipendente da acquisire
    eVar_Indirizzo_Y2 = 14          ' Variabile dipendente da acquisire
    
    eScheda_IndirizzoY3 = 15        ' Source address della MMS6252 per la var di comando
    eVar_Tipo_Y3 = 16               ' Tipo della Variabile dipendente da acquisire
    eVar_Indirizzo_Y3 = 17          ' Variabile dipendente da acquisire
    
    eMax
End Enum


' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagPuntiAcquisizione
    x1 As Double
    x2 As Double
    Y1 As Double
    Y2 As Double
    Y3 As Double
End Type


Public Type TagTEST_ACQUISIZIONE_PRE
    bOldOn As Boolean                               ' difu del comando di start del modulo
    Return As Integer                               ' Stato del modulo
    State As Integer
    NomeFile As String                              ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean                           ' Flag segnalazione primo avvio del programma
    Timer As Long
    
    CURVE_XY As TagChannelRampa                     ' File di configurazione della rampa usata per effettuare l'acquisizione
    CURVE_COMANDO As TagChannel                     ' File di configurazione delle curve comando mA in funzione di V teorico

    
    Punto() As TagPuntiAcquisizione                 ' Punti della acquisizione
    Tempo_sec_start As Double                       ' Istante di avvio della sequenza espresso in secondi
    NumeroPuntiAcquisiti As Long                    ' numero di punti acquisiti
    OldCommandInMilliVolts As Double
    OldTest(eMax) As String                         '(aumenta la dim dell'array perchè ho aggiunto il 2°set di curve limite)
    
    
    Scheda_IndirizzoY1 As Integer
    Scheda_Indirizzo_X1 As Integer
    Scheda_Indirizzo_X2 As Integer
    Scheda_IndirizzoY2 As Integer
    Scheda_IndirizzoY3 As Integer
    
    
    
    
    
    TimeA As Long
    TimeB As Long
    Delta As Long
    DeltaMax As Long
End Type


'----------------
' RampSettings()
'----------------
' Funzione di inizializzazione dei settaggi del grafico. Carica le curve limite
Private Sub RampSettings(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)

    On Error GoTo err:

    Dim bUpdate As Boolean
    Dim i As Integer

    ' RAMPE
    ' -----
    If (Context.OldTest(eNomeFileRampaXY) <> Test.Parameter(eNomeFileRampaXY)) Then

        ' Inizializzazione dei file della rampa da applicare
        If (Update_CURVE_XY(Context.CURVE_XY, Get_Nome_Assoluto_File(Test.Parameter(eNomeFileRampaXY), IndiceSezione, Infos.Label1.Caption)) = True) Then
            Prepare_MsgBox Infos, "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI RAMPE", vbCritical
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eNomeFileRampaXY) = Test.Parameter(eNomeFileRampaXY)
    End If

    ' CURVE COMANDO
    ' -------------
    If (Context.OldTest(eNomeFileCurvaComando) <> Test.Parameter(eNomeFileCurvaComando)) Then
        ' Inizializzazione dei comandi da applicare
        If (Update_CURVECOMANDO(Context.CURVE_COMANDO, Get_Nome_Assoluto_File(Test.Parameter(eNomeFileCurvaComando), IndiceSezione, Infos.Label1.Caption)) = True) Then
            Prepare_MsgBox Infos, "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI COMANDO", vbCritical
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eNomeFileCurvaComando) = Test.Parameter(eNomeFileCurvaComando)
    End If

    Exit Sub

err:
    Debug.Print err.Description
    
End Sub 'RampSettings()


'-------------------------
' "INFO_ACQUISIZIONE_PRE"
'-------------------------
Private Sub INFO_ACQUISIZIONE_PRE(Infos As TagInfosSingleTest, Optional bClear As Boolean = False)

    ' Visualizza GRAFICO
    If (Infos.CWGraph1.Visible <> False) Then
        Infos.CWGraph1.Visible = False
    End If
    
    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
    End If
    

    ' Visualizza TEXT BOX
    If (Infos.RichTextBox1.Visible <> True) Then
        Infos.RichTextBox1.Visible = True
    End If
    
    If (bClear = True) Then
        Infos.RichTextBox1.Text = ""
        Infos.CWGraph1.ClearData
        Infos.CWGraph2.ClearData
    End If

End Sub 'INFO_ACQUISIZIONE_GENERICA()


Private Function TEST_ACQUISIZIONE_PRE_INIT_STRUCT(Context1 As TagTEST_ACQUISIZIONE_PRE)
    Context1.bFirstTime = False
    Context1.bOldOn = False
    Context1.Delta = 0
    Context1.DeltaMax = 0
    Context1.NomeFile = ""
    Context1.NumeroPuntiAcquisiti = 0
    Context1.OldCommandInMilliVolts = 0

    Context1.Return = 0
    Context1.State = 0
    Context1.Tempo_sec_start = 0
    Context1.TimeA = 0
    Context1.TimeB = 0
    Context1.Timer = 0

End Function

'-------------------------
' "TEST_ACQUISIZIONE_PRE"
'-------------------------
Public Sub TEST_ACQUISIZIONE_PRE(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_ACQUISIZIONE_PRE, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "ACQUISIZIONE_PRE") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        INFO_ACQUISIZIONE_PRE Infos
        
        If (bStop = True And bOn = True) Then
        
            TEST_ACQUISIZIONE_PRE_INIT_STRUCT Context1
            INFO_ACQUISIZIONE_PRE Infos, True
            
            
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then        ' fronte di salita del test ( AVVIO )
            Init Test, Infos, IndiceSezione, Context1               ' init delle variabili del modulo
        
        End If
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        Task Test, Infos, IndiceSezione, Context1                   ' Task del modulo
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        
        bReturn = Context1.Return
    End If
    
End Sub 'TEST_ACQUISIZIONE_PRE()


'-----------------
' Task del modulo
'-----------------
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)
    
    Select Case Context.State
        Case eInit: Init Test, Infos, IndiceSezione, Context
        Case eStart: Start Test, Infos, IndiceSezione, Context
        Case eRun: Run Test, Infos, IndiceSezione, Context
        Case esavefile: SaveFile Test, Infos, IndiceSezione, Context
        Case eEnd: Fine Test, Infos, IndiceSezione, Context
        Case eError: ErrorTest Test, Infos, IndiceSezione, Context
        Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
End Sub 'Task()


'------
' Init
'------
' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)
    
    Context.Return = eTEST_BUSY 'eTEST_BUSY_NON_BLOCCANTE
     
    Context.State = eStart
    
    Context.Timer = GetTickCount + SPEGNINENTO_TIMEOUT_MS
    Context.bFirstTime = True
    
    RampSettings Test, Infos, IndiceSezione, Context
    
    Context.NumeroPuntiAcquisiti = 0
    
    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
    
    Infos.RichTextBox1.Text = ""
    Infos.RichTextBox1.BackColor = vbWhite
    
    
    If (CInt(Test.Parameter(eScheda_IndirizzoY1)) = -1) Then
        Context.Scheda_IndirizzoY1 = IndiceSezione + 1
    Else
        Context.Scheda_IndirizzoY1 = CInt(Test.Parameter(eScheda_IndirizzoY1))
    End If
    
    
    If (CInt(Test.Parameter(eScheda_IndirizzoY2)) = -1) Then
        Context.Scheda_IndirizzoY2 = IndiceSezione + 1
    Else
        Context.Scheda_IndirizzoY2 = CInt(Test.Parameter(eScheda_IndirizzoY2))
    End If
    
    
    
    If (CInt(Test.Parameter(eScheda_IndirizzoY3)) = -1) Then
        Context.Scheda_IndirizzoY3 = IndiceSezione + 1
    Else
        Context.Scheda_IndirizzoY3 = CInt(Test.Parameter(eScheda_IndirizzoY3))
    End If
    
        
        
    If (CInt(Test.Parameter(eScheda_Indirizzo_X1)) = -1) Then
        Context.Scheda_Indirizzo_X1 = CInt(Test.Parameter(eScheda_Indirizzo_X1))
    Else
        Context.Scheda_Indirizzo_X1 = CInt(Test.Parameter(eScheda_Indirizzo_X1))
    End If
    
    
    
    If (CInt(Test.Parameter(eScheda_Indirizzo_X2)) = -1) Then
        Context.Scheda_Indirizzo_X2 = IndiceSezione + 1
    Else
        Context.Scheda_Indirizzo_X2 = CInt(Test.Parameter(eScheda_Indirizzo_X2))
    End If
    
        
    Prepare_Textbox_Parametri Test, Infos
    
End Sub 'Init()


'-------
' Start
'-------
Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)
    
    Context.Tempo_sec_start = GetTickCount / 1000
    
    ' INFO
    '------
    Infos.RichTextBox1.Text = "NOME FILE CURVA COMANDO    = " + Test.Parameter(eNomeFileCurvaComando) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "NOME FILE RAMPA APPLICATA  = " + Test.Parameter(eNomeFileRampaXY) + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = ">>> M O D U L O   A C Q U I S I Z I O N E   P R E <<<" + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + ">>> RUN  <<<"
    Context.State = eRun

End Sub 'Start()


'------
' Run()
'------
Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)
    
    Dim Tempo_sec As Double
    Dim CommandInMilliVolts As Double
    Dim State As MltState
    Dim bPunti_Y_Validi As Boolean
    Dim Percentual As Integer
    
    On Error Resume Next
    
    
    If (Context.Timer <= GetTickCount) Then
        
        Context.TimeA = GetTickCount
        
        Tempo_sec = GetTickCount / 1000 - Context.Tempo_sec_start
        Context.Timer = GetTickCount + 1        ' aggiornamento Timer ogni 50 mSec
        
        'trappola
        If (Tempo_sec < 0) Then
            Prepare_MsgBox Infos, "TEMPO NEGATIVO", vbCritical
            Exit Sub
        End If
        
        '---------------
        ' RAMPA "t-->V" (tempo-->volt)
        '---------------
        ' Prelevo il comando dalla rampa applicata
        If (Get_CURVE_XY_VALUE(Context.CURVE_XY, Tempo_sec, CommandInMilliVolts) = eERROR_Rampe) Then
            
            Context.State = esavefile
        
        Else
            SetCalibratedValue 0, "INTERNAL", 1, CommandInMilliVolts
            '--------------------------------
            ' CURVA COMANDO "V-->Percentual"
            '--------------------------------
            'If (Get_CURVECOMANDO_VALUE(Context.CURVE_COMANDO, CommandInMilliVolts / 1000, State, Percentual) = eError_CURVE_COMANDO_CAN) Then
            If (Get_CURVECOMANDO_VALUE(Context.CURVE_COMANDO, CommandInMilliVolts / 1000, State, Percentual) = eError_CURVE_COMANDO_CAN) Then
                
                
                Context.State = eError
            Else
                
                UPDATE_OUTPUT Test, Context, Percentual, State
                
            End If
        End If
        
        
        'If (Context.State <> eError And Context.State <> eSaveFile) Then
        If (Context.State <> esavefile) Then
        
        
            Context.Punto(Context.NumeroPuntiAcquisiti).Y1 = GetCalibratedValue_EVO(Context.Scheda_IndirizzoY1, Test.Parameter(eVar_Tipo_Y1), CInt(Test.Parameter(eVar_Indirizzo_Y1)))
            Context.Punto(Context.NumeroPuntiAcquisiti).Y2 = GetCalibratedValue_EVO(Context.Scheda_IndirizzoY2, Test.Parameter(eVar_Tipo_Y2), CInt(Test.Parameter(eVar_Indirizzo_Y2)))
            Context.Punto(Context.NumeroPuntiAcquisiti).Y3 = GetCalibratedValue_EVO(Context.Scheda_IndirizzoY3, Test.Parameter(eVar_Tipo_Y3), CInt(Test.Parameter(eVar_Indirizzo_Y3)))
            Context.Punto(Context.NumeroPuntiAcquisiti).x1 = GetCalibratedValue_EVO(Context.Scheda_Indirizzo_X1, Test.Parameter(eVar_Tipo_X1), CInt(Test.Parameter(eVar_Indirizzo_X1)))
            Context.Punto(Context.NumeroPuntiAcquisiti).x2 = GetCalibratedValue_EVO(Context.Scheda_Indirizzo_X2, Test.Parameter(eVar_Tipo_X2), CInt(Test.Parameter(eVar_Indirizzo_X2)))
                
           
            If (Context.bFirstTime = True) Then
                Context.bFirstTime = False
                
                Context.Punto(Context.NumeroPuntiAcquisiti).x1 = GetCalibratedValue_EVO(Context.Scheda_Indirizzo_X1, Test.Parameter(eVar_Tipo_X1), CInt(Test.Parameter(eVar_Indirizzo_X1)))
                Context.Punto(Context.NumeroPuntiAcquisiti).x2 = GetCalibratedValue_EVO(Context.Scheda_Indirizzo_X2, Test.Parameter(eVar_Tipo_X2), CInt(Test.Parameter(eVar_Indirizzo_X2)))
            
            End If
                                                        
            'If (bPunti_Y_Validi = True) Then
                Context.NumeroPuntiAcquisiti = Context.NumeroPuntiAcquisiti + 1
                ReDim Preserve Context.Punto(Context.NumeroPuntiAcquisiti)
            'End If

        End If 'END (Context.State <> eSaveFile)
        
        Context.OldCommandInMilliVolts = CommandInMilliVolts
        
        Context.TimeB = GetTickCount
        
        Context.Delta = Context.TimeB - Context.TimeA
        
        If (Context.DeltaMax < Context.Delta) Then
            Context.DeltaMax = Context.Delta
        End If
        
    End If '(Context.Timer < GetTickCount)
    
End Sub 'Run()


Private Sub UPDATE_OUTPUT(ByRef Test As TagSingoloTest, Context As TagTEST_ACQUISIZIONE_PRE, ByVal Percentual As Integer, ByVal State As Integer)


    SetCalibratedValue CInt(Test.Parameter(eScheda_Indirizzo_X1)), Test.Parameter(eVar_Tipo_X1), Test.Parameter(eVar_Tipo_X1), Percentual
    SetCalibratedValue CInt(Test.Parameter(eScheda_Indirizzo_X2)), Test.Parameter(eVar_Tipo_X2), Test.Parameter(eVar_Tipo_X2), State
    
    
    
End Sub 'UPDATE_OUTPUT()


'-----------
' SaveFile()
'-----------
Private Sub SaveFile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)

    ' INFO
    '------
    Infos.RichTextBox1.Text = vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "Salvataggio Curva Acquisizione X-Y in File.csv" + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = vbLf + Infos.RichTextBox1.Text
                

    SalvaPunti Test, Infos, Context, IndiceSezione
    Context.State = eEnd
    
    
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    
End Sub 'SaveFile()


'------
' Fine
'------
Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)

    If (Infos.CWGraph1.PlotAreaColor = vbRed) Then
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
    
End Sub 'Fine()


'------------
' ErrorTest()
'------------
Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE)
    
    Infos.RichTextBox1.BackColor = vbRed
    
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = " ! ! ! ! ! ERRORE SUL TEST ! ! ! ! ! " + vbLf + Infos.RichTextBox1.Text
    End If
    
    Context.Return = eTEST_ERROR

End Sub 'ErrorTest()


'-------------
' SalvaPunti()
'-------------
Private Sub SalvaPunti(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_ACQUISIZIONE_PRE, IndiceSezione As Integer)
    
    Dim FileName As String
    Dim FileName_Post As String
    Dim bSaveLastAcq As Boolean
    Dim PathDirectory As String
    Dim PathDirectoryAll As String
    Dim i As Long
    Dim ora As String
    Dim oggi As String
    
On Error GoTo err:
    
    ora = Format(Time, ("HH.MM.SS"))
    ora = Replace(ora, ".", "")

    oggi = Format(Date, ("yy/mm/dd"))
    oggi = Replace(oggi, "/", "")
    
    PathDirectory = CreaDirectory(Infos, Settings.FolderGraphSaved)
    PathDirectoryAll = CreaDirectory(Infos, Settings.FolderGraphSaved_EVO)
    
    If (PathDirectory = "") Then
        PathDirectory = Settings.FolderGraphSaved + "\"
    End If
    If (PathDirectoryAll = "") Then
        PathDirectoryAll = Settings.FolderGraphSaved_EVO + "\"
    End If
    
    ' PERCORSO SALVATAGGIO FILE
    '--------------------------
    ' se presente ID Seriale Valvola
    If (Infos.Label1 <> "") Then
        'FileName = PathDirectory + Test.Name + "_" + Infos.Label1.Caption + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
        FileName = PathDirectoryAll + Test.Name + "_" + contextd.DistrSelected + "_" + Infos.Label1.Caption + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
        FileName_Post = PathDirectory + Test.Name + ".csv"
    
    ' se non presente ID Seriale Valvola
    Else
        'FileName = PathDirectory + Test.Name + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
        FileName = PathDirectoryAll + Test.Name + "_" + contextd.DistrSelected + "_" + oggi + "_" + ora + "_" + CStr(IndiceSezione) + ".csv"
        FileName_Post = PathDirectory + Test.Name + ".csv"
    End If
    
    ' INFO ------------------------------------------------------------------------
    Infos.RichTextBox1.Text = "FILE = " + FileName + vbLf + Infos.RichTextBox1.Text
    '---------------------------------------------------------------------------------------
    
    On Error Resume Next
    Close #22
    Close #23
    
    Open FileName For Append As #22
    Open FileName_Post For Output As #23
    
    Print #22, "X1 ; X2 ; Y1 ; Y2; Y3"
    Print #23, "X1 ; X2 ; Y1 ; Y2; Y3"
    Do While (i < UBound(Context.Punto))
        Print #22, CStr(Context.Punto(i).x1) + ";" + CStr(Context.Punto(i).x2) + ";" + CStr(Context.Punto(i).Y1) + ";" + CStr(Context.Punto(i).Y2) + ";" + CStr(Context.Punto(i).Y3)
        Print #23, CStr(Context.Punto(i).x1) + ";" + CStr(Context.Punto(i).x2) + ";" + CStr(Context.Punto(i).Y1) + ";" + CStr(Context.Punto(i).Y2) + ";" + CStr(Context.Punto(i).Y3)
        i = i + 1
    Loop

    Close #22
    Close #23
    
    ' Voglio salvare l'ultima acq
    bSaveLastAcq = True
    
    If (Infos.Label1.Caption = "") Then
            bSaveLastAcq = False
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
        Print #22, "X1 ; X2 ; Y1 ; Y2; Y3"
        Print #22, "X1 ; X2 ; Y1 ; Y2; Y3"
        Do While (i < UBound(Context.Punto))
            Print #22, CStr(Context.Punto(i).x1) + ";" + CStr(Context.Punto(i).x2) + ";" + CStr(Context.Punto(i).Y1) + ";" + CStr(Context.Punto(i).Y2) + ";" + CStr(Context.Punto(i).Y2)
            i = i + 1
        Loop
    
        Close #22
    End If
    
Exit Sub

err:
    Prepare_MsgBox Infos, "ERRORE DURANTE IL SALVATAGGIO DELLA CURVA", vbCritical
    
End Sub 'SalvaPunti()



Private Function CreaDirectory(Infos As TagInfosSingleTest, ByVal StartDirectory As String) As String
    
    On Error GoTo err:
    
    Dim PathDirectory  As String
        
    CreaDirectory = StartDirectory + "\" + Infos.codicePRodotto + "\"
    
    MkDir (CreaDirectory)
    Exit Function
    
err:

End Function 'CreaDirectory()



Public Function DirExists(ByVal Path As String) As Boolean

    On Error Resume Next
    
    'Legge l'attributo e si assicura che si tratti di una directory
    FileExists = GetAttr(Path) And vbDirectory
    
    'Se avviene un errore la Function restituisce False
    
End Function 'DirExists()



Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFileRampaXY                      = " + Test.Parameter(eNomeFileRampaXY) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFileCurvaComando                      = " + Test.Parameter(eNomeFileCurvaComando) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eScheda_Indirizzo_X1               = " + Test.Parameter(eScheda_Indirizzo_X1) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Tipo_X1        = " + Test.Parameter(eVar_Tipo_X1) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Indirizzo_X1   = " + Test.Parameter(eVar_Indirizzo_X1) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eScheda_Indirizzo_X2   = " + Test.Parameter(eScheda_Indirizzo_X2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Tipo_X2 = " + Test.Parameter(eVar_Tipo_X2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Indirizzo_X2 = " + Test.Parameter(eVar_Indirizzo_X2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eScheda_IndirizzoY1                  = " + Test.Parameter(eScheda_IndirizzoY1) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Tipo_Y1                  = " + Test.Parameter(eVar_Tipo_Y1) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Indirizzo_Y1                     = " + Test.Parameter(eVar_Indirizzo_Y1) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eScheda_IndirizzoY2                 = " + Test.Parameter(eScheda_IndirizzoY2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Tipo_Y2                 = " + Test.Parameter(eVar_Tipo_Y2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Indirizzo_Y2                 = " + Test.Parameter(eVar_Indirizzo_Y2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eScheda_IndirizzoY3                 = " + Test.Parameter(eScheda_IndirizzoY3) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Tipo_Y3                 = " + Test.Parameter(eVar_Tipo_Y3) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eVar_Indirizzo_Y3                 = " + Test.Parameter(eVar_Indirizzo_Y3) + vbCrLf


End Sub


