Attribute VB_Name = "Module_TEST_ACQUISIZIONE_GENERICA"

'**************************************************************
'       M O D U L O :    Module_TEST_ACQUISIZIONE_GENERICA.bas
'**************************************************************

Private Declare Function GetTickCount Lib "kernel32" () As Long

Private Const SPEGNINENTO_TIMEOUT_MS = 200
Private Const ACCENSIONE_TIMEOUT_MS = 200

Private y(2) As Double

Const ASSE_X = 1
Const ASSE_Y1 = 2
Const ASSE_Y2 = 3

Const PLOT_Y1 = 5
Const PLOT_Y2 = 6


' 1 - Nome File della rampa applicata ( Tempo -> V )
' 2 - Nome File della curva di comando ( V -> mA )
' 3 - Lista delle Variabili da graficare
' 4 - Nome File del limite inferiore
' 5 - Nome File del limite superiore


'---------------------------
' Stato macchina del modulo
'---------------------------
Private Enum eMachineState
    eInit = 0
    eStart = 1
    eRun = 2
    eSaveFile = 3
    eEnd = 4
    eError = 5
End Enum


'----------------
' PARAMETRI TEST
'----------------
Private Enum eTestParameter
    
    eNomeFileRampaXY = 1                    ' Nome del file della RAMPA XY
    eNomeFileCurvaComando = 2               ' Nome del file della CURVA DI COMANDO
    
    eIndirizzoScheda = 3                    ' Source address della MMS6252 per la var di comando
    eTipoVarIndipendenteXY = 4              ' Tipo della Variabile indipendente di comando
    eVarIndipendenteXY = 5                  ' Variabile indipendente di comando
    
    NomeFileCurvaLimiteInferiore = 6        ' Nome file Curva LIMITE INFERIORE
    NomeFileCurvaLimiteSuperiore = 7        ' Nome file Curva LIMITE SUPERIORE
    eNomeFileCWGraphSezione = 8             ' Nome file Impostazioni GRAFICO CWGraph
    NomeFileCurvaLimiteInferiore_2 = 9      ' Nome file Curva LIMITE INFERIORE
    NomeFileCurvaLimiteSuperiore_2 = 10     ' Nome file Curva LIMITE SUPERIORE
    
    eMax
End Enum


' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagPuntiAcquisizione
    x As Double
    y As Double
    Z As Double
End Type

Public Type TagTEST_ACQUISIZIONE_GENERICA
    bOldOn As Boolean                               ' difu del comando di start del modulo
    Return As Integer                               ' Stato del modulo
    State As Integer
    NomeFile As String                              ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean                           ' Flag segnalazione primo avvio del programma
    Timer As Long
    
    CURVE_LIMITE(2) As CurveLimitePlusTag           ' Struttura dati delle curve limite
    CURVE_XY As TagChannelRampa                     ' File di configurazione della rampa usata per effettuare l'acquisizione
    CURVE_COMANDO As TagChannel                     ' File di configurazione delle curve comando mA in funzione di V teorico
    
    IMPOSTAZ_CWGRAPH() As TagImpostazioniGrafico    ' File di configurazione per il Grafico "CWGraph"
    
    Punto() As TagPuntiAcquisizione                 ' Punti della acquisizione
    Tempo_sec_start As Double                       ' Istante di avvio della sequenza espresso in secondi
    NumeroPuntiAcquisiti As Long                    ' numero di punti acquisiti
    OldCommandInMilliVolts As Double
    
    OldTest(eMax) As String   '(aumenta la dim dell'array perchè ho aggiunto il 2°set di curve limite)
End Type


'-------------------------------------------------------
' Questa funzione imposta gli oggetti grafici "CWGraph1"
'-------------------------------------------------------
Public Function ImpostaCWGraphLoad_GRAPH(ByRef CWGraph1 As CWGraph, ByRef Context As TagTEST_ACQUISIZIONE_GENERICA)

'    On Error GoTo err:

    ' GRAFICO
    '--------
    ' asse X
    ' (1° riga dati del file di impostazione dei CWGraph1)
    CWGraph1.Axes(ASSE_X).Caption = Context.IMPOSTAZ_CWGRAPH(0).NomeVariabile                ' "NOME"
    CWGraph1.Axes(ASSE_X).Minimum = Context.IMPOSTAZ_CWGRAPH(0).MinVariabile                 ' "MIN"
    CWGraph1.Axes(ASSE_X).Maximum = Context.IMPOSTAZ_CWGRAPH(0).MaxVariabile                 ' "MAX"
    CWGraph1.Axes(ASSE_X).Ticks.MajorUnitsInterval = Context.IMPOSTAZ_CWGRAPH(0).MajorTick   ' "MAJOR TICK"
    CWGraph1.Axes(ASSE_X).Ticks.MinorUnitsInterval = Context.IMPOSTAZ_CWGRAPH(0).MinorTick   ' "MINOR TICK"
    ' asse Y1
    ' (2° riga dati del file di impostazione dei CWGraph1)
    CWGraph1.Axes(ASSE_Y1).Caption = Context.IMPOSTAZ_CWGRAPH(1).NomeVariabile                ' "NOME"
    CWGraph1.Axes(ASSE_Y1).Minimum = Context.IMPOSTAZ_CWGRAPH(1).MinVariabile                 ' "MIN"
    CWGraph1.Axes(ASSE_Y1).Maximum = Context.IMPOSTAZ_CWGRAPH(1).MaxVariabile                 ' "MAX"
    CWGraph1.Axes(ASSE_Y1).Ticks.MajorUnitsInterval = Context.IMPOSTAZ_CWGRAPH(1).MajorTick   ' "MAJOR TICK"
    CWGraph1.Axes(ASSE_Y1).Ticks.MinorUnitsInterval = Context.IMPOSTAZ_CWGRAPH(1).MinorTick   ' "MINOR TICK"
    ' asse Y2 (opzionale)
    ' (3° riga dati del file di impostazione dei CWGraph1)
    CWGraph1.Axes(ASSE_Y2).Caption = Context.IMPOSTAZ_CWGRAPH(2).NomeVariabile                ' "NOME"
    CWGraph1.Axes(ASSE_Y2).Minimum = Context.IMPOSTAZ_CWGRAPH(2).MinVariabile                 ' "MIN"
    CWGraph1.Axes(ASSE_Y2).Maximum = Context.IMPOSTAZ_CWGRAPH(2).MaxVariabile                 ' "MAX"
    CWGraph1.Axes(ASSE_Y2).Ticks.MajorUnitsInterval = Context.IMPOSTAZ_CWGRAPH(2).MajorTick   ' "MAJOR TICK"
    CWGraph1.Axes(ASSE_Y2).Ticks.MinorUnitsInterval = Context.IMPOSTAZ_CWGRAPH(2).MinorTick   ' "MINOR TICK"

'err:

End Function 'ImpostaCWGraphLoad_GRAPH()


'-------------------------------------------------------
' Questa funzione plotta i punti nei grafici "CWGraph1"
'-------------------------------------------------------
Public Sub Graph_Points(ByRef CWGraph1 As CWGraph, ByRef Context As TagTEST_ACQUISIZIONE_GENERICA, ByRef ulPointIndex As Long, ByVal ClearData As Boolean)
    
    Dim i As Long
    Dim TempY1 As Double
    Dim TempY2 As Double
    Dim TempX As Double
    Dim DeltaError As Double
    
    If (ClearData = True) Then
        CWGraph1.Plots(PLOT_Y1).ClearData
        CWGraph1.Plots(PLOT_Y2).ClearData
    End If
    
    CWGraph1.Plots(PLOT_Y1).Visible = False
    CWGraph1.Plots(PLOT_Y2).Visible = False
    
    i = ulPointIndex
    Do While (i < UBound(Context.Punto))
        If (i < 2) Then
            CWGraph1.Plots(PLOT_Y1).ChartXvsY Context.Punto(i).x, Context.Punto(i).y
            CWGraph1.Plots(PLOT_Y2).ChartXvsY Context.Punto(i).x, Context.Punto(i).Z
        Else
            
            TempY2 = (Context.Punto(i - 1).Z)
            
            TempY1 = (Context.Punto(i - 1).y)
            TempX = Context.Punto(i).x
           
            If (TempY1 > Context.Punto(i).y) Then
                DeltaError = TempY1 - Context.Punto(i).y
            Else
                DeltaError = Context.Punto(i).y - TempY1
            End If
            
            If (DeltaError > 0.5) Then
                TempY1 = Context.Punto(i).y
                TempX = Context.Punto(i).x
            End If
            
            CWGraph1.Plots(PLOT_Y1).ChartXvsY Context.Punto(i).x, TempY1
        End If
        
        i = i + 1
    Loop
    
    ulPointIndex = UBound(Context.Punto)
    CWGraph1.Plots(PLOT_Y1).Visible = True
    
End Sub 'Graph_Points()


'----------------
' GraphSettings()
'----------------
' Funzione di inizializzazione dei settaggi del grafico. Carica le curve limite / Impostazioni di CWGraph
Private Sub GraphSettings(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)

    On Error GoTo err:
    
    Dim bUpdate As Boolean
    Dim i As Integer
    Dim par As CurveLimitePlusParametersTag
    

    ' inizializzazione delle curve limite
    par.Indice(eInferiore) = 1
    par.Indice(eSuperiore) = 2

    CurveLimitePlusInit Context.CURVE_LIMITE(0), par
    CurveLimitePlusInit Context.CURVE_LIMITE(1), par


    ' RAMPE
    ' -----
    If (Context.OldTest(eNomeFileRampaXY) <> Test.Parameter(eNomeFileRampaXY)) Then

        ' Inizializzazione dei file della rampa da applicare
        If (Update_CURVE_XY(Context.CURVE_XY, Settings.FolderRampeXY + "\" + Test.Parameter(eNomeFileRampaXY)) = True) Then
            MsgBox "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI RAMPE", vbCritical
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eNomeFileRampaXY) = Test.Parameter(eNomeFileRampaXY)
    End If

    ' CURVE COMANDO
    ' -------------
    If (Context.OldTest(eNomeFileCurvaComando) <> Test.Parameter(eNomeFileCurvaComando)) Then
        ' Inizializzazione dei comandi da applicare
        If (Update_CURVECOMANDO(Context.CURVE_COMANDO, Settings.FolderFileCurveComando + "\" + Test.Parameter(eNomeFileCurvaComando)) = True) Then
            MsgBox "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI COMANDO", vbCritical
            Context.State = eError
        End If

        bUpdate = True
        Context.OldTest(eNomeFileCurvaComando) = Test.Parameter(eNomeFileCurvaComando)
    End If

    ' CURVE LIMITE INF e SUP
    ' ----------------------
    If (Context.OldTest(NomeFileCurvaLimiteInferiore) <> Test.Parameter(NomeFileCurvaLimiteInferiore) Or _
        Context.OldTest(NomeFileCurvaLimiteSuperiore) <> Test.Parameter(NomeFileCurvaLimiteSuperiore)) Then
    
        Infos.CWGraph1.ClearData
    End If

    ' curva limite inf
    If (Context.OldTest(NomeFileCurvaLimiteInferiore) <> Test.Parameter(NomeFileCurvaLimiteInferiore)) Then
    
        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(0), Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteInferiore), eInferiore) = eERROR_CurveLimite) Then
            MsgBox "ERRORE CARICAMENTO CURVE LIMITE INFERIORE", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph1, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteInferiore), eInferiore
        End If


        Context.OldTest(NomeFileCurvaLimiteInferiore) = Test.Parameter(NomeFileCurvaLimiteInferiore)
        bUpdate = True
    End If


    ' curva limite inf
    If (Context.OldTest(NomeFileCurvaLimiteInferiore_2) <> Test.Parameter(NomeFileCurvaLimiteInferiore_2)) Then
    
        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(1), Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteInferiore_2), eInferiore) = eERROR_CurveLimite) Then
            MsgBox "ERRORE CARICAMENTO CURVE LIMITE INFERIORE 2", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph1, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteInferiore_2), eInferiore_2
        End If


        Context.OldTest(NomeFileCurvaLimiteInferiore_2) = Test.Parameter(NomeFileCurvaLimiteInferiore_2)
        bUpdate = True
    End If


    ' curva limite sup
    If (Context.OldTest(NomeFileCurvaLimiteSuperiore) <> Test.Parameter(NomeFileCurvaLimiteSuperiore)) Then

        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(0), Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteSuperiore), eSuperiore) = eERROR_CurveLimite) Then
            MsgBox "ERRORE CARICAMENTO CURVE LIMITE SUPERIORE 1", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph1, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteSuperiore), eSuperiore
        End If


        Context.OldTest(NomeFileCurvaLimiteSuperiore) = Test.Parameter(NomeFileCurvaLimiteSuperiore)
        bUpdate = True
    End If
    
    ' curva limite sup
    If (Context.OldTest(NomeFileCurvaLimiteSuperiore_2) <> Test.Parameter(NomeFileCurvaLimiteSuperiore_2)) Then

        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(1), Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteSuperiore_2), eSuperiore) = eERROR_CurveLimite) Then
            MsgBox "ERRORE CARICAMENTO CURVE LIMITE SUPERIORE 2", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph1, Settings.FolderFileCurveLimite + "\" + Test.Parameter(NomeFileCurvaLimiteSuperiore_2), eSuperiore_2
        End If


        Context.OldTest(NomeFileCurvaLimiteSuperiore_2) = Test.Parameter(NomeFileCurvaLimiteSuperiore_2)
        bUpdate = True
    End If
    

    ' IMPOSTAZIONI "CWGraph1"
    ' ----------------------
    If (Context.OldTest(eNomeFileCWGraphSezione) <> Test.Parameter(eNomeFileCWGraphSezione)) Then
            
        If (ImpostaCWGraphLoad(Context.IMPOSTAZ_CWGRAPH(), Settings.FolderConfigurazioneTest + "\" + Test.Parameter(eNomeFileCWGraphSezione))) Then
            
            ImpostaCWGraphLoad_GRAPH Infos.CWGraph1, Context
            
        Else
            MsgBox "ERRORE CARICAMENTO CONFIGURAZIONE GRAFICO SEZIONE", vbCritical
        End If
        
        Context.OldTest(eNomeFileCWGraphSezione) = Test.Parameter(eNomeFileCWGraphSezione)
        bUpdate = True
        
    End If
    
    'If (bUpdate = True) Then
        Infos.CWGraph1.PlotAreaColor = vbWhite
        Infos.CWGraph1.Plots(PLOT_Y1).PointColor = vbBlue
'        Infos.CWGraph1.Plots(PLOT_Y1).LineColor = vbBlue
    'End If

    Exit Sub

err:
    Debug.Print err.Description
    
End Sub 'GraphSettings()


'------------------------------
' "INFO_ACQUISIZIONE_GENERICA"
'------------------------------
Private Sub INFO_ACQUISIZIONE_GENERICA(Infos As TagInfosSingleTest)

    If (Infos.CWGraph1.Visible <> True) Then
        Infos.CWGraph1.Visible = True
    End If
    
    If (Infos.RichTextBox1.Visible = True) Then
        Infos.RichTextBox1.Visible = False
    End If
    
End Sub 'INFO_ACQUISIZIONE_GENERICA()



Private Function TEST_ACQUISIZIONE_GENERICA_INIT_STRUCT(Context1 As TagTEST_ACQUISIZIONE_GENERICA)

    Context1.bFirstTime = False
    Context1.bOldOn = False
    Context1.NomeFile = ""
    Context1.NumeroPuntiAcquisiti = 0
    Context1.OldCommandInMilliVolts = 0

    Context1.Return = 0
    Context1.State = 0
    Context1.Tempo_sec_start = 0
    Context1.Timer = 0

End Function

'------------------------------
' "TEST_ACQUISIZIONE_GENERICA"
'------------------------------
Public Sub TEST_ACQUISIZIONE_GENERICA(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_ACQUISIZIONE_GENERICA, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "ACQUISIZIONE_GENERICA") Then
        Context1.Return = eTEST_STOP
        
    Else
                
        INFO_ACQUISIZIONE_GENERICA Infos
        
        
        If (bStop = True And bOn = True) Then
        
            TEST_ACQUISIZIONE_GENERICA_INIT_STRUCT Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            Context1.State = eError
            
        ElseIf (bStop = True) Then
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then        ' fronte di salita del test ( AVVIO )
            Init Test, Infos, IndiceSezione, Context1               ' init delle variabili del modulo
        
        End If
        
        Task Test, Infos, IndiceSezione, Context1                   ' Task del modulo
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        bReturn = Context1.Return
    End If
    
    
    
    
   
    
End Sub 'TEST_ACQUISIZIONE_GENERICA()


'-----------------
' Task del modulo
'-----------------
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)
    
    Select Case Context.State
        Case eInit: Init Test, Infos, IndiceSezione, Context
        Case eStart: Start Test, Infos, IndiceSezione, Context
        Case eRun: Run Test, Infos, IndiceSezione, Context
        Case eSaveFile: SaveFile Test, Infos, IndiceSezione, Context
        Case eEnd: Fine Test, Infos, IndiceSezione, Context
        Case eError: ErrorTest Test, Infos, IndiceSezione, Context
        Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
End Sub 'Task()


'------
' Init
'------
' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)
    
    Context.Return = eTEST_BUSY 'eTEST_BUSY_NON_BLOCCANTE
     
    Context.State = eStart
    
    Context.Timer = GetTickCount + SPEGNINENTO_TIMEOUT_MS
    Context.bFirstTime = True
    
    GraphSettings Test, Infos, IndiceSezione, Context
    
    Context.NumeroPuntiAcquisiti = 0
    
    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
    
End Sub 'Init()


'-------
' Start
'-------
Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)
    
    Context.State = eRun
    Context.Tempo_sec_start = GetTickCount / 1000
    
    'SPOSTARE TUTTI I RIFERIMENTI AL GRAPH NEL RICHT TEXT
    Infos.CWGraph1.Plots(PLOT_Y1).ClearData
    Infos.CWGraph1.Plots(PLOT_Y2).ClearData
    
    ' LOG
    '-----
    AggiungiLogCh IndiceSezione, "| MODULO: POST_PROCESSING| Inizio ---------------------- "

End Sub 'Start()


'------
' Run()
'------
Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)
    
    Dim Tempo_sec As Double
    Dim CommandInMilliVolts As Double
    Dim State As MltState
    Dim bPunti_Y_Validi As Boolean
    Dim Percentual As Integer
    
    On Error Resume Next
    
    
    If (Context.Timer < GetTickCount) Then
        
        Tempo_sec = GetTickCount / 1000 - Context.Tempo_sec_start
        Context.Timer = GetTickCount + 50        ' aggiornamento Timer ogni 50 mSec
        
        '---------------
        ' RAMPA "t-->V" (tempo-->volt)
        '---------------
        ' Prelevo il comando dalla rampa applicata
        If (Get_CURVE_XY_VALUE(Context.CURVE_XY, Tempo_sec, CommandInMilliVolts) = eERROR_Rampe) Then
            
            Context.State = eSaveFile
        Else
            '--------------------------------
            ' CURVA COMANDO "V-->Percentual"
            '--------------------------------
            'If (Get_CURVECOMANDO_VALUE(Context.CURVE_COMANDO, CommandInMilliVolts / 1000, State, Percentual) = eError_CURVE_COMANDO_CAN) Then
            If (Get_CURVECOMANDO_XY_VALUE(Context.CURVE_COMANDO, CommandInMilliVolts / 1000, State, Percentual) = eError_CURVE_COMANDO_CAN) Then
                
                ' LOG
                '-----
                AggiungiLogCh IndiceSezione, "| MODULO: POST_PROCESSING| ESITO = >>> F A I L <<< | ERROR CARICAMENTO CURVA COMANDO 'V-->I' (volt-->mA)"
                                              
                Context.State = eError
            Else
                
                UPDATE_OUTPUT Test, Context, Percentual
                
            End If
        End If
        
        
        'If (Context.State <> eError And Context.State <> eSaveFile) Then
        If (Context.State <> eSaveFile) Then
        
            Context.Punto(Context.NumeroPuntiAcquisiti).y = GetCalibratedValue(&HA0, "AIN16", Context.IMPOSTAZ_CWGRAPH(1).AddrVariabile)
            Context.Punto(Context.NumeroPuntiAcquisiti).Z = GetCalibratedValue(&HA0, "AIN16", Context.IMPOSTAZ_CWGRAPH(2).AddrVariabile)
           
            If (Context.bFirstTime = True) Then
                Context.bFirstTime = False
                
                'Context.Punto(Context.NumeroPuntiAcquisiti).x = CommandInMilliVolts
                Context.Punto(Context.NumeroPuntiAcquisiti).x = GetCalibratedValue(&HA0, "AIN16", Context.IMPOSTAZ_CWGRAPH(0).AddrVariabile)
            
            End If
            
            '--------------------
            ' Carica CURVE LIMITE
            '--------------------
            If ((CurveLimitePlusCheck(Context.CURVE_LIMITE(0), _
                                          Context.Punto(Context.NumeroPuntiAcquisiti).x, _
                                          Context.Punto(Context.NumeroPuntiAcquisiti).y) = eERROR_CurveLimite) And _
                (CurveLimitePlusCheck(Context.CURVE_LIMITE(1), _
                                          Context.Punto(Context.NumeroPuntiAcquisiti).x, _
                                          Context.Punto(Context.NumeroPuntiAcquisiti).y) = eERROR_CurveLimite) _
                                          ) Then
                If (Infos.CWGraph1.PlotAreaColor <> vbRed) Then
                    Distributore.Text2 = Context.Punto(Context.NumeroPuntiAcquisiti).x
                    Distributore.Text1 = Context.Punto(Context.NumeroPuntiAcquisiti).y
                    Infos.CWGraph1.PlotAreaColor = vbRed    ' se l'acquisizione esce dalla curve limite lo sfondo del grafico diventa rosso
                End If
            End If

            '-------------------
            ' PLOTTO su CWGraph1
            '-------------------
            Infos.CWGraph1.Plots(PLOT_Y1).ChartXvsY Context.Punto(Context.NumeroPuntiAcquisiti).x, Context.Punto(Context.NumeroPuntiAcquisiti).y

            'If (bPunti_Y_Validi = True) Then
                Context.NumeroPuntiAcquisiti = Context.NumeroPuntiAcquisiti + 1
                ReDim Preserve Context.Punto(Context.NumeroPuntiAcquisiti)
            'End If


            Context.Punto(Context.NumeroPuntiAcquisiti).x = GetCalibratedValue(&HA0, "AIN16", Context.IMPOSTAZ_CWGRAPH(0).AddrVariabile)
            'Context.Punto(Context.NumeroPuntiAcquisiti).x = Percentual


        End If 'END (Context.State <> eSaveFile)
        
        Context.OldCommandInMilliVolts = CommandInMilliVolts
        
    End If '(Context.Timer < GetTickCount)
    
End Sub 'Run()


Private Sub UPDATE_OUTPUT(ByRef Test As TagSingoloTest, Context As TagTEST_ACQUISIZIONE_GENERICA, ByVal Percentual As Integer)

    Dim TIPO_VARIABILE_X As String
    Dim INDIRIZZO_VARIABILE_X As Integer
    Dim PPDO_IndexMessage As Integer
    Dim PPDO_AdcIndexIntoMessage As Integer
    Dim Output As Double
    
    TIPO_VARIABILE_X = Test.Parameter(eTipoVarIndipendenteXY)
    INDIRIZZO_VARIABILE_X = CInt(Test.Parameter(eVarIndipendenteXY))

    SetCalibratedValue CInt(Test.Parameter(eIndirizzoScheda)), Test.Parameter(eTipoVarIndipendenteXY), INDIRIZZO_VARIABILE_X, Percentual
    
End Sub 'UPDATE_OUTPUT()


'-----------
' SaveFile()
'-----------
Private Sub SaveFile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)

    ' LOG
    '-----
    AggiungiLogCh IndiceSezione, "| MODULO: POST_PROCESSING| SALVATAGGIO CURVA ACQUISIZIONE X-Y"
                

    SalvaPunti Test, Infos, Context, IndiceSezione
    Context.State = eEnd
    
End Sub 'SaveFile()


'------
' Fine
'------
Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)

    If (Infos.CWGraph1.PlotAreaColor = vbRed) Then
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
    
End Sub 'Fine()


'------------
' ErrorTest()
'------------
Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_GENERICA)
    
    Infos.RichTextBox1.BackColor = vbRed
    
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = " ! ! ! ! ! ERRORE SUL TEST ! ! ! ! ! " + vbLf + Infos.RichTextBox1.Text
    End If
    
    Context.Return = eTEST_ERROR

End Sub 'ErrorTest()


'-------------
' SalvaPunti()
'-------------
Private Sub SalvaPunti(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_ACQUISIZIONE_GENERICA, IndiceSezione As Integer)
    
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
    PathDirectoryAll = CreaDirectory(Infos, Settings.FolderGraphSaved_AllACQ)
    
    If (PathDirectory = "") Then
        PathDirectory = Settings.FolderGraphSaved + "\"
    End If
    If (PathDirectoryAll = "") Then
        PathDirectoryAll = Settings.FolderGraphSaved_AllACQ + "\"
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
    ' LOG ----------------------------------------------------------------------------------
    AggiungiLogCh IndiceSezione, "| MODULO: POST_PROCESSING| FILE = " & FileName
    AggiungiLogCh IndiceSezione, "| MODULO: POST_PROCESSING| Fine ------------------------ "
    '---------------------------------------------------------------------------------------
    
    On Error Resume Next
    Close #22
    Close #23
    
    Open FileName For Append As #22
    Open FileName_Post For Output As #23
    
    Print #22, "ASSE X ; ASSE Y ; ASSE Y2"
    Print #23, "X ; Y1 ; Y2"
    Do While (i < UBound(Context.Punto))
        Print #22, CStr(Context.Punto(i).x) + ";" + CStr(Context.Punto(i).y) + ";"; CStr(Context.Punto(i).Z)
        Print #23, CStr(Context.Punto(i).x) + ";" + CStr(Context.Punto(i).y) + ";" + CStr(Context.Punto(i).Z)
        i = i + 1
    Loop

    Close #22
    Close #23
    
    ' Voglio salvare l'ultima acq
    bSaveLastAcq = True
    
    If (Infos.Label1.Caption = "") Then
'ddd
'        If (MsgBox("Attenzione non è stata associata nessun codice al modulo, voler procedere al salvataggio della curva ?", vbYesNo) = vbYes) Then
'            Infos.Label1.Caption = Immissione_SerialNumber()
'        Else
            ' Non voglio salvare l'ultima acq
            bSaveLastAcq = False
'        End If
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
        Print #22, "X ; Y1 ; Y2; Y3"
        Do While (i < UBound(Context.Punto))
            Print #22, CStr(Context.Punto(i).x) + ";" + CStr(Context.Punto(i).y) + ";" + CStr(Context.Punto(i).Z)
            i = i + 1
        Loop
    
        Close #22
    End If
    
Exit Sub

err:
    MsgBox "ERRORE DURANTE IL SALVATAGGIO DELLA CURVA", vbCritical
    
End Sub 'SalvaPunti()



Private Function CreaDirectory(Infos As TagInfosSingleTest, ByVal StartDirectory As String) As String
    
    On Error GoTo err:
    
    Dim PathDirectory  As String
        
    CreaDirectory = StartDirectory + "\" + Infos.CodiceProdotto + "\"
    
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

End Function 'CreaDirectory()



Public Function DirExists(ByVal Path As String) As Boolean

    On Error Resume Next
    'Legge l'attributo e si assicura che si tratti di una directory
    FileExists = GetAttr(Path) And vbDirectory
    
    'Se avviene un errore la Function restituisce False
End Function 'DirExists()

