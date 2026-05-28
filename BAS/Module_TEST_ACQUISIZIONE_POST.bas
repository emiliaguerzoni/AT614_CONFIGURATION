Attribute VB_Name = "Module_TEST_ACQUISIZIONE_POST"
 '**************************************************************
'       M O D U L O :    Module_TEST_ACQUISIZIONE_POST.bas
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
' 1 - Nome File della rampa applicata ( in V vs Tempo )
' 2 - Nome File della curva di comando
' 3 - Lista delle Variabili da graficare
' 4 - Nome File del limite inferiore
' 5 - Nome File del limite superiore


'---------------------------d b
' Stato macchina del modulo
'---------------------------
'Private Enum eMachineState
'    eInit = 0
'    eStart = 1
'    eRun = 2
'    eEnd = 3
'    eError = 4
'End Enum
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
    eColonnaX = 1                           ' Tipo della Variabile indipendente di comando
    eColonnaY = 2                           ' Variabile indipendente di comando
    
    NomeFileIngresso = 3                    ' NomeFile.csv dell'ACQUISIZIONE XY
    
    eNomeFileCWGraphSezione = 4             ' Nome file Impostazioni GRAFICO CWGraph
    
    NomeFileCurvaLimiteInferiore = 5        ' Nome file Curva LIMITE INFERIORE
    NomeFileCurvaLimiteSuperiore = 6        ' Nome file Curva LIMITE SUPERIORE
    
    NomeFileCurvaLimiteInferiore_2 = 7      ' Nome file Curva LIMITE INFERIORE
    NomeFileCurvaLimiteSuperiore_2 = 8      ' Nome file Curva LIMITE SUPERIORE
    
    eNomeColonnaX = 9                       ' Nome da attribuire alla colonna X del file di uscita (nella 1°riga per la 1°colonna)
    eNomeColonnaY = 10                      ' Nome da attribuire alla colonna Y del file di uscita (nella 1°riga per la 2°colonna)
    
    eColonnaY1 = 11                           ' Variabile indipendente di comando
    eNomeColonnaY1 = 12                      ' Nome da attribuire alla colonna Y del file di uscita (nella 1°riga per la 2°colonna)
    eMax
End Enum


' Struttura dati usata per definire e gestire i punti dell'acquisizione
'''Public Type TagPuntiAcquisizione
'''    x As Double
'''    y As Double
'''    Z As Double
'''End Type
Public Type TagPuntiAcquisizione
    X As Double
    Y As Double
    Z As Double
    min1 As Double
    max1 As Double
    min2 As Double
    max2 As Double
End Type


Public Type TagTEST_ACQUISIZIONE_POST
    bOldOn As Boolean                               ' difu del comando di start del modulo
    Return As Integer                               ' Stato del modulo
    State As Integer
    ErrorString As String
    NomeFile As String                              ' Nome del file in Ingresso
    bFirstTime As Boolean                           ' Flag segnalazione primo avvio del programma
    Timer As Long
    Punto() As TagPuntiAcquisizione                 ' Punti della acquisizione
    X() As Double
    Y() As Double
    Z() As Double
    NumeroPuntiAcquisiti As Long                    ' numero di punti acquisiti
    OldTest(eMax) As String                         '(aumenta la dim dell'array perch? ho aggiunto il 2?set di curve limite)
    Time(10) As Long
    CURVE_LIMITE(2) As CurveLimitePlusTag           ' Struttura dati delle curve limite
    IMPOSTAZ_CWGRAPH() As TagImpostazioniGrafico    ' File di configurazione per il Grafico "CWGraph"
    
    Uscita() As TagPuntiAcquisizione                ' Vettore dei punti di uscita
End Type



'--------------------------
' "INFO_ACQUISIZIONE_POST"
'--------------------------
Private Sub INFO_ACQUISIZIONE_POST(Infos As TagInfosSingleTest)

    ' Visualizza GRAFICO
    If (Infos.CWGraph1.Visible <> False) Then
        Infos.CWGraph1.Visible = False
        Infos.CWGraph1.TrackMode = 1
    End If
    
    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
        Infos.CWGraph2.ClearData
        Infos.CWGraph2.TrackMode = 1
    End If
    

    ' Visualizza TEXT BOX
    If (Infos.RichTextBox1.Visible <> False) Then
        Infos.RichTextBox1.Visible = False
    End If

End Sub 'INFO_ACQUISIZIONE_GENERICA()

Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eColonnaX                      = " + Test.Parameter(eColonnaX) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eColonnaY                      = " + Test.Parameter(eColonnaY) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileIngresso               = " + Test.Parameter(NomeFileIngresso) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFileCWGraphSezione        = " + Test.Parameter(eNomeFileCWGraphSezione) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileCurvaLimiteInferiore   = " + Test.Parameter(NomeFileCurvaLimiteInferiore) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileCurvaLimiteSuperiore   = " + Test.Parameter(NomeFileCurvaLimiteSuperiore) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileCurvaLimiteInferiore_2 = " + Test.Parameter(NomeFileCurvaLimiteInferiore_2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "NomeFileCurvaLimiteSuperiore_2 = " + Test.Parameter(NomeFileCurvaLimiteSuperiore_2) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeColonnaX                  = " + Test.Parameter(eNomeColonnaX) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeColonnaY                  = " + Test.Parameter(eNomeColonnaY) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eColonnaY1                     = " + Test.Parameter(eColonnaY1) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeColonnaY1                 = " + Test.Parameter(eNomeColonnaY1) + vbCrLf


End Sub

'-------------------------------------------------------
' Questa funzione imposta gli oggetti grafici "CWGraph1"
'-------------------------------------------------------
Public Function ImpostaCWGraphLoad_GRAPH_POST(ByRef CWGraph1 As CWGraph, ByRef Context As TagTEST_ACQUISIZIONE_POST)

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

End Function 'ImpostaCWGraphLoad_GRAPH_POST()


'-------------------------------------------------------
' Questa funzione plotta i punti nei grafici "CWGraph1"
'-------------------------------------------------------
Public Sub Graph_Points_post(ByRef CWGraph1 As CWGraph, ByRef Context As TagTEST_ACQUISIZIONE_POST, ByRef ulPointIndex As Long, ByVal ClearData As Boolean)
    
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
            CWGraph1.Plots(PLOT_Y1).ChartXvsY Context.Punto(i).X, Context.Punto(i).Y
            CWGraph1.Plots(PLOT_Y2).ChartXvsY Context.Punto(i).X, Context.Punto(i).Z
        Else
            
            TempY2 = (Context.Punto(i - 1).Z)
            
            TempY1 = (Context.Punto(i - 1).Y)
            TempX = Context.Punto(i).X
           
            If (TempY1 > Context.Punto(i).Y) Then
                DeltaError = TempY1 - Context.Punto(i).Y
            Else
                DeltaError = Context.Punto(i).Y - TempY1
            End If
            
            If (DeltaError > 0.5) Then
                TempY1 = Context.Punto(i).Y
                TempX = Context.Punto(i).X
            End If
            
            CWGraph1.Plots(PLOT_Y1).ChartXvsY Context.Punto(i).X, TempY1
        End If
        
        i = i + 1
    Loop
    
    ulPointIndex = UBound(Context.Punto)
    CWGraph1.Plots(PLOT_Y1).Visible = True
    
End Sub 'Graph_Points_post()


'----------------
' GraphSettings()
'----------------
' Funzione di inizializzazione dei settaggi del grafico. Carica le curve limite / Impostazioni di CWGraph
Private Sub GraphSettings(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)

    On Error GoTo err:
    
    Dim bUpdate As Boolean
    Dim i As Integer
    Dim par As CurveLimitePlusParametersTag
    

    ' inizializzazione delle curve limite
    par.Indice(eInferiore) = 1
    par.Indice(eSuperiore) = 2

    CurveLimitePlusInit Context.CURVE_LIMITE(0), par
    CurveLimitePlusInit Context.CURVE_LIMITE(1), par



    ' CURVE LIMITE INF e SUP
    ' ----------------------
    'If (Context.OldTest(NomeFileCurvaLimiteInferiore) <> Test.Parameter(NomeFileCurvaLimiteInferiore) Or _
    '    Context.OldTest(NomeFileCurvaLimiteSuperiore) <> Test.Parameter(NomeFileCurvaLimiteSuperiore)) Then
    
        Infos.CWGraph2.ClearData
   ' End If

    ' curva limite inf
    'If (Context.OldTest(NomeFileCurvaLimiteInferiore) <> Test.Parameter(NomeFileCurvaLimiteInferiore)) Then
    
        
        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(0), Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteInferiore), IndiceSezione, Infos.Label1.Caption), eInferiore) = eERROR_CurveLimite) Then
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CURVE LIMITE INFERIORE", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph2, Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteInferiore), IndiceSezione, Infos.Label1.Caption), eInferiore
        End If


        Context.OldTest(NomeFileCurvaLimiteInferiore) = Test.Parameter(NomeFileCurvaLimiteInferiore)
        bUpdate = True
   ' End If


    ' curva limite inf
    'If (Context.OldTest(NomeFileCurvaLimiteInferiore_2) <> Test.Parameter(NomeFileCurvaLimiteInferiore_2)) Then
    
        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(1), Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteInferiore_2), IndiceSezione, Infos.Label1.Caption), eInferiore) = eERROR_CurveLimite) Then
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CURVE LIMITE INFERIORE 2", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph2, Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteInferiore_2), IndiceSezione, Infos.Label1.Caption), eInferiore_2
        End If


        Context.OldTest(NomeFileCurvaLimiteInferiore_2) = Test.Parameter(NomeFileCurvaLimiteInferiore_2)
        bUpdate = True
    'End If


    ' curva limite sup
    'If (Context.OldTest(NomeFileCurvaLimiteSuperiore) <> Test.Parameter(NomeFileCurvaLimiteSuperiore)) Then

        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(0), Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteSuperiore), IndiceSezione, Infos.Label1.Caption), eSuperiore) = eERROR_CurveLimite) Then
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CURVE LIMITE SUPERIORE 1", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph2, Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteSuperiore), IndiceSezione, Infos.Label1.Caption), eSuperiore
        End If


        Context.OldTest(NomeFileCurvaLimiteSuperiore) = Test.Parameter(NomeFileCurvaLimiteSuperiore)
        bUpdate = True
   ' End If
    
    ' curva limite sup
    'If (Context.OldTest(NomeFileCurvaLimiteSuperiore_2) <> Test.Parameter(NomeFileCurvaLimiteSuperiore_2)) Then

        If (CurveLimitePlusLoad(Context.CURVE_LIMITE(1), Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteSuperiore_2), IndiceSezione, Infos.Label1.Caption), eSuperiore) = eERROR_CurveLimite) Then
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CURVE LIMITE SUPERIORE 2", vbCritical
        Else
            CurveLimitePlusLoad_GRAPH Infos.CWGraph2, Get_Nome_Assoluto_File(Test.Parameter(NomeFileCurvaLimiteSuperiore_2), IndiceSezione, Infos.Label1.Caption), eSuperiore_2
        End If


        Context.OldTest(NomeFileCurvaLimiteSuperiore_2) = Test.Parameter(NomeFileCurvaLimiteSuperiore_2)
        bUpdate = True
    'End If
    

    ' IMPOSTAZIONI "CWGraph1"
    ' ----------------------
    'If (Context.OldTest(eNomeFileCWGraphSezione) <> Test.Parameter(eNomeFileCWGraphSezione)) Then
            
        If (ImpostaCWGraphLoad(Context.IMPOSTAZ_CWGRAPH(), Get_Nome_Assoluto_File(Test.Parameter(eNomeFileCWGraphSezione), IndiceSezione, Infos.Label1.Caption))) Then
            
            ImpostaCWGraphLoad_GRAPH_POST Infos.CWGraph2, Context
            
        Else
            Prepare_MsgBox Infos, "ERRORE CARICAMENTO CONFIGURAZIONE GRAFICO SEZIONE", vbCritical
        End If
        
        Context.OldTest(eNomeFileCWGraphSezione) = Test.Parameter(eNomeFileCWGraphSezione)
        bUpdate = True
        
    'End If
    
    Infos.CWGraph2.PlotAreaColor = vbWhite
    Infos.CWGraph2.Plots(PLOT_Y1).PointColor = vbBlue

    Exit Sub

err:
    Debug.Print err.Description
    
End Sub 'GraphSettings()


Private Function TEST_ACQUISIZIONE_POST_INIT_STRUCT(Context1 As TagTEST_ACQUISIZIONE_POST)
    
    Context1.bFirstTime = False
    Context1.bOldOn = False
    Context1.NomeFile = ""
    Context1.NumeroPuntiAcquisiti = 0

    Context1.Return = 0
    Context1.State = 0
    Context1.Timer = 0
    
    
End Function


'--------------------------
' "TEST_ACQUISIZIONE_POST"
'--------------------------
Public Sub TEST_ACQUISIZIONE_POST(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_ACQUISIZIONE_POST, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "ACQUISIZIONE_POST") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        
        
        If (bStop = True And bOn = True) Then
            TEST_ACQUISIZIONE_POST_INIT_STRUCT Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            Context1.ErrorString = "STOP DA PARTE DELL'UTENTE"
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then        ' fronte di salita del test ( AVVIO )
            'Init Test, Infos, IndiceSezione, Context1               ' init delle variabili del modulo
            INFO_ACQUISIZIONE_POST Infos
            
            Init_Event Test, Infos, IndiceSezione, Context1               ' init delle variabili del modulo
            
            
            
        End If
        
        Task Test, Infos, IndiceSezione, Context1                   ' Task del modulo
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If
        
        bReturn = Context1.Return
    End If
    
End Sub 'TEST_ACQUISIZIONE_POST()


'-----------------
' Task del modulo
'-----------------
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)
    
    Select Case Context.State
        Case eInit: Init Test, Infos, IndiceSezione, Context
        Case eStart:
            Start Test, Infos, IndiceSezione, Context
            Run Test, Infos, IndiceSezione, Context
        'Case eRun:
        
        Case eEnd: Fine Test, Infos, IndiceSezione, Context
        Case eError: ErrorTest Test, Infos, IndiceSezione, Context
        Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select

End Sub 'Task()



Private Sub Init_Event(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)
    Init Test, Infos, IndiceSezione, Context
    GraphSettings Test, Infos, IndiceSezione, Context
End Sub
'------
' Init
'------
' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)
    
    Context.Return = eTEST_BUSY 'eTEST_BUSY_NON_BLOCCANTE
     
    Prepare_Textbox_Parametri Test, Infos

    Context.State = eStart
    

End Sub 'Init()


'-------
' Start
'-------
Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)
    
    On Error GoTo err:
    Context.Return = eTEST_BUSY 'eTEST_BUSY_NON_BLOCCANTE
     
    Context.State = eStart
    
    Context.Timer = GetTickCount + SPEGNINENTO_TIMEOUT_MS
    Context.bFirstTime = True
    
    'GraphSettings Test, Infos, IndiceSezione, Context
    
    Context.NumeroPuntiAcquisiti = 0
    
    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
    
    ReDim Context.X(Context.NumeroPuntiAcquisiti)
    ReDim Context.Y(Context.NumeroPuntiAcquisiti)
    ReDim Context.Z(Context.NumeroPuntiAcquisiti)
    
    Context.NomeFile = Get_Nome_Assoluto_File(Test.Parameter(NomeFileIngresso), IndiceSezione, Infos.Label1.Caption)
    Context.Time(0) = GetTickCountEvo
    
    
    
    
    
    Context.State = eRun
     
    'SPOSTARE TUTTI I RIFERIMENTI AL GRAPH NEL RICHT TEXT TEXT
    Infos.CWGraph2.Plots(PLOT_Y1).ClearData
    Infos.CWGraph2.Plots(PLOT_Y2).ClearData
    
    ' LOG
    '-----
    'AggiungiLogCh IndiceSezione, "| MODULO: ACQUISIZ_POST  | Inizio ---------------------- "
    'Context.Time(1) = GetTickCountEvo
err:
End Sub 'Start()


Private Function Run_Load_Dati(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST) As Boolean
    
    Dim fileId As Long
    Dim LineaDati As String
    Dim fileContent As String
    Dim Linee As Variant
    Dim i As Long
    
    Dim bZetaEnable As Boolean
    
    On Error GoTo err:
    
    Context.bFirstTime = True
    
    fileId = FreeFile
    Open Context.NomeFile For Input As #fileId
                            
'    Do While (EOF(fileId) = False)
'
'        If (Context.bFirstTime = True) Then
'           Line Input #fileId, LineaDati
'           Context.bFirstTime = False
'        End If
'
'        Line Input #fileId, LineaDati
'        Dati = Split(LineaDati, ";")                    ' Prelevo i dati
'
'
'        Context.X(Context.NumeroPuntiAcquisiti) = CDbl(Dati(CInt(Test.Parameter(eColonnaX))))
'        Context.Y(Context.NumeroPuntiAcquisiti) = CDbl(Dati(CInt(Test.Parameter(eColonnaY))))
'
'        Context.Punto(Context.NumeroPuntiAcquisiti).X = CDbl(Dati(CInt(Test.Parameter(eColonnaX))))
'        Context.Punto(Context.NumeroPuntiAcquisiti).Y = CDbl(Dati(CInt(Test.Parameter(eColonnaY))))
'
'        Context.NumeroPuntiAcquisiti = Context.NumeroPuntiAcquisiti + 1
'        ReDim Preserve Context.Punto(Context.NumeroPuntiAcquisiti)
'
'        ReDim Preserve Context.X(Context.NumeroPuntiAcquisiti)
'        ReDim Preserve Context.Y(Context.NumeroPuntiAcquisiti)
'
'    Loop
'
        
        
'    ReDim Preserve Context.Punto(Context.NumeroPuntiAcquisiti)
'
'    ReDim Preserve Context.X(Context.NumeroPuntiAcquisiti)
'    ReDim Preserve Context.Y(Context.NumeroPuntiAcquisiti)
        
'Close #fileId

    fileContent = Input$(LOF(1), #fileId)
    Close #fileId
    Linee = Split(fileContent, vbCr)
    i = 1
    
    Context.NumeroPuntiAcquisiti = UBound(Linee) - 1
    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
    
    ReDim Context.X(Context.NumeroPuntiAcquisiti)
    ReDim Context.Y(Context.NumeroPuntiAcquisiti)
    ReDim Context.Z(Context.NumeroPuntiAcquisiti)
    
    If (UBound(Test.Parameter) >= eColonnaY1) Then
        
        If (IsNumeric(Test.Parameter(eColonnaY1)) = True) Then
            bZetaEnable = True
        End If
    End If
    
    
    
    Do While (i < UBound(Linee))
        Linee(i) = Replace(Linee(i), vbLf, "")
        Dati = Split(Linee(i), ";")                    ' Prelevo i dati
        
        Context.X(i - 1) = CDbl(Dati(CInt(Test.Parameter(eColonnaX))))
        Context.Y(i - 1) = CDbl(Dati(CInt(Test.Parameter(eColonnaY))))
        
        If (bZetaEnable = True) Then
            Context.Z(i - 1) = CDbl(Dati(CInt(Test.Parameter(eColonnaY1))))
            Context.Punto(i - 1).Z = CDbl(Dati(CInt(Test.Parameter(eColonnaY1))))
        End If
                
        
        Context.Punto(i - 1).X = CDbl(Dati(CInt(Test.Parameter(eColonnaX))))
        Context.Punto(i - 1).Y = CDbl(Dati(CInt(Test.Parameter(eColonnaY))))
        
        i = i + 1
    Loop

    
    Run_Load_Dati = False
    Exit Function

err:
    Context1.ErrorString = "ECCEZIONE Run_Load_Dati=" + err.Description
    Context.State = eError
    Debug.Print err.Description
    Close #fileId
    
    Run_Load_Dati = True
End Function 'Run_Load_Dati()



Private Function Run_Test_Dati(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST) As Boolean
    
    On Error GoTo err:

    Dim i As Long
    
    i = 0
    
    If (Context.NumeroPuntiAcquisiti = 0) Then
        Infos.CWGraph2.PlotAreaColor = vbRed
    Else
        'Infos.CWGraph2.Visible = False
        Do While (i < Context.NumeroPuntiAcquisiti)
             
            '--------------------
            ' Check CURVE LIMITE
            '--------------------
            If ((CurveLimitePlusCheck_2(Context.CURVE_LIMITE(0), _
                                          Context.Punto(i).X, _
                                          Context.Punto(i).Y, _
                                          Context.Punto(i).min1, Context.Punto(i).max1) = eERROR_CurveLimite) And _
                (CurveLimitePlusCheck_2(Context.CURVE_LIMITE(1), _
                                          Context.Punto(i).X, _
                                          Context.Punto(i).Y, _
                                          Context.Punto(i).min2, Context.Punto(i).max2) = eERROR_CurveLimite) _
                                          ) Then

                If (Infos.CWGraph2.PlotAreaColor <> vbRed) Then
                    Distributore.Text2 = Context.Punto(i).X
                    Distributore.Text1 = Context.Punto(i).Y
                    Infos.CWGraph2.PlotAreaColor = vbRed
                End If

            End If
        

             '-------------------
             ' PLOTTO su CWGraph1
             '-------------------
'             Infos.CWGraph2.Plots(PLOT_Y1).ChartXvsY Context.Punto(i).X, Context.Punto(i).Y
       
            i = i + 1
        Loop
        'Infos.CWGraph2.Visible = True
        'ddd -----------------------------------------------------
        'ReDim Preserve Context.Punto(Context.NumeroPuntiAcquisiti)
        
        If (UBound(Context.Y) > 0) Then
            ReDim Preserve Context.Y(UBound(Context.Y) - 1)
        End If
        
        If (UBound(Context.Z) > 0) Then
            ReDim Preserve Context.Z(UBound(Context.Z) - 1)
        End If
        
        If (UBound(Context.X) > 0) Then
            ReDim Preserve Context.X(UBound(Context.X) - 1)
        End If
        
        Infos.CWGraph2.ChartLength = UBound(Context.X)
        Infos.CWGraph2.Plots(PLOT_Y1).ChartXvsY Context.X, Context.Y
        
        Infos.CWGraph2.Plots(PLOT_Y2).ChartXvsY Context.X, Context.Z
        
        '---------------------------------------------------------
    
    End If
    Run_Test_Dati = False
    Exit Function

err:
    Run_Test_Dati = True
    Infos.CWGraph2.PlotAreaColor = vbRed

End Function 'Run_Test_Dati()



'------
' Run()
'------

Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)
   On Error GoTo err:
   ' Context.Time(2) = GetTickCountEvo
    
    If (Run_Load_Dati(Test, Infos, IndiceSezione, Context) = True) Then
        Context1.ErrorString = "FUNZIONE Run_Load_Dati in ERRORE"
        Context.State = eError
    
    ElseIf (Run_Test_Dati(Test, Infos, IndiceSezione, Context) = True) Then
        Context1.ErrorString = "FUNZIONE Run_Test_Dati in ERRORE"
        Context.State = eError
        
    Else
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
        Context.State = eEnd
    End If
    
    Infos.CWGraph2.Visible = True
       
err:
End Sub


'------
' Fine
'------
Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)

    If (Infos.CWGraph2.PlotAreaColor = vbRed) Then
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
    
End Sub 'Fine()


'------------
' ErrorTest()
'------------
Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_POST)
    
    If (Infos.RichTextBox1.BackColor <> vbRed) Then
        Infos.RichTextBox1.BackColor = vbRed
    End If
    
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = " ------------------------------------ " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = Context.ErrorString + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = " ------------------------------------ " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = " ! ! ! ! ! ERRORE SUL TEST ! ! ! ! ! " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1 = " ------------------------------------ " + vbLf + Infos.RichTextBox1.Text
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    End If
    
    Context.Return = eTEST_ERROR

End Sub 'ErrorTest()






