Attribute VB_Name = "Module_TEST_POST_PROCESSING"
 
'**************************************************************
'       M O D U L O :    Module_TEST_POST_PROCESSING.bas
'**************************************************************

Private Declare Function GetTickCount Lib "kernel32" () As Long

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
    eNomeFileTipoTest = 1               ' Nome del tipo di test da eseguire
    eMax
End Enum


' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagPuntiAcquisizione
    X As Double
    Y As Double
End Type

' Struttura dati necessaria per gestire la modalità "mediana" del test di post processing
Public Type TagTEST_POST_PROCESSING_MEDIANA
    ' Parametri
    ParametroFiltro As String                    ' Valore della dimensione del filtro software in virgola mobile da applicare
    ' Dati
    MediaMobile() As TagPuntiAcquisizione        ' Vettore della media mobile
    Indice_Punto_Uscita As Long                  ' Indice del punto di uscita
End Type


' Struttura dati usata per il test dei minimi quadrati da applicare alla curva di ingresso
Public Type TagTEST_POST_PROCESSING_MINIMI_QUADRATI
    ' Parametri
   
    ' Dati
    Media_Campionaria_X As Double
    Media_Campionaria_Y As Double
    
    Covarianza As Double
    S_quadro_x As Double
    s_quadro_y As Double
    
    coefficiente_correlazione As Double
    
    coefficiente_A As Double
    coefficiente_B As Double
End Type


' Struttura dati usata per il test dei sottoinsiemi.
' Questa funzione non manipola la tabella dei dati in ingresso, ma
' ne crea un sottoinsieme della stessa applicando alcune regole.
' Le regole che si possono applicare sono :
' prendi i valori asse X maggiore di ...
' prendi i valori asse X minore  di ...
' prendi i valori asse Y maggiore di ...
' prendi i valori asse Y minore  di ...
' se i parametri non sono definiti non si applicaca la logica relativa
Public Type TagTEST_POST_PROCESSING_SOTTOINSIEME
    ' parametri
    Parametro_Maggiore_X As String
    Parametro_Minore_X As String
    Parametro_Maggiore_Y As String
    Parametro_Minore_Y As String
        
    ' Variabili
    Temp() As TagPuntiAcquisizione          ' Vettore della media mobile
    
End Type


Public Type TagTEST_POST_PROCESSING_CHOOSE_COLUMNS
    NoUse As String
End Type


Public Type TagTEST_POST_PROCESSING_UNISCI
    NomeFileDaUnire() As String
End Type

Public Type TagTEST_POST_PROCESSING_SFOLTISCI
    NumeroPuntiAccettatiContinui As String
    NumeroPuntiEliminatiContinui As String
End Type



' Struttura dati usata per generare un grafico in uscita dell'isteresi rispetto
Public Type TagTEST_POST_PROCESSING_ISTERESI
    Parametro_Isteresi_Asse As String
End Type


' Struttura dati usata per generare una curva unica partendo una curva con valor medio
Public Type TagTEST_POST_PROCESSING_HYS_VALOR_MEDIO
    Parametro_Isteresi_Asse As String
    Parametro_Ordine As String
End Type


' Struttura dati necessaria per gestire lo shifting della curva sui due assi di un valore constante
Public Type TagTEST_POST_PROCESSING_SHIFTING
    ParametroX As String                   ' Quantità dello scostamento sull'asse X
    ParametroY As String                   ' Quantità dello scontamento sull'asse Y
End Type



Public Type TagTEST_TOTAL
    Nome_Tipo_Test As String                        ' Nome del tipo di test applicato
    Nome_File_Uscita() As String                      ' Percorso del file di  uscita
    Nome_File_Ingresso As String                    ' Percorso del file in ingresso
    Nome_Colonna_X As String                        ' Nome da associare alla colonna X del file di uscita
    Nome_Colonna_Y As String                        ' Nome da associare alla colonna Y del file di uscita

    Indice_Colonna_X As String                      ' Indice della colonna del campo X del file di ingresso
    Indice_Colonna_Y As String                      ' Indice della colonna del campo Y del file di ingresso
    
    
    FileTest As String
    
    Test_Choose_Columns As TagTEST_POST_PROCESSING_CHOOSE_COLUMNS
    Test_Mediana As TagTEST_POST_PROCESSING_MEDIANA     '
    Test_Shifting As TagTEST_POST_PROCESSING_SHIFTING
    Test_MinimiQuadrati As TagTEST_POST_PROCESSING_MINIMI_QUADRATI
    Test_Sottoinsieme As TagTEST_POST_PROCESSING_SOTTOINSIEME
    Test_Isteresi As TagTEST_POST_PROCESSING_ISTERESI
    Test_Valor_Medio As TagTEST_POST_PROCESSING_HYS_VALOR_MEDIO
    Test_Unisci As TagTEST_POST_PROCESSING_UNISCI
    Test_Sfoltisci As TagTEST_POST_PROCESSING_SFOLTISCI
    
End Type




Public Type TagTEST_POST_PROCESSING
    bOldOn As Boolean                               ' difu del comando di start del modulo
    Return As Integer                               ' Stato del modulo
    State As Integer
    NomeFile As String                              ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean                           ' Flag segnalazione primo avvio del programma
    Timer As Long
    Punto() As TagPuntiAcquisizione                 ' Punti della acquisizione
    NumeroPuntiAcquisiti As Long                    ' numero di punti acquisiti
    OldTest(eMax) As String                         '(aumenta la dim dell'array perchè ho aggiunto il 2°set di curve limite)
    
    StartTimer As Long
    IndiceTest As Long
    
    FileTest As String                              ' Nome del test da applicare al post processing
    'Nome_Tipo_Test As String                        ' Nome del tipo di test applicato
    'Test_Mediana As TagTEST_POST_PROCESSING_MEDIANA     '
    'Test_Shifting As TagTEST_POST_PROCESSING_SHIFTING
    'Test_MinimiQuadrati As TagTEST_POST_PROCESSING_MINIMI_QUADRATI
    'Test_Sottoinsieme As TagTEST_POST_PROCESSING_SOTTOINSIEME
    'Test_ParallelOffset As TagTEST_POST_PROCESSING_PARALLEL_OFFSET
    'Test_Isteresi As TagTEST_POST_PROCESSING_ISTERESI
    'Test_Valor_Medio As TagTEST_POST_PROCESSING_HYS_VALOR_MEDIO
    Test() As TagTEST_TOTAL
    Uscita() As TagPuntiAcquisizione                ' Vettore dei punti di uscita
End Type



'------------------------
' "INFO_POST_PROCESSING"
'------------------------
Private Sub INFO_POST_PROCESSING(Infos As TagInfosSingleTest)

    ' Visualizza GRAFICO
    If (Infos.CWGraph1.Visible <> False) Then
        Infos.CWGraph1.Visible = False
    End If
    
    ' Visualizza TEXT BOX
    If (Infos.RichTextBox1.Visible <> True) Then
        Infos.RichTextBox1.Visible = True
    End If
    
    If (Infos.CWGraph2.Visible <> False) Then
        Infos.CWGraph2.Visible = False
    End If
    
End Sub 'INFO_POST_PROCESSING()



Private Sub TEST_POST_PROCESSING_INIT_STRUCT(Context1 As TagTEST_POST_PROCESSING)
    
    Context1.bFirstTime = False
    Context1.bOldOn = False
    Context1.FileTest = ""
    Context1.NomeFile = ""
    Context1.NumeroPuntiAcquisiti = 0
    
    Context1.Return = 0
    Context1.State = 0
    
End Sub


'------------------------
' "TEST_POST_PROCESSING"
'------------------------
Public Sub TEST_POST_PROCESSING(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_POST_PROCESSING, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "POST_PROCESSING") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        INFO_POST_PROCESSING Infos
        
        If (bStop = True And bOn = True) Then
            TEST_POST_PROCESSING_INIT_STRUCT Context1
        
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then        ' fronte di salita del test ( AVVIO )
            Context1.StartTimer = GetTickCountEvo
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
    
End Sub 'TEST_POST_PROCESSING()


'-----------------
' Task del modulo
'-----------------
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)
    
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
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)
    
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    Dim TestFileName As String
    Dim FileTest As String
    Dim ContenutoFile As String
    
    
    
    On Error GoTo err:
    
    If (Infos.RichTextBox1.BackColor <> vbWhite) Then
        Infos.RichTextBox1.BackColor = vbWhite
    End If
        
    
    Context.Return = eTEST_BUSY_NON_BLOCCANTE
     
    Context.State = eStart
    
    Context.IndiceTest = 0
    
    Context.bFirstTime = True
    ContenutoFile = FileText(Get_Nome_Assoluto_File(Test.Parameter(eNomeFileTipoTest)))
    
    
    ' Inizializzo il numero dei test da svolgere a 0
    ReDim Context.Test(0)
    
    ' Il file è un contenitore di percorsi assoluti di file. Questi file descrivono i test da svolgere
    Righe = Split(ContenutoFile, vbLf)
    
    i = 0
    Do While (i <= UBound(Righe))
    
        Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
        Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
 
        Select Case NOME
        Case "FILE_TEST":
            ReDim Context.Test(UBound(Context.Test)).Nome_File_Uscita(0)
            
            Context.Test(UBound(Context.Test)).Nome_File_Uscita(0) = ""
            Context.Test(UBound(Context.Test)).Nome_File_Ingresso = ""
            Context.Test(UBound(Context.Test)).Nome_Colonna_X = "Colonna X"
            Context.Test(UBound(Context.Test)).Nome_Colonna_Y = "Colonna Y"
            Context.Test(UBound(Context.Test)).Indice_Colonna_X = "0"
            Context.Test(UBound(Context.Test)).Indice_Colonna_Y = "1"
            
            Context.Test(UBound(Context.Test)).FileTest = FileText(Get_Nome_Assoluto_File(Valore))
            Context.Test(UBound(Context.Test)).Nome_Tipo_Test = Init_NomeTest(Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test)))
            
            Init_MediaMobile Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_Choose_Columns Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_Sfoltisci Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_Unisci Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_Shifting Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_MinimiQuadrati Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_Sottoinsieme Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
    
            Init_Isteresi Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            Init_Valor_Medio_Hys Test, Infos, IndiceSezione, Context.Test(UBound(Context.Test))
            
            ReDim Preserve Context.Test(UBound(Context.Test) + 1)
            
        
        End Select

        i = i + 1
    Loop
    
    
    If (Infos.RichTextBox1.Text <> "") Then
        Infos.RichTextBox1.Text = ""
    End If
    
    

    
    
    ' INIT di tutti i sotto test gestiti da questa versione del modulo di post procesing
    'ReDim Context.Uscita(0)
        
    Context.NumeroPuntiAcquisiti = 0
    
    Prepare_Textbox_Parametri Test, Infos
    
    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
    Exit Sub

err:
    Context.State = eError
    Infos.RichTextBox1.Text = " ERROR INIT INDICE " + CStr(UBound(Context.Test)) + vbCrLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = " NOME=" + NOME + vbCrLf + " VALORE=" + Valore + vbCrLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
End Sub 'Init()


'-------
' Start
'-------
Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)
    
    Context.State = eRun
    
    ' LOG
    '-----
'    AggiungiLogCh IndiceSezione, "| MODULO: TEST_POST_PROCESSING | Inizio ---------------------- "

End Sub 'Start()


Private Sub Load_InputFile(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL, ByVal IndiceSezione As Integer)
    
    Dim InputFile As String
    Dim InputRighe As Variant
    Dim bFirstTime As Boolean
    Dim Dati As Variant
    Dim i As Long
    
    On Error GoTo err:
    
    bFirstTime = True
    
    ' Carico il file di ingresso in memoria
    InputFile = FileText(Get_Nome_Assoluto_File(Test.Nome_File_Ingresso))
    InputRighe = Split(InputFile, vbCrLf)
    
    ReDim Context.Punto(0)
    
    i = 0
    Do While (i < UBound(InputRighe))
        
        If (bFirstTime = True) Then
            bFirstTime = False
            i = i + 1
        End If
    
        Dati = Split(InputRighe(i), ";")
        
        If (UBound(Dati) >= CInt(Test.Indice_Colonna_X) And UBound(Dati) >= CInt(Test.Indice_Colonna_X)) Then
            Context.Punto(UBound(Context.Punto)).X = CDbl(Clear_Dato(Dati(Test.Indice_Colonna_X)))
            Context.Punto(UBound(Context.Punto)).Y = CDbl(Clear_Dato(Dati(Test.Indice_Colonna_Y)))
            ReDim Preserve Context.Punto(UBound(Context.Punto) + 1)
        End If
        i = i + 1
    Loop
    
    ReDim Preserve Context.Punto(UBound(Context.Punto) - 1)
    
    Exit Sub

err:
    Debug.Print err.Description

End Sub 'Load_InputFile()



Private Sub Load_InputFile_unisci(FileName As String, Infos As TagInfosSingleTest, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL, ByVal IndiceSezione As Integer)
    
    Dim InputFile As String
    Dim InputRighe As Variant
    Dim bFirstTime As Boolean
    Dim Dati As Variant
    Dim i As Long
    
    On Error GoTo err:
    
    bFirstTime = True
    
    ' Carico il file di ingresso in memoria
    InputFile = FileText(Get_Nome_Assoluto_File(FileName))
    InputRighe = Split(InputFile, vbCrLf)
    
    ReDim Context.Punto(0)
    
    i = 0
    Do While (i < UBound(InputRighe))
        
        If (bFirstTime = True) Then
            bFirstTime = False
            i = i + 1
        End If
    
        Dati = Split(InputRighe(i), ";")
        
        If (UBound(Dati) >= CInt(Test.Indice_Colonna_X) And UBound(Dati) >= CInt(Test.Indice_Colonna_X)) Then
            Context.Punto(UBound(Context.Punto)).X = CDbl(Clear_Dato(Dati(Test.Indice_Colonna_X)))
            Context.Punto(UBound(Context.Punto)).Y = CDbl(Clear_Dato(Dati(Test.Indice_Colonna_Y)))
            ReDim Preserve Context.Punto(UBound(Context.Punto) + 1)
        End If
        i = i + 1
    Loop
    
    ReDim Preserve Context.Punto(UBound(Context.Punto) - 1)
    
    Exit Sub

err:
    Debug.Print err.Description

End Sub 'Load_InputFile()



'------
' Run()
'------
Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)
    Dim i As Long
    Dim j As Long
    Dim bFirstTime As Boolean
    
    ' Carico i punti X-Y in funzione dei parametri di ingresso in memoria
    
    
    
    
    bFirstTime = True
    
    Do While (i < UBound(Context.Test))
    
    
        
                
        ReDim Context.Uscita(0)
        
        Load_InputFile Test, Infos, Context, Context.Test(i), IndiceSezione
        
    
        Run_MediaMobile Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Shifhing Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_MinimiQuadrati Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Sottoinsieme Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Isteresi Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Choose_Columns Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Sfoltisci Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Unisci Test, Infos, IndiceSezione, Context, Context.Test(i)
        Run_Valor_Medio_Hys Test, Infos, IndiceSezione, Context, Context.Test(i)
    
        
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "FILE IN:" + Context.Test(i).Nome_File_Ingresso + vbLf
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "TEST:" + Context.Test(i).Nome_Tipo_Test + " PUNTI : " + CStr(UBound(Context.Uscita)) + vbLf
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
        
        SalvaPunti Test, Infos, Context, IndiceSezione, Context.Test(i)
                
                
        If (Context.State = eError) Then
            i = UBound(Context.Test)
        Else
            i = i + 1
        End If
    Loop
    
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " TEST TIME = " + CStr(GetTickCountEvo - Context.StartTimer) + " (ms)" + vbCrLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
End Sub 'Run()


'-----------
' SaveFile()
'-----------
Private Sub SaveFile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)

    ' LOG
    '-----
    'AggiungiLogCh IndiceSezione, "| MODULO: TEST_POST_PROCESSING | SALVATAGGIO DATI POST PROCESSING"
                
    'SalvaPunti Test, Infos, Context, IndiceSezione
    
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    
    Context.State = eEnd
    
End Sub 'SaveFile()


'------
' Fine
'------
Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)

    If (Infos.CWGraph1.PlotAreaColor = vbRed) Then
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
    
End Sub 'Fine()


'------------
' ErrorTest()
'------------
Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING)
    
    Infos.RichTextBox1.BackColor = vbRed
    
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = " ! ! ! ! ! ERRORE SUL TEST ! ! ! ! ! " + vbLf + Infos.RichTextBox1.Text
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    End If
    
    Context.Return = eTEST_ERROR

End Sub 'ErrorTest()


'-------------
' SalvaPunti()
'-------------
Private Sub SalvaPunti(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_POST_PROCESSING, IndiceSezione As Integer, Test As TagTEST_TOTAL)
    
    
    Dim i As Long
    Dim Index As Long
    
    Dim ora As String
    Dim oggi As String
        
        
On Error GoTo err:

    ora = Format(Time, ("HH.MM.SS"))
    ora = Replace(ora, ".", "")

    oggi = Format(Date, ("yy/mm/dd"))
    oggi = Replace(oggi, "/", "")
    
    
    Index = 0
    Do While (Index < UBound(Test.Nome_File_Uscita))
    
        Open Test.Nome_File_Uscita(Index) For Output As #22     ' per file POST di lavoro
        '--------------------------------------------------
        
        ' CREA 1° RIGA FILE
        '--------------------------------------------------
            
        Print #22, Test.Nome_Colonna_X + ";" + Test.Nome_Colonna_Y   ' per file ALL
    
        '--------------------------------------------------
        
        ' SCRIVI DATI IN FILE
        '--------------------------------------------------
        Do While (i <= UBound(Context.Uscita))
                    
            Print #22, CStr(Context.Uscita(i).X) + ";" + CStr(Context.Uscita(i).Y)
    
            i = i + 1
        Loop
        '--------------------------------------------------
        
        ' CHIUDI FILE
        '----------------------------
            
        Close #22       ' per file ALL
        
        Index = Index + 1
    Loop
    '----------------------------
    

    
Exit Sub

err:
    Prepare_MsgBox Infos, "ERRORE DURANTE IL SALVATAGGIO DELLA CURVA " + err.Description, vbCritical
    
End Sub 'SalvaPunti()


Private Function FileText(FileName As String) As String
    Dim handle As Integer
    handle = FreeFile
    
    Open FileName For Input As #handle
    
    FileText = Input$(LOF(handle), handle)
    
    Close #handle
    
End Function


Private Function Init_NomeTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL) As String
    
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    Righe = Split(Context.FileTest, vbLf)
    
        
    
    i = 0
    Do While (i <= UBound(Righe))
        Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
        Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
 

        Select Case NOME
        Case "NOME_TEST": Init_NomeTest = Valore
        End Select

        i = i + 1
    Loop
End Function


' Funzione di inizializzazione
Private Sub Init_MediaMobile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "MEDIANA") Then
    
        Righe = Split(Context.FileTest, vbLf)
        
    
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
     
    
            Select Case NOME
            Case "FILTRO": Context.Test_Mediana.ParametroFiltro = Valore
            Case "NOME_FILE_USCITA":
            
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)
                
            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
        Loop
    
        ReDim Context.Test_Mediana.MediaMobile(CInt(Context.Test_Mediana.ParametroFiltro))
        Context.Test_Mediana.Indice_Punto_Uscita = 0
        
    
    End If
End Sub


Private Sub Init_Choose_Columns(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "CHOOSE_COLUMNS") Then
    
        Righe = Split(Context.FileTest, vbLf)
        
    
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
                 
            Select Case NOME
            Case "NOME_FILE_USCITA":
                
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
        Loop
                    
    
    End If
End Sub

Private Sub Init_Sfoltisci(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "SFOLTISCI") Then
    
        Righe = Split(Context.FileTest, vbLf)
        
    
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
                 
            Select Case NOME
            Case "NUMERO_PUNTI_OK"
                Context.Test_Sfoltisci.NumeroPuntiAccettatiContinui = Valore
                
            Case "NUMERO_PUNTI_NOK"
                Context.Test_Sfoltisci.NumeroPuntiEliminatiContinui = Valore
            Case "NOME_FILE_USCITA":
                
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
        Loop
                    
    
    End If
End Sub



Private Sub Init_Unisci(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "UNISCI") Then
    
        Righe = Split(Context.FileTest, vbLf)
        Context.Nome_File_Ingresso = ""
        ReDim Context.Test_Unisci.NomeFileDaUnire(0)
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
                 
            Select Case NOME
            Case "NOME_FILE_USCITA":
                
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            Case "NOME_FILE_INGRESSO":
                If (Context.Nome_File_Ingresso = "") Then
                    Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                Else
                    Context.Test_Unisci.NomeFileDaUnire(UBound(Context.Test_Unisci.NomeFileDaUnire)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                    ReDim Preserve Context.Test_Unisci.NomeFileDaUnire(UBound(Context.Test_Unisci.NomeFileDaUnire) + 1)
                End If
                
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
        Loop
                    
    
    End If
End Sub




Private Sub Run_MediaMobile(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)
    On Error GoTo err:
    Dim Index_Vettore_media_Mobile As Long
    Dim i As Long
    Dim CalcY As Double
    Dim CalcX As Double
    Dim j As Long
    
    Dim Primo_Punto_Valido As Boolean
    
    If (Test.Nome_Tipo_Test <> "MEDIANA") Then
        Exit Sub
    End If
    
    
    
    i = 0
    Do While (i < UBound(Context.Punto))
                        
        Test.Test_Mediana.MediaMobile(Index_Vettore_media_Mobile) = Context.Punto(i)
            
        Index_Vettore_media_Mobile = Index_Vettore_media_Mobile + 1
        
        If (Index_Vettore_media_Mobile > UBound(Test.Test_Mediana.MediaMobile)) Then
            Index_Vettore_media_Mobile = 0
            Primo_Punto_Valido = True
        End If
        
    
        If (Primo_Punto_Valido = True) Then
        
            j = 0
            CalcY = 0
            CalcX = 0
            Do While (j <= UBound(Test.Test_Mediana.MediaMobile))
                CalcY = CalcY + Test.Test_Mediana.MediaMobile(j).Y
                CalcX = CalcX + Test.Test_Mediana.MediaMobile(j).X
                j = j + 1
            Loop
            
            Context.Uscita(UBound(Context.Uscita)).Y = CalcY / j
            Context.Uscita(UBound(Context.Uscita)).X = CalcX / j
            
            ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
            
        End If
        
        i = i + 1
        
    Loop


    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If


    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
    Exit Sub
err:
    Context.State = eError
End Sub


Private Sub Init_Shifting(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "SHIFTING") Then
    
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " SHIFTING " + vbLf
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
        
        Righe = Split(Context.FileTest, vbLf)
        
        Context.Test_Shifting.ParametroX = "0"
        Context.Test_Shifting.ParametroY = "0"
        
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
     
    
            Select Case NOME
            Case "ASSE_X": Context.Test_Shifting.ParametroX = Valore
            Case "ASSE_Y": Context.Test_Shifting.ParametroY = Valore
            Case "NOME_FILE_USCITA":
                
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
            
        Loop
    
        
    
    End If
    
End Sub




Private Sub Run_Choose_Columns(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)

    Dim i As Long

    
    If (Test.Nome_Tipo_Test <> "CHOOSE_COLUMNS") Then
        Exit Sub
    End If
                 
    i = 0
    Do While (i <= UBound(Context.Punto))
                        
                        
        Context.Uscita(UBound(Context.Uscita)).X = Context.Punto(i).X
        Context.Uscita(UBound(Context.Uscita)).Y = Context.Punto(i).Y
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
        i = i + 1
        
    Loop
    
        
       
    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If
    
    

End Sub



Private Sub Run_Sfoltisci(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)

    Dim i As Long
    Dim StatoMacchina As String
    Dim IndexNumeroOk As Long
    Dim IndexNumeroNok As Long
    Dim Index As Long
    
    If (Test.Nome_Tipo_Test <> "SFOLTISCI") Then
        Exit Sub
    End If
                 
             
    i = 0
    StatoMacchina = "ACCETTA"
    Do While (i <= UBound(Context.Punto))
        
        Select Case StatoMacchina
        Case "ACCETTA"
        
            If (IndexNumeroOk < Test.Test_Sfoltisci.NumeroPuntiAccettatiContinui) Then
                Context.Uscita(UBound(Context.Uscita)).X = Context.Punto(i).X
                Context.Uscita(UBound(Context.Uscita)).Y = Context.Punto(i).Y
                ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
            Else
                StatoMacchina = "NON ACCETTA"
            End If
            IndexNumeroNok = 0
            IndexNumeroOk = IndexNumeroOk + 1
            
        Case "NON ACCETTA"
        
            If (IndexNumeroNok < Test.Test_Sfoltisci.NumeroPuntiEliminatiContinui) Then
                ' scarto il punto
            Else
                StatoMacchina = "ACCETTA"
            End If
        
            IndexNumeroNok = IndexNumeroNok + 1
            IndexNumeroOk = 0
            
            
        End Select
        
        i = i + 1
    Loop
    
       
    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If
    

End Sub




Private Sub Run_Unisci(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)

    Dim i As Long
    Dim Index As Long
    
    If (Test.Nome_Tipo_Test <> "UNISCI") Then
        Exit Sub
    End If
                 
             
    i = 0
    Do While (i <= UBound(Context.Punto))
                        
        
        Context.Uscita(UBound(Context.Uscita)).X = Context.Punto(i).X
        Context.Uscita(UBound(Context.Uscita)).Y = Context.Punto(i).Y
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
        i = i + 1
        
    Loop
    
    Do While (Index < UBound(Test.Test_Unisci.NomeFileDaUnire))
        
        Infos.RichTextBox1.Text = "UNISCO " + Test.Test_Unisci.NomeFileDaUnire(Index) + vbCrLf
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
        
        Load_InputFile_unisci Test.Test_Unisci.NomeFileDaUnire(Index), Infos, Context, Test, IndiceSezione
        
        
        i = 0
        Do While (i <= UBound(Context.Punto))
                            
                            
            Context.Uscita(UBound(Context.Uscita)).X = Context.Punto(i).X
            Context.Uscita(UBound(Context.Uscita)).Y = Context.Punto(i).Y
            ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
            i = i + 1
            
        Loop
        
        Index = Index + 1
    Loop
    
    
    
    
       
    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If
    

End Sub






Private Sub Run_Shifhing(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)
    On Error GoTo err:
    Dim Index_Vettore_media_Mobile As Long
    Dim i As Long
    Dim CalcY As Double
    Dim CalcX As Double
    Dim j As Long
    Dim Shifting_X As Double
    Dim Shifting_Y As Double
    
    Dim Primo_Punto_Valido As Boolean
    
    If (Test.Nome_Tipo_Test <> "SHIFTING") Then
        Exit Sub
    End If
    
    If (Test.Test_Shifting.ParametroX = "MIN") Then
        i = 0
        
        Do While (i <= UBound(Context.Punto))
            If (i = 0) Then
            
                Shifting_X = Context.Punto(i).X
                
            ElseIf (Context.Punto(i).X < Shifting_X) Then
            
                Shifting_X = Context.Punto(i).X
                
            End If
            
            i = i + 1
        Loop
        
        Shifting_X = -Shifting_X
    Else
        Shifting_X = CDbl(Test.Test_Shifting.ParametroX)
    End If
    
    If (Test.Test_Shifting.ParametroY = "MIN") Then
        i = 0
        
        Do While (i <= UBound(Context.Punto))
            If (i = 0) Then
            
                Shifting_Y = Context.Punto(i).Y
                
            ElseIf (Context.Punto(i).X < Shifting_Y) Then
            
                Shifting_Y = Context.Punto(i).Y
                
            End If
            
            i = i + 1
        Loop
        
        Shifting_Y = -Shifting_Y
    Else
        Shifting_Y = CDbl(Test.Test_Shifting.ParametroY)
    End If
    
        
    
    i = 0
    Do While (i <= UBound(Context.Punto))
                        
                        
        Context.Uscita(UBound(Context.Uscita)).X = Context.Punto(i).X + Shifting_X
        Context.Uscita(UBound(Context.Uscita)).Y = Context.Punto(i).Y + Shifting_Y
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
        i = i + 1
        
    Loop
    
        
       
    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If
    
    
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
    Exit Sub
err:

    Context.State = eError
    
    
End Sub


Private Sub Init_MinimiQuadrati(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    
    Dim i As Long
    Dim NOME As String
    Dim Valore As String
    Dim Riche As Variant
    
    If (Context.Nome_Tipo_Test <> "MINIMI_QUADRATI") Then
        Exit Sub
    End If
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + vbLf + "MINIMI_QUADRATI" + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    
    ' Parametri
    Context.Test_MinimiQuadrati.Media_Campionaria_X = 0
    Context.Test_MinimiQuadrati.Media_Campionaria_Y = 0
    
    ' Dati
    Context.Test_MinimiQuadrati.Covarianza = 0
    Context.Test_MinimiQuadrati.S_quadro_x = 0
    Context.Test_MinimiQuadrati.s_quadro_y = 0

    Context.Test_MinimiQuadrati.coefficiente_correlazione = 0
    
        Righe = Split(Context.FileTest, vbLf)
        
    
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
         
            Select Case NOME
            Case "NOME_FILE_USCITA":
            
                
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
        Loop
    

End Sub


Private Sub Run_MinimiQuadrati(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)
    
    Dim Calc1 As Double
    Dim Calc2 As Double
    Dim Calc3 As Double
    Dim i As Long
    
    On Error GoTo err
    
    If (Test.Nome_Tipo_Test <> "MINIMI_QUADRATI") Then
        Exit Sub
    End If
    
    
    ' Calcolo medie campionarie
    i = 0
    Calc1 = 0
    Calc2 = 0
    Do While (i < UBound(Context.Punto))
                            
        Calc1 = Calc1 + Context.Punto(i).X
        Calc2 = Calc2 + Context.Punto(i).Y
            
        i = i + 1
        
    Loop
    
    If (UBound(Context.Punto) <= 0) Then
        Context.State = eError
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " NESSUN PUNTO DA CALCOLARE"
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
        Exit Sub
    End If
    Test.Test_MinimiQuadrati.Media_Campionaria_X = Calc1 / i
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " Media Campionaria X = " + CStr(Test.Test_MinimiQuadrati.Media_Campionaria_X) + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    
    Test.Test_MinimiQuadrati.Media_Campionaria_Y = Calc2 / i
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " Media Campionaria Y = " + CStr(Test.Test_MinimiQuadrati.Media_Campionaria_Y) + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    
    ' Calcolo della covarianza
    ' Calcolo di s_quadro_x
    ' Calcolo di s_quadro_y
    i = 0
    Calc1 = 0
    Calc2 = 0
    Calc3 = 0
    
    Do While (i < UBound(Context.Punto))
                                    
        ' Calcolo della covarianza
        Calc1 = (Context.Punto(i).X - Test.Test_MinimiQuadrati.Media_Campionaria_X) * (Context.Punto(i).Y - Test.Test_MinimiQuadrati.Media_Campionaria_Y)
        
        ' Calcolo di s_quadro_x
        Calc2 = (Context.Punto(i).X - Test.Test_MinimiQuadrati.Media_Campionaria_X) * (Context.Punto(i).X - Test.Test_MinimiQuadrati.Media_Campionaria_X)
        
        ' Calcolo di s_quadro_y
        Calc3 = (Context.Punto(i).Y - Test.Test_MinimiQuadrati.Media_Campionaria_Y) * (Context.Punto(i).Y - Test.Test_MinimiQuadrati.Media_Campionaria_Y)
        
        i = i + 1
        
    Loop
        
    Test.Test_MinimiQuadrati.Covarianza = 1 / (i - 1) * Calc1
    Test.Test_MinimiQuadrati.S_quadro_x = 1 / (i - 1) * Calc2
    Test.Test_MinimiQuadrati.s_quadro_y = 1 / (i - 1) * Calc3
    
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " Covarianza = " + CStr(Test.Test_MinimiQuadrati.Covarianza) + vbLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " S_quadro_x = " + CStr(Test.Test_MinimiQuadrati.S_quadro_x) + vbLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + " s_quadro_y = " + CStr(Test.Test_MinimiQuadrati.s_quadro_y) + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    
    ' coefficiente di correlazione
    Test.Test_MinimiQuadrati.coefficiente_correlazione = (Test.Test_MinimiQuadrati.Covarianza) / (Sqr(Test.Test_MinimiQuadrati.S_quadro_x) * Sqr(Test.Test_MinimiQuadrati.s_quadro_y))
    
    Infos.RichTextBox1 = Infos.RichTextBox1.Text + " coefficiente di correlazione = " + CStr(Test.Test_MinimiQuadrati.s_quadro_y) + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    
    ' Calcolo i coefficienti della retta
    Test.Test_MinimiQuadrati.coefficiente_A = Test.Test_MinimiQuadrati.Covarianza / Test.Test_MinimiQuadrati.S_quadro_x
    Test.Test_MinimiQuadrati.coefficiente_B = Test.Test_MinimiQuadrati.Media_Campionaria_Y - Test.Test_MinimiQuadrati.coefficiente_A * Test.Test_MinimiQuadrati.Media_Campionaria_X
    
    Infos.RichTextBox1 = Infos.RichTextBox1.Text + " Y = Ax + B" + vbLf
    Infos.RichTextBox1 = Infos.RichTextBox1.Text + " Dove : " + vbLf
    Infos.RichTextBox1 = Infos.RichTextBox1.Text + " A = " + CStr(Test.Test_MinimiQuadrati.coefficiente_A) + vbLf
    Infos.RichTextBox1 = Infos.RichTextBox1.Text + " B = " + CStr(Test.Test_MinimiQuadrati.coefficiente_B) + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    
    
    ' Creo il file di uscita
    i = 0
    Do While (i < UBound(Context.Punto))
    
        Context.Uscita(UBound(Context.Uscita)).X = Context.Punto(i).X
        Context.Uscita(UBound(Context.Uscita)).Y = Context.Punto(i).X * Test.Test_MinimiQuadrati.coefficiente_A + Test.Test_MinimiQuadrati.coefficiente_B
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) + 1)
    
        i = i + 1
    Loop

    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
        
    Exit Sub
        
err:
    Infos.RichTextBox1 = err.Description + vbLf
    Context.State = eError
    Exit Sub
        
'err:
'    Debug.Print err.Description
        
End Sub


Private Sub Init_Valor_Medio_Hys(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "VALORE_MEDIO_HYS") Then
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "VALOR MEDIO HYS " + vbLf + vbLf
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
        
        Righe = Split(Context.FileTest, vbLf)
        
        i = 0
        
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
     
    
            Select Case NOME
            Case "ORDINA": Context.Test_Valor_Medio.Parametro_Ordine = Valore
            Case "NOME_FILE_USCITA":
            
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
            
        Loop
        
    End If
End Sub


Private Sub Run_Valor_Medio_Hys(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)


    On Error GoTo err:
    Dim i As Long
    Dim TipoTest As String

    
    
    If (Test.Nome_Tipo_Test <> "VALORE_MEDIO_HYS") Then
        Exit Sub
    End If
        
    ' sottoinsieme asse X
    i = 0
    Do While (i <= UBound(Context.Punto))
        Run_Valor_Medio_Hys_Calc Context.Punto(i).X, Context.Punto, Context.Uscita
        i = i + 1
    Loop
    ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    
    If (Test.Test_Valor_Medio.Parametro_Ordine = "CRESCENTE") Then
        BubbleSort1DArray Context.Uscita, True, Context.Uscita
    End If
    
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "PUNTI : " + CStr(UBound(Context.Uscita)) + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
    

    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
    Exit Sub
err:
    Context.State = eError
    Debug.Print err.Description
End Sub



Function BubbleSort1DArray(ByRef Vin() As TagPuntiAcquisizione, bAscending As Boolean, ByRef vRet() As TagPuntiAcquisizione) As Boolean
    ' Sorts the single dimension list array, ascending or descending
    ' Returns sorted list in vRet if supplied, otherwise in vIn modified
On Error GoTo err:
    Dim First As Long, Last As Long
    Dim i As Long, j As Long, bWasMissing As Boolean
    Dim Temp As TagPuntiAcquisizione, vW() As TagPuntiAcquisizione
    
    First = LBound(Vin)
    Last = UBound(Vin)
    
    ReDim vW(First To Last)
    
    i = First
    Do While (i <= Last)
        vW(i).X = Vin(i).X
        vW(i).Y = Vin(i).Y
        i = i + 1
    Loop
    
    
    If bAscending = True Then
        For i = First To Last - 1
            For j = i + 1 To Last
                If vW(i).X > vW(j).X Then
                
                Temp.X = vW(j).X
                Temp.Y = vW(j).Y
                
                vW(j).X = vW(i).X
                vW(j).Y = vW(i).Y
                
                vW(i).X = Temp.X
                vW(i).Y = Temp.Y
                
                End If
            Next j
        Next i
    Else 'descending sort
        For i = First To Last - 1
            For j = i + 1 To Last
                If vW(i).X < vW(j).X Then
                Temp.X = vW(j).X
                Temp.Y = vW(j).Y
                
                vW(j).X = vW(i).X
                vW(j).Y = vW(i).Y
                
                
                vW(i).X = Temp.X
                vW(i).Y = Temp.Y
                
                End If
            Next j
        Next i
    End If
  

   'transfers
     ReDim vRet(First To Last)
     
    i = First
    Do While (i <= Last)
        vRet(i).X = vW(i).X
        vRet(i).Y = vW(i).Y
        i = i + 1
    Loop
    
     
   
   BubbleSort1DArray = True
    Exit Function
    
err:
    Debug.Print err.Description
End Function


Private Sub Init_Isteresi(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)


    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "ISTERESI") Then
        Infos.RichTextBox1.Text = " ISTERESI " + vbLf
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)

        Righe = Split(Context.FileTest, vbLf)

        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
     
    
            Select Case NOME
            Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
            Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
            Case "NOME_FILE_USCITA":
            
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)

            
            Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
            Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
            Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
            
            End Select
    
            i = i + 1
        Loop


    End If
End Sub


Private Function Run_Isteresi_create_curve(X As Double, Punto1 As TagPuntiAcquisizione, Punto2 As TagPuntiAcquisizione) As Double
    Dim k As Double
    Dim c As Double
    Dim Y As Double
    On Error GoTo err:
    
    If ((Punto2.X - Punto1.X) <> 0) Then
    
        k = (Punto2.Y - Punto1.Y) / (Punto2.X - Punto1.X)
        c = Punto2.Y - k * Punto2.X
        If (k <> 0) Then
            Debug.Print CStr(k)
        End If
        
        Y = k * X + c
        
        Run_Isteresi_create_curve = Y
    Else
        Run_Isteresi_create_curve = Punto1.Y
    End If
        
    Exit Function
err:
    Debug.Print err.Description
End Function

 
Private Function Run_Valor_Medio_Hys_Calc(X As Double, Punti() As TagPuntiAcquisizione, Uscita() As TagPuntiAcquisizione)
    Dim i As Long
    Dim Y() As Double
    Dim bFirst As Boolean
    Dim bUscitaPresente As Boolean
    Dim min As Double
    Dim max As Double
    Dim Isteresi As Double
    Dim Index() As Long
    On Error GoTo err:
    ' escludo l'ultimo punto
    i = 0
    bFirst = True
    Do While (i < UBound(Punti) - 1)
        If (Punti(i).X < Punti(i + 1).X) Then
            min = Punti(i).X
            max = Punti(i + 1).X
        Else
            max = Punti(i).X
            min = Punti(i + 1).X
        End If
        
        If (X >= min And X <= max) Then
            If (bFirst = True) Then
                bFirst = False
                ReDim Y(0)
                ReDim Index(0)
            Else
                ReDim Preserve Y(UBound(Y) + 1)
                ReDim Preserve Index(UBound(Index) + 1)
                
            End If
            
            Y(UBound(Y)) = Run_Isteresi_create_curve(X, Punti(i), Punti(i + 1))
            Index(UBound(Index)) = i
        End If
        
        i = i + 1
    Loop
    
    
    ' controllo se il punto X è già presente nella tabella di "Uscita"
    ' in tal caso non deve essere utilizzato
    i = 0
    bUscitaPresente = False
    Do While (i < UBound(Uscita) - 1 And UBound(Uscita) >= 2)
        If (Uscita(i).X < Uscita(i + 1).X) Then
            min = Uscita(i).X
            max = Uscita(i + 1).X
        Else
            max = Uscita(i).X
            min = Uscita(i + 1).X
        End If
    
        If (X >= min And X <= max) Then
            ' Non utilizzo il punto
            bUscitaPresente = True
            i = UBound(Uscita)
        End If
        
        i = i + 1
    Loop
    
    If (bUscitaPresente = False) Then
    
        ' ora controllo quanti punti ho trovato nella tabella
        If (bFirst = False And UBound(Y) >= 1) Then
            
            i = 0
            min = Y(0)
            max = Y(0)
            
            Do While (i <= UBound(Y))
                If (min > Y(i)) Then
                    min = Y(i)
                End If
                
                If (max < Y(i)) Then
                    max = Y(i)
                End If
                
                i = i + 1
            Loop
            
                
            Isteresi = max - min
            Uscita(UBound(Uscita)).X = X
            Uscita(UBound(Uscita)).Y = min + Isteresi / 2
            
            ReDim Preserve Uscita(UBound(Uscita) + 1)
        
        ElseIf (bFirst = False) Then
        
            Uscita(UBound(Uscita)).X = X
            Uscita(UBound(Uscita)).Y = Y(0)
            
            ReDim Preserve Uscita(UBound(Uscita) + 1)
                        
        End If
    
    End If
    
    Exit Function
err:
Debug.Print err.Description
End Function


Private Function Run_Isteresi_Calc(X As Double, Punti() As TagPuntiAcquisizione, Uscita() As TagPuntiAcquisizione)
    Dim i As Long
    Dim Y() As Double
    Dim linea() As Double
    Dim bFirst As Boolean
    Dim bUscitaPresente As Boolean
    
    Dim min As Double
    Dim max As Double
    Dim Isteresi As Double
    On Error GoTo err:
    ' escludo l'ultimo punto
    i = 0
    bFirst = True
    Do While (i <= UBound(Punti) - 1)
        If (Punti(i).X < Punti(i + 1).X) Then
            min = Punti(i).X
            max = Punti(i + 1).X
        Else
            max = Punti(i).X
            min = Punti(i + 1).X
        End If
        
        If (X >= min And X <= max) Then
            If (bFirst = True) Then
                bFirst = False
                ReDim Y(0)
                ReDim linea(0)
            Else
                ReDim Preserve Y(UBound(Y) + 1)
                ReDim Preserve linea(UBound(linea) + 1)
            End If
            
            Y(UBound(Y)) = Run_Isteresi_create_curve(X, Punti(i), Punti(i + 1))
           linea(UBound(linea)) = i
        End If
        
        i = i + 1
    Loop
    
    ' controllo se il punto X è già presente nella tabella di "Uscita"
    ' in tal caso non deve essere utilizzato
    i = 0
    bUscitaPresente = False
    Do While (i < UBound(Uscita) - 1 And UBound(Uscita) >= 2)
        If (Uscita(i).X < Uscita(i + 1).X) Then
            min = Uscita(i).X
            max = Uscita(i + 1).X
        Else
            max = Uscita(i).X
            min = Uscita(i + 1).X
        End If
    
        If (X >= min And X <= max) Then
            ' Non utilizzo il punto
            bUscitaPresente = True
            i = UBound(Uscita)
        End If
        
        i = i + 1
    Loop
    
    ' ora controllo quanti punti ho trovato nella tabella
    If (bFirst = False And bUscitaPresente = False And UBound(Y) >= 1) Then
        
        
        i = 0
        min = Y(0)
        max = Y(0)
        
        Do While (i <= UBound(Y))
            If (min > Y(i)) Then
                min = Y(i)
            End If
            
            If (max < Y(i)) Then
                max = Y(i)
            End If
            
            i = i + 1
        Loop
        
            
        Isteresi = max - min
        Uscita(UBound(Uscita)).X = X
        Uscita(UBound(Uscita)).Y = Isteresi
        
        ReDim Preserve Uscita(UBound(Uscita) + 1)
    
    End If
    Exit Function
err:
Debug.Print err.Description

End Function 'Run_Isteresi_Calc()


Private Sub Run_Isteresi(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)


    On Error GoTo err:
    Dim i As Long
    Dim TipoTest As String
    
    
    If (Test.Nome_Tipo_Test <> "ISTERESI") Then
        Exit Sub
    End If
            
        
        
    ' sottoinsieme asse X
    i = 0
    Do While (i <= UBound(Context.Punto))
        Run_Isteresi_Calc Context.Punto(i).X, Context.Punto, Context.Uscita
        i = i + 1
    Loop
    
    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If

    
err:
    Debug.Print err.Description
End Sub









Private Sub Init_Sottoinsieme(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_TOTAL)


    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    
    If (Context.Nome_Tipo_Test = "SOTTOINSIEME") Then
        
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "SOTTOINSIEME" + vbLf
        Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
        
        Context.Test_Sottoinsieme.Parametro_Maggiore_X = ""
        Context.Test_Sottoinsieme.Parametro_Maggiore_Y = ""
        Context.Test_Sottoinsieme.Parametro_Minore_X = ""
        Context.Test_Sottoinsieme.Parametro_Minore_Y = ""
                        
            
        Righe = Split(Context.FileTest, vbLf)
            
        i = 0
        Do While (i <= UBound(Righe))
            Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
            Get_Coppia_Variabile_Valore Righe(i), NOME, Valore
     
    
            Select Case NOME
                Case "MAGGIORE_X": Context.Test_Sottoinsieme.Parametro_Maggiore_X = Valore
                Case "MAGGIORE_Y": Context.Test_Sottoinsieme.Parametro_Maggiore_Y = Valore
                Case "MINORE_Y": Context.Test_Sottoinsieme.Parametro_Minore_Y = Valore
                Case "MINORE_X": Context.Test_Sottoinsieme.Parametro_Minore_X = Valore
                Case "NOME_COLONNA_X": Context.Nome_Colonna_X = Valore
                Case "NOME_COLONNA_Y": Context.Nome_Colonna_Y = Valore
                Case "NOME_FILE_USCITA":
                    
                Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita)) = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                ReDim Preserve Context.Nome_File_Uscita(UBound(Context.Nome_File_Uscita) + 1)
                    
                    
                Case "NOME_FILE_INGRESSO": Context.Nome_File_Ingresso = Get_Nome_Assoluto_File(Valore, IndiceSezione, Infos.Label1.Caption)
                Case "INDICE_COLONNA_Y": Context.Indice_Colonna_Y = Valore
                Case "INDICE_COLONNA_X": Context.Indice_Colonna_X = Valore
                
            End Select
 
            i = i + 1
        Loop

    End If
    
End Sub 'Init_Sottoinsieme()





Private Function Run_Sottoinsieme_Tipo_Test(Maggiore As String, Minore As String) As String
    
    Dim tipo_Test As String

    ' Nessun parametro di x definito
    If (Maggiore = "" And Minore = "") Then
        tipo_Test = "TUTTI"
    ElseIf (Maggiore = "") Then
        tipo_Test = "MINORE"
    ElseIf (Minore = "") Then
        tipo_Test = "MAGGIORE"
    ElseIf (CDbl(Maggiore) > CDbl(Minore)) Then
        tipo_Test = "ESTERNI"
    Else
        tipo_Test = "INTERNI"
    End If
    
    Run_Sottoinsieme_Tipo_Test = tipo_Test
    
End Function




Private Function Run_Sottoinsieme_Esegui_Test(ByVal Valore As Double, Minore As String, Maggiore As String, TipoTest As String, Punto As TagPuntiAcquisizione, Uscita() As TagPuntiAcquisizione)
        
    
    Select Case TipoTest
    
    Case "TUTTI"
    
        Uscita(UBound(Uscita)).X = Punto.X
        Uscita(UBound(Uscita)).Y = Punto.Y
        ReDim Preserve Uscita(UBound(Uscita) + 1)
        
    Case "MINORE"
        
        If (Valore < CDbl(Minore)) Then
            Uscita(UBound(Uscita)).X = Punto.X
            Uscita(UBound(Uscita)).Y = Punto.Y
            ReDim Preserve Uscita(UBound(Uscita) + 1)
        End If
    
    Case "MAGGIORE"
    
        If (Valore > CDbl(Maggiore)) Then
            Uscita(UBound(Uscita)).X = Punto.X
            Uscita(UBound(Uscita)).Y = Punto.Y
            ReDim Preserve Uscita(UBound(Uscita) + 1)
        End If
    
    Case "ESTERNI"
    
        If (Valore > CDbl(Maggiore) Or Valore < CDbl(Minore)) Then
            Uscita(UBound(Uscita)).X = Punto.X
            Uscita(UBound(Uscita)).Y = Punto.Y
            ReDim Preserve Uscita(UBound(Uscita) + 1)
        End If
    
    Case "INTERNI"
    
        If (Valore < CDbl(Minore) And Valore > CDbl(Maggiore)) Then
            Uscita(UBound(Uscita)).X = Punto.X
            Uscita(UBound(Uscita)).Y = Punto.Y
            ReDim Preserve Uscita(UBound(Uscita) + 1)
        End If
    
    End Select

End Function 'Run_Sottoinsieme_Esegui_Test()





Private Sub Run_Sottoinsieme(ByRef Testx As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_POST_PROCESSING, Test As TagTEST_TOTAL)
    On Error GoTo err:
    Dim i As Long
    Dim TipoTest As String
    
    
    If (Test.Nome_Tipo_Test <> "SOTTOINSIEME") Then
        Exit Sub
    End If
    

    TipoTest = Run_Sottoinsieme_Tipo_Test(Test.Test_Sottoinsieme.Parametro_Maggiore_X, Test.Test_Sottoinsieme.Parametro_Minore_X)
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "TIPO TEST : " + TipoTest + vbLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "VALORE 1: " + Test.Test_Sottoinsieme.Parametro_Minore_X + vbLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "VALORE 2: " + Test.Test_Sottoinsieme.Parametro_Maggiore_X + vbLf
    Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)

    
    
    ReDim Test.Test_Sottoinsieme.Temp(0)
    ' sottoinsieme asse X
    i = 0
    Do While (i <= UBound(Context.Punto))
        
        
'        Run_Sottoinsieme_Esegui_Test(
'                                    Valore As Double,
'                                    Minimo As String,
'                                    Massimo As String,
'                                    TipoTest As String,
'                                    Punto As TagPuntiAcquisizione,
'                                    Uscita() As TagPuntiAcquisizione)
        
        
        Run_Sottoinsieme_Esegui_Test Context.Punto(i).X, _
                                        Test.Test_Sottoinsieme.Parametro_Minore_X, _
                                        Test.Test_Sottoinsieme.Parametro_Maggiore_X, _
                                        TipoTest, _
                                        Context.Punto(i), _
                                        Test.Test_Sottoinsieme.Temp

        i = i + 1
    Loop
    
    
    If (UBound(Test.Test_Sottoinsieme.Temp) > 0) Then
        ReDim Preserve Test.Test_Sottoinsieme.Temp(UBound(Test.Test_Sottoinsieme.Temp) - 1)
    End If
    
    TipoTest = Run_Sottoinsieme_Tipo_Test(Test.Test_Sottoinsieme.Parametro_Maggiore_Y, Test.Test_Sottoinsieme.Parametro_Minore_Y)
    
    ' sottoinsieme asse Y
    i = 0
    Do While (i <= UBound(Test.Test_Sottoinsieme.Temp))
        
        Run_Sottoinsieme_Esegui_Test Test.Test_Sottoinsieme.Temp(i).Y, _
                                        Test.Test_Sottoinsieme.Parametro_Minore_Y, _
                                        Test.Test_Sottoinsieme.Parametro_Maggiore_Y, _
                                        TipoTest, _
                                        Test.Test_Sottoinsieme.Temp(i), _
                                        Context.Uscita

    
        i = i + 1
    Loop
    
    If (UBound(Context.Uscita) > 0) Then
        ReDim Preserve Context.Uscita(UBound(Context.Uscita) - 1)
    End If
    
        
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
    Exit Sub
err:
    Context.State = eError
    
    
End Sub 'Run_Sottoinsieme()



Private Function CreaDirectory(Infos As TagInfosSingleTest, ByVal StartDirectory As String) As String
    
    On Error Resume Next:
    
    Dim PathDirectory  As String
    Dim subfolder As Variant
    Dim TempStr As String
    Dim i As Long
    
    
    subfolder = Split(Infos.codicePRodotto.Caption, "\")
    
    TempStr = subfolder(0) + "\"
    TempStr = StartDirectory + "\" + TempStr
    
    MkDir (TempStr)
    
    i = 1
    Do While (i <= UBound(subfolder))
        
        TempStr = TempStr + subfolder(i) + "\"
    
        MkDir (TempStr)
        i = i + 1
    Loop
    
    
    
    
        
'    CreaDirectory = StartDirectory + "\" + Infos.CodiceProdotto + "\"
'
'    MkDir (CreaDirectory)
    Exit Function
    
err:
    Debug.Print err.Description
End Function 'CreaDirectory()

Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Dim FileTest As String
    FileTest = FileText(Get_Nome_Assoluto_File(Test.Parameter(eNomeFile)))
    
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFileTipoTest                    = " + Test.Parameter(eNomeFileTipoTest) + vbCrLf
        Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + FileTest + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
        

End Sub



