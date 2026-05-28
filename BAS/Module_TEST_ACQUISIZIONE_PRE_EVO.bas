Attribute VB_Name = "Module_TEST_ACQUISIZIONE_PRE_EVO"

'**************************************************************
'       M O D U L O :    Module_TEST_ACQUISIZIONE_PRE_EVO.bas
'**************************************************************

Private Declare Function GetTickCount Lib "kernel32" () As Long

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
    ePresetValue = 6
    ePresetValue_wait = 7
End Enum


'----------------
' PARAMETRI TEST
'----------------
Private Enum eTestParameter
    eNomeFile = 1


    eMax
End Enum


' struttura dati delle informazioni necessarie per la configurazione del file di uscita
Private Type TagTuplaFiles
    IdScheda As Integer
    TipoVariabile As String
    AddrVariabile As Integer
    NOME As String
End Type


' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagPuntiAcquisizione
    Y() As Double
End Type


Private Type TagPresetValue
     Tupla As TagTuplaFiles
     Value As Double
     Delay As Long
End Type


Public Type TagTEST_ACQUISIZIONE_PRE_EVO
    ErrorString As String
    bOldOn As Boolean                               ' difu del comando di start del modulo
    Return As Integer                               ' Stato del modulo
    State As Integer
    NomeFile As String                              ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean                           ' Flag segnalazione primo avvio del programma
    Timer As Long
    
    FileTest As String                              ' Contenuto del File di settings

    CURVE_XY As TagChannelRampa                     ' File di configurazione della rampa usata per effettuare l'acquisizione
    CURVE_COMANDO As TagChannel                     ' File di configurazione delle curve comando mA in funzione di V teorico
    
    cfg_Punto() As TagTuplaFiles                    ' Vettore di configurazione dei test
    Punto() As TagPuntiAcquisizione                 ' Punti della acquisizione
    NumeroPuntiAcquisiti As Long                    ' numero di punti acquisiti
    Tempo_sec_start As Double
    OldCommandInMilliVolts As Double

    
    PresetValue() As TagPresetValue                 ' Variabili da settare prima di attivare la sequenza
    PresetIndex As Long
    
    
    
    NomeFileRampaXY As String
    NomeFileCurvaComando As String
    
    Var_Command_Scheda As String
    Var_Command_Tipo As String
    Var_Command_Indirizzo As String
    
    Var_State_Scheda As String
    Var_State_Tipo As String
    Var_State_Indirizzo As String
    
    Var_Nome_File_Uscita() As String
    
    FileNameSettingVariables As String
    
    
    
    
    
    
    
    TimeA As Long
    TimeB As Long
    Delta As Long
    DeltaMax As Long
End Type


'----------------
' RampSettings()
'----------------
' Funzione di inizializzazione dei settaggi del grafico. Carica le curve limite
Private Sub RampSettings(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)

    On Error GoTo err:

    Dim bUpdate As Boolean
    Dim i As Integer

    ' RAMPE
    ' -----
   

        ' Inizializzazione dei file della rampa da applicare
       ' If (Update_CURVE_XY(Context.CURVE_XY, Settings.FolderRampeXY + "\" + Context.NomeFileRampaXY) = True) Then
            If (Update_CURVE_XY(Context.CURVE_XY, Get_Nome_Assoluto_File(Context.NomeFileRampaXY, IndiceSezione, Infos.Label1.Caption)) = True) Then
                
                Context.ErrorString = "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI RAMPE " + Get_Nome_Assoluto_File(Context.NomeFileRampaXY, IndiceSezione, Infos.Label1.Caption)
                Prepare_MsgBox Infos, Context.ErrorString, vbCritical
                
                Context.State = eError
            
            End If
       ' End If:

        bUpdate = True
        
    'End If

    ' CURVE COMANDO
    ' -------------
   
        ' Inizializzazione dei comandi da applicare
      '  If (Update_CURVECOMANDO(Context.CURVE_COMANDO, Settings.FolderFileCurveComando + "\" + Context.NomeFileCurvaComando) = True) Then
            
            If (Update_CURVECOMANDO(Context.CURVE_COMANDO, Get_Nome_Assoluto_File(Context.NomeFileCurvaComando, IndiceSezione, Infos.Label1.Caption)) = True) Then
            
                Context.ErrorString = "ERRORE DURANTE L'APERTURA DEL FILE DELLA CURVA DI COMANDO"
                Prepare_MsgBox Infos, Context.ErrorString, vbCritical
                
                Context.State = eError
            End If
            
       ' End If

        bUpdate = True
        
    'End If

    Exit Sub

err:
    Debug.Print err.Description
    
End Sub 'RampSettings()


'-------------------------
' "INFO_ACQUISIZIONE_PRE_EVO"
'-------------------------
Private Sub INFO_ACQUISIZIONE_PRE_EVO(Infos As TagInfosSingleTest, Optional bClear As Boolean = False)

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


Private Sub TEST_ACQUISIZIONE_PRE_EVO_INIT_STRUCT(Context1 As TagTEST_ACQUISIZIONE_PRE_EVO)
    
    Context1.bFirstTime = False
    Context1.bOldOn = False
    Context1.Delta = 0
    Context1.DeltaMax = 0
    Context1.FileTest = ""
    Context1.NomeFile = ""
    Context1.NumeroPuntiAcquisiti = 0
    Context1.OldCommandInMilliVolts = 0
    
    Context1.Return = 0
    Context1.State = 0
    Context1.Tempo_sec_start = 0
    Context1.TimeA = 0
    Context1.TimeB = 0
    Context1.Timer = 0
    Context1.State = eEnd
    
End Sub


'-------------------------
' "TEST_ACQUISIZIONE_PRE_EVO"
'-------------------------
Public Sub TEST_ACQUISIZIONE_PRE_EVO(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagTEST_ACQUISIZIONE_PRE_EVO, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "ACQUISIZIONE_PRE_EVO") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        INFO_ACQUISIZIONE_PRE_EVO Infos
        
        
        If (bStop = True And bOn = True) Then
        
            TEST_ACQUISIZIONE_PRE_EVO_INIT_STRUCT Context1
            INFO_ACQUISIZIONE_PRE_EVO Infos, True
            
            
            
        ElseIf (bStop = True And Context1.State <> eEnd) Then
        
            Context.ErrorString = "STOP DA PARTE DELL'UTENTE"
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
    
End Sub 'TEST_ACQUISIZIONE_PRE_EVO()


'-----------------
' Task del modulo
'-----------------
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    
    Select Case Context.State
        Case eInit: Init Test, Infos, IndiceSezione, Context
        Case eStart: Start Test, Infos, IndiceSezione, Context
        Case eRun: Run Test, Infos, IndiceSezione, Context
        Case esavefile: SaveFile Test, Infos, IndiceSezione, Context
        Case eEnd: Fine Test, Infos, IndiceSezione, Context
        Case eError: ErrorTest Test, Infos, IndiceSezione, Context
        Case ePresetValue: PresetValue Test, Infos, IndiceSezione, Context
        Case ePresetValue_wait: PresetValue_wait Test, Infos, IndiceSezione, Context
        Case Else: ErrorTest Test, Infos, IndiceSezione, Context
        
    End Select
    

    
End Sub 'Task()





Private Sub Open_input_file(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    On Error GoTo err:
    Dim FileTest As String
    Dim Righe As Variant
    Dim i As Long
    Dim NOME As String
    Dim Valore As String
    
    FileTest = FileText(Get_Nome_Assoluto_File(Test.Parameter(eNomeFile)))
    Righe = Split(FileTest, vbLf)
    
    ReDim Context.Var_Nome_File_Uscita(0)
    ReDim Context.PresetValue(0)
    
    i = 0
    Do While (i <= UBound(Righe))
    
        Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
        Get_Coppia_Variabile_Valore Righe(i), NOME, Valore

        Select Case NOME
        Case "VAR_COMMAND_SCHEDA": Context.Var_Command_Scheda = Valore
        Case "VAR_COMMAND_TIPO": Context.Var_Command_Tipo = Valore
        Case "VAR_COMMAND_INDIRIZZO": Context.Var_Command_Indirizzo = Valore
        Case "VAR_STATE_SCHEDA": Context.Var_State_Scheda = Valore
        Case "VAR_STATE_TIPO": Context.Var_State_Tipo = Valore
        Case "VAR_STATE_INDIRIZZO": Context.Var_State_Indirizzo = Valore
        Case "NOMEFILECURVACOMANDO": Context.NomeFileCurvaComando = Valore
        Case "NOMEFILERAMPAXY": Context.NomeFileRampaXY = Valore
        Case "FILENAMESETTINGVARIABLES": Context.FileNameSettingVariables = Valore
        
        Case "PRESET_VARIABLE":
            UpdatePresetValue Context.PresetValue(UBound(Context.PresetValue)), Valore, IndiceSezione
            ReDim Preserve Context.PresetValue(UBound(Context.PresetValue) + 1)
            
        Case "NOME_FILE_USCITA":
            Context.Var_Nome_File_Uscita(UBound(Context.Var_Nome_File_Uscita)) = Valore
            ReDim Preserve Context.Var_Nome_File_Uscita(UBound(Context.Var_Nome_File_Uscita) + 1)
        End Select
        i = i + 1
    Loop
        
    Exit Sub
err:
Debug.Print err.Description

End Sub



Private Sub PresetValue_wait(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    
    
    If (Context.Timer < GetTickCountEvo) Then
        Context.State = ePresetValue
'        If (Context.PresetIndex > UBound(Context.PresetValue)) Then
'            Context.State = eStart
'        Else
'            Context.State = ePresetValue
'        End If
    End If
        
        
    


End Sub


Private Sub PresetValue(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
        
    
    Infos.RichTextBox1.Text = CStr(Context.PresetValue(Context.PresetIndex).Tupla.IdScheda) + "." + Context.PresetValue(Context.PresetIndex).Tupla.TipoVariabile + "." + CStr(Context.PresetValue(Context.PresetIndex).Tupla.AddrVariabile) + " = " + CStr(Context.PresetValue(Context.PresetIndex).Value) + vbLf + Infos.RichTextBox1.Text
    SetCalibratedValue Context.PresetValue(Context.PresetIndex).Tupla.IdScheda, Context.PresetValue(Context.PresetIndex).Tupla.TipoVariabile, Context.PresetValue(Context.PresetIndex).Tupla.AddrVariabile, Context.PresetValue(Context.PresetIndex).Value
    
    Context.PresetIndex = Context.PresetIndex + 1
    
    If (Context.PresetIndex > UBound(Context.PresetValue)) Then
        Context.State = eStart
    Else
        Context.State = ePresetValue_wait
        Context.Timer = GetTickCountEvo + Context.PresetValue(Context.PresetIndex).Delay
    End If
    
    
    'Context.Timer = GetTickCountEvo + Context.PresetValue(Context.PresetIndex).Delay
    

    
    
        
        
    'Context.State = ePresetValue_wait

    


End Sub




Private Sub UpdatePresetValue(ByRef PresetValue As TagPresetValue, Valore As String, IndiceSezione As Integer)
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
    
    PresetValue.Value = CDbl(Dati(1))
    

End Sub











'------
' Init
'------
' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    
    Context.Return = eTEST_BUSY
    
    ' Apro il file di ingresso
    Open_input_file Test, Infos, IndiceSezione, Context
    
     
    Context.State = ePresetValue
    
    RampSettings Test, Infos, IndiceSezione, Context
    
    Context.NumeroPuntiAcquisiti = 0

    ' Prelevo il contenuto del file di configurazione
    Context.FileTest = FileText(Get_Nome_Assoluto_File(Context.FileNameSettingVariables))

    ' Inizializzo la struttura in base alla configurazione del file
    Init_NomeTest Test, Infos, IndiceSezione, Context

    
    If (CInt(Context.Var_Command_Scheda) = -1) Then
        Context.Var_Command_Scheda = CStr(IndiceSezione + 1)
    Else
        Context.Var_Command_Scheda = CInt(Context.Var_Command_Scheda)
    End If
    
    
    If (CInt(Context.Var_State_Scheda) = -1) Then
        Context.Var_State_Scheda = CStr(IndiceSezione + 1)
    Else
        Context.Var_State_Scheda = CInt(Context.Var_State_Scheda)
    End If

    Context.Timer = GetTickCountEvo
    



    ReDim Context.Punto(Context.NumeroPuntiAcquisiti)
    
    Infos.RichTextBox1.Text = ""
    Infos.RichTextBox1.BackColor = vbWhite
    
    
    Prepare_Textbox_Parametri Test, Infos
    
    Context.PresetIndex = 0
    Context.Timer = GetTickCountEvo + Context.PresetValue(Context.PresetIndex).Delay
End Sub 'Init()


'-------
' Start
'-------
Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    Dim i As Long
    
    ' INFO
    '------
    Infos.RichTextBox1.Text = "NOME FILE CURVA COMANDO    = " + Context.NomeFileCurvaComando + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "NOME FILE RAMPA APPLICATA  = " + Context.NomeFileRampaXY + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = ">>> MODULO   ACQUISIZIONE  PRE  EVO <<<" + vbLf + Infos.RichTextBox1.Text
    
    
        'Init Test, Infos, IndiceSezione, Context
    
    

    Context.bFirstTime = True
    
    Context.Tempo_sec_start = GetTickCountEvo / 1000
    
    Context.State = eRun
    

End Sub 'Start()


'------
' Run()
'------
Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    
    Dim Tempo_sec As Double
    Dim CommandInMilliVolts As Double
    Dim State As MltState
    Dim bPunti_Y_Validi As Boolean
    Dim Percentual As Integer
    Dim OldTimer As Long
    
    On Error Resume Next
    
    Context.Timer = GetTickCountEvo + 100
    OldTimer = GetTickCountEvo - 1
    
    Do While (Context.Timer > GetTickCountEvo And Context.State = eRun)
        
        If (OldTimer <= GetTickCountEvo) Then
            
            OldTimer = GetTickCountEvo + 2
            
            PPDO_Manager
        
        
            If (Context.bFirstTime = True) Then
                'Context.Tempo_sec_start = GetTickCountEvo / 1000
            End If
            
            Tempo_sec = GetTickCountEvo / 1000 - Context.Tempo_sec_start
            
            
            '---------------
            ' RAMPA "t-->V" (tempo-->volt)
            '---------------
            ' Prelevo il comando dalla rampa applicata
            If (Get_CURVE_XY_VALUE(Context.CURVE_XY, Tempo_sec, CommandInMilliVolts) = eERROR_Rampe) Then
                
                Context.State = esavefile
            
            Else
                '--------------------------------
                ' CURVA COMANDO "V-->Percentual"
                '--------------------------------
                SetCalibratedValue 0, "INTERNAL", 1, CommandInMilliVolts
                
                
                If (Get_CURVECOMANDO_VALUE(Context.CURVE_COMANDO, CommandInMilliVolts / 1000, State, Percentual) = eError_CURVE_COMANDO_CAN) Then
                    
                    ' LOG
                    '-----
 
                    Context.ErrorString = "LA FUNZIONE Get_CURVECOMANDO_VALUE() HA RITORNATO UN ERRORE"
                    Context.State = eError
                Else
                    
                    UPDATE_OUTPUT Test, Context, Percentual, State
                    
                End If
            End If
                    
            If (Context.State <> esavefile) Then
            
                ' PANDA TODO
               
                
                
                If (Context.bFirstTime = True) Then
                    Context.bFirstTime = False
                    
                    Memorizza_Dati Context.cfg_Punto, Context.Punto, IndiceSezione
                    
                Else
                    
                    ReDim Preserve Context.Punto(UBound(Context.Punto) + 1)
                    Memorizza_Dati Context.cfg_Punto, Context.Punto, IndiceSezione
                    
                End If
                                                            
        
            End If 'END (Context.State <> eSaveFile)
        End If
    Loop
    'End If
        
    Context.OldCommandInMilliVolts = CommandInMilliVolts
    
    
End Sub 'Run()


Private Sub UPDATE_OUTPUT(ByRef Test As TagSingoloTest, Context As TagTEST_ACQUISIZIONE_PRE_EVO, ByVal Percentual As Integer, ByVal State As Integer)

    Dim PPDO_IndexMessage As Integer
    Dim PPDO_AdcIndexIntoMessage As Integer
    Dim Output As Double
        
    
    
    SetCalibratedValue Context.Var_Command_Scheda, _
                        Context.Var_Command_Tipo, _
                        CInt(Context.Var_Command_Indirizzo), _
                        Percentual
    
    
    SetCalibratedValue Context.Var_State_Scheda, _
                        Context.Var_State_Tipo, _
                        CInt(Context.Var_State_Indirizzo), _
                        State
    
    
End Sub 'UPDATE_OUTPUT()


'-----------
' SaveFile()
'-----------
Private Sub SaveFile(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    On Error GoTo err:
    ' INFO
    '------
    Infos.RichTextBox1.Text = vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = "Salvataggio Curva Acquisizione X-Y in File.csv" + vbLf + Infos.RichTextBox1.Text
    Infos.RichTextBox1.Text = vbLf + Infos.RichTextBox1.Text
    ' LOG
    '-----

                

    SalvaPunti Test, Infos, Context, IndiceSezione
    
    SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    
err:
End Sub 'SaveFile()


'------
' Fine
'------
Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)

    If (Infos.CWGraph1.PlotAreaColor = vbRed) Then
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
    
End Sub 'Fine()


'------------
' ErrorTest()
'------------
Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    
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

End Sub 'ErrorTest()


'-------------
' SalvaPunti()
'-------------
Private Sub SalvaPunti(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, Context As TagTEST_ACQUISIZIONE_PRE_EVO, IndiceSezione As Integer)
    
    Dim FileName As String
    Dim i As Long
    Dim Index As Long
    
    On Error GoTo err:
        
            
    Index = 0
    
    
    
    Do While (Index < UBound(Context.Var_Nome_File_Uscita))
        
        FileName = Get_Nome_Assoluto_File(Context.Var_Nome_File_Uscita(Index), IndiceSezione, Infos.Label1.Caption)
        Open FileName For Output As #23    ' per file POST di lavoro
        
        SaveFile_Create_FirstRow 23, Context.cfg_Punto  ' per file ALL
        i = 0
        Do While (i < UBound(Context.Punto))
            
            SaveFile_Create_Row 23, Context.Punto(i)        ' per file POST di lavoro
                       
            i = i + 1
        Loop
        
        Index = Index + 1
        
        Close #23
    Loop
    
    '----------------------------
        SalvaRisultatiTest Settings.FolderLOG, Infos, Infos.RichTextBox1.Text
    Context.State = eEnd
Exit Sub

err:

    Context.ErrorString = "ERRORE DURANTE IL SALVATAGGIO DELLA CURVA " + FileName + " " + err.Description
    Prepare_MsgBox Infos, Context.ErrorString, vbCritical
    Context.State = eError
    
    
    Close #23
End Sub 'SalvaPunti()



Private Function CreaDirectory(Infos As TagInfosSingleTest, ByVal StartDirectory As String) As String
    
    On Error GoTo err:
    
    Dim PathDirectory  As String
    Dim subfolder As Variant
    Dim TempStr As String
    Dim i As Long
    
    
    subfolder = Split(Infos.codicePRodotto, "\")
    
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



Public Function DirExists(ByVal Path As String) As Boolean

    On Error Resume Next
    
    'Legge l'attributo e si assicura che si tratti di una directory
    FileExists = GetAttr(Path) And vbDirectory
    
    'Se avviene un errore la Function restituisce False
    
End Function 'DirExists()



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



' Funzione di inizializzazione della struttura del modulo attraverso la lettura del file di configurazione specificio degli ingressi da loggare
Private Sub Init_NomeTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagTEST_ACQUISIZIONE_PRE_EVO)
    
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    
    Dim NOME As String
    Dim Valore As String
    
    Righe = Split(Context.FileTest, vbLf)
    
        
    i = 0
    ReDim Context.cfg_Punto(0)
    Do While (i <= UBound(Righe))
        Righe(i) = Rimuovi_Commento_Da_Stringa(Righe(i), "//")
        
        Get_Tupla Righe(i), Context.cfg_Punto, "."
        
        i = i + 1
    Loop
    
    If (UBound(Context.cfg_Punto) >= 1) Then
        ReDim Preserve Context.cfg_Punto(UBound(Context.cfg_Punto) - 1)
    End If

End Sub


Private Function Get_Tupla(ByVal Riga As String, ByRef Tupla() As TagTuplaFiles, Optional Separatore As String = "=")
    
    Dim Dati As Variant
    Dim i As Integer


    Riga = Replace(Riga, vbLf, "")
    Riga = Replace(Riga, vbCr, "")
    Dati = Split(Riga, Separatore)
    
    '
    If (UBound(Dati) <> 3) Then
        Exit Function
    Else
        Tupla(UBound(Tupla)).IdScheda = CLng(Dati(0))
        Tupla(UBound(Tupla)).TipoVariabile = Dati(1)
        Tupla(UBound(Tupla)).AddrVariabile = CLng(Dati(2))
        Tupla(UBound(Tupla)).NOME = Dati(3)
        ReDim Preserve Tupla(UBound(Tupla) + 1)
    
    End If
    
End Function

' Questa funzione legge i dati dal sistema e li scrive nella struttura dinamica dei punti acquisiti
Private Function Memorizza_Dati(ByRef Tupla() As TagTuplaFiles, Punto() As TagPuntiAcquisizione, ByVal IndiceSezione As Integer)
    
    Dim i As Long
    
    ' Modifico la dimensione del vettore dei dati acquisiti
    ReDim Punto(UBound(Punto)).Y(UBound(Tupla))
    
    i = 0
    
    
    Do While (i <= UBound(Tupla))
    
        If (Tupla(i).IdScheda = -1) Then
            Tupla(i).IdScheda = IndiceSezione + 1
        End If
        
    
        Punto(UBound(Punto)).Y(i) = GetCalibratedValue_EVO(Tupla(i).IdScheda, Tupla(i).TipoVariabile, Tupla(i).AddrVariabile)
        i = i + 1
    Loop
            
End Function



Private Function SaveFile_Create_FirstRow(IndexFile As Integer, ByRef Tupla() As TagTuplaFiles)
    
    Dim Row As String
    Dim i As Long
    Dim bFirstTime As Boolean
    
    bFirstTime = True
    i = 0
    Do While (i <= UBound(Tupla))
        If (bFirstTime = True) Then
            
            bFirstTime = False
            Row = Tupla(i).NOME
        
        Else
        
            Row = Row + ";" + Tupla(i).NOME
            
        End If
        i = i + 1
    Loop

    Print #IndexFile, Row
    
End Function



'Private Function SaveFile_Create_FirstRow_PcDc(IndexFile As Integer, ByRef Tupla() As TagTuplaFiles)
'
'    Dim Row As String
'    Dim i As Long
'    Dim bFirstTime As Boolean
'
'    bFirstTime = True
'    i = 1 ' salto il "tempo" (con i = 0 nella 1° colonna ho il Tempo)
'    Do While (i <= UBound(Tupla))
'        If (bFirstTime = True) Then
'
'            bFirstTime = False
'            Row = Tupla(i).NOME
'
'        Else
'
'            Row = Row + ";" + Tupla(i).NOME
'
'        End If
'        i = i + 1
'    Loop
'
'    Print #IndexFile, Row
'
'End Function


Private Function SaveFile_Create_Row(IndexFile As Integer, Punto As TagPuntiAcquisizione)

    Dim Row As String
    Dim i As Long
    Dim bFirstTime As Boolean
    
    bFirstTime = True
    i = 0
    Do While (i <= UBound(Punto.Y))
        If (bFirstTime = True) Then
            
            bFirstTime = False
            Row = CStr(Punto.Y(i))
        
        Else
        
            Row = Row + ";" + CStr(Punto.Y(i))
            
        End If
        
        i = i + 1
    Loop
    
    Print #IndexFile, Row

End Function


'Private Function SaveFile_Create_Row_PcDc(IndexFile As Integer, Punto As TagPuntiAcquisizione)
'
'    Dim Row As String
'    Dim i As Long
'    Dim bFirstTime As Boolean
'
'    bFirstTime = True
'    i = 1 ' salto il "tempo" (con i = 0 nella 1° colonna ho il Tempo)
'    Do While (i <= UBound(Punto.y))
'        If (bFirstTime = True) Then
'
'            bFirstTime = False
'            'Row = CStr(Punto.y(i))
'            Row = Format(CStr(Punto.y(i)), "##0.000")
'        Else
'
'            'Row = Row + vbTab + ";" + vbTab + CStr(Punto.y(i))
'            Row = Row + ";" + Format(CStr(Punto.y(i)), "##0.000")
'
'        End If
'
'        i = i + 1
'    Loop
'
'    Print #IndexFile, Row
'
'End Function



Private Function GetDelayValue(ByRef PresetValue As TagPresetValue, ByRef Valore As String, IndiceSezione As Integer) As String
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


Private Sub Prepare_Textbox_Parametri(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest)
    Dim FileTest As String
    Infos.RichTextBox1.Text = "LISTA PARAMETRI" + vbCrLf _
                                              + "-------------------------------" + vbCrLf + _
                                              CStr(Now) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "eNomeFile                      = " + Test.Parameter(eNomeFile) + vbCrLf
    
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + FileTest + vbCrLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------------" + vbCrLf
    
End Sub

