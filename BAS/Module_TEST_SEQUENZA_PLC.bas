Attribute VB_Name = "Module_TEST_SEQUENZA_PLC"

'**************************************************************
'       M O D U L O :    Module_TEST_SEQUENZA_PLC.bas
'**************************************************************


'**************************************************************
'       M O D U L O :    Module_SEQUENZA_PLC.bas
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
    NomeFileIngresso = 1        ' Nome del file di ingresso
    Bloccante = 2               ' Abilita la modalità bloccante o meno del test globale
    eMax
End Enum


' Struttura dati usata per definire e gestire i punti dell'acquisizione
Public Type TagIstruzione
    NOME As String
    par() As String
    
    RIGA As String
End Type


 Type TagTupla
    ID As Integer
    TYPE As String
    Addr As Integer
 End Type
 
 
 Public Type TagStackCall
    R(18) As Double             ' Registri interni dell'ALU disponibili
    s(18) As String             ' Stringhe interne
    p(18) As String             ' Stringhe interne
    LR As Long                  ' Indice relativo alla gestione della chiamate alle funzioni
 End Type
 

Public Type TagAlu

    ' ALU
    n As Boolean                        ' Negative result from ALU
    Z As Boolean                        ' Zero Sesult from ALU
    c As Boolean                        ' ALU operation cauded carry ????
    R(18) As Double                     ' Registri interni dell'ALU disponibili
    s(18) As String                     ' Stringhe interne
    p(18) As String                     ' Stringhe interne
    PC As Long                          ' Program counter
    
    SP_POINTER As Long
    SP_STRINGHE As Long
    SP As Long                          ' Indice relativo allo stack pointer
    LR As Long                          ' Indice relativo alla gestione della chiamate alle funzioni
    
    SP_CALL As Long
    STACK_CALL(100) As TagStackCall     ' Vettore delle chiamate a funzioni
    STACK(100) As Long                  ' Area dello stack
    STACK_POINTER(100) As String        ' Area dello stack
    STACK_STRINGHE(100) As String
    CurrentFunction As String           ' Nome della funzione corrente
    Timer As Long                       ' Timer interno
    
    IndiceSezione As Integer            ' Indice della sezione di riferimento
    Infos As TagInfosSingleTest
End Type


Public Type TagDefine
    Parola As String                    ' Vettore delle definizioni
    Valore As String
End Type

Public Type TagPreprocessor
    IndiceDefine As Long
    DEFINE() As TagDefine
    
    IndiceIfDef As Long
    IFDEF() As Boolean
End Type


Public Type TagSEQUENZA_PLC
    bOldOn As Boolean                   ' difu del comando di start del modulo
    Return As Integer                   ' Stato del modulo
    State As Integer
    
    NomeFile As String                  ' Nome del file da utilizzare per la creazione della curva XY
    bFirstTime As Boolean               ' Flag segnalazione primo avvio del programma
    
    
    ' Codice da eseguire
    PASSI() As TagIstruzione            ' Vettore dei passi da eseguire
    ALU As TagAlu
    
    PREPROCESSOR As TagPreprocessor
    
    FileTest As String                  ' Nome del test da applicare al post processing

End Type


'------------------------
' "INFO_SEQUENZA_PLC"
'------------------------
Private Sub INFO_SEQUENZA_PLC(Infos As TagInfosSingleTest)

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
    
End Sub 'INFO_SEQUENZA_PLC()


' Questa funzione ricerca il punto di ingresso del file
' in pratica cerca la funzione denominata main
' true con errore
Private Function SEQUENZA_PLC_INIT_PC(ALU As TagAlu, PASSI() As TagIstruzione) As Boolean
On Error GoTo err:
    Dim i As Long
    Dim bReturn As Boolean
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    i = 0
    
    Do While (bExit = False)
        If (IsArrayInitialized(PASSI(i).par) = True) Then
            
            If (PASSI(i).NOME = "FUNCTION" And PASSI(i).par(0) = Chr(34) + "main" + Chr(34)) Then
                bExit = True
                bFind = True
                
            Else
                i = i + 1
            End If
            
        Else
            i = i + 1
        End If
            
        If (i > UBound(PASSI)) Then
            bExit = True
        End If
    
        
    Loop
    
    
    If (bFind = True) Then
        ALU.PC = i
        SEQUENZA_PLC_INIT_PC = False
    Else
        SEQUENZA_PLC_INIT_PC = True
    End If
    
    Exit Function
err:
    SEQUENZA_PLC_INIT_PC = True
    Debug.Print err.Description
End Function



Private Sub SEQUENZA_PLC_INIT_ALU(ALU As TagAlu)
    ALU.n = False
    ALU.Z = False
    ALU.c = False
    ALU.PC = 0
    
    
    ALU.SP_CALL = 0
    
    ALU.SP = 0
    ALU.LR = 0
    ALU.SP_POINTER = 0
    ALU.SP_STRINGHE = 0
    
    i = 0
    Do While (i < UBound(ALU.R))
        ALU.R(i) = 0
        i = i + 1
    Loop
    
    i = 0
    Do While (i < UBound(ALU.s))
        ALU.s(i) = 0
        i = i + 1
    Loop
    
    i = 0
    Do While (i < UBound(ALU.p))
        ALU.p(i) = 0
        i = i + 1
    Loop
    
    
        
    
End Sub

Private Sub SEQUENZA_PLC_INIT_STRUCT(Context1 As TagSEQUENZA_PLC)
    Dim i As Long
    
    Context1.bFirstTime = False
    Context1.bOldOn = False
    Context1.FileTest = ""
    Context1.Return = 0
    Context1.State = 0
    
    Context1.PREPROCESSOR.IndiceDefine = 0
    Context1.PREPROCESSOR.IndiceIfDef = 0
    
    ReDim Context1.PREPROCESSOR.IFDEF(0)
    ReDim Context1.PREPROCESSOR.DEFINE(0)
    
    SEQUENZA_PLC_INIT_ALU Context1.ALU

    
    
    
    
End Sub

'------------------------
' "SEQUENZA_PLC"
'------------------------
Public Sub SEQUENZA_PLC(ByRef Test As TagSingoloTest, IndiceSezione As Integer, Infos As TagInfosSingleTest, Context1 As TagSEQUENZA_PLC, ByRef bOn As Boolean, ByRef bStop As Boolean, ByRef bReturn As Integer)
   
    If (Test.ID <> "SEQUENZA_PLC") Then
        Context1.Return = eTEST_STOP
        Context1.State = eEnd
    Else
                
        INFO_SEQUENZA_PLC Infos
        
        If (bStop = True And bOn = True) Then
            SEQUENZA_PLC_INIT_STRUCT Context1
            Set Context1.ALU.Infos.CWGraph1 = Infos.CWGraph1
            Set Context1.ALU.Infos.RichTextBox1 = Infos.RichTextBox1
            Set Context1.ALU.Infos.Label1 = Infos.Label1
            
        ElseIf (bStop = True And Context1.State <> eEnd) Then
            Context1.State = eError
        
        ElseIf (bOn = True And Context1.bOldOn = False) Then        ' fronte di salita del test ( AVVIO )
            Init_Test Test, Infos, IndiceSezione, Context1               ' init delle variabili del modulo
        
        End If
        
        Task Test, Infos, IndiceSezione, Context1                   ' Task del modulo
        
        If (bStop = True) Then
             Context1.bOldOn = False
        Else
            Context1.bOldOn = bOn                     ' memorizzo il fronte
        End If

        bReturn = Context1.Return
    End If
    
End Sub 'SEQUENZA_PLC()


'-----------------
' Task del modulo
'-----------------
Private Sub Task(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC)
    
    Select Case Context.State
        Case eInit: Init Test, Infos, IndiceSezione, Context
        Case eStart: Start Test, Infos, IndiceSezione, Context
        Case eRun: Run Test, Infos, IndiceSezione, Context
        Case eEnd: Fine Test, Infos, IndiceSezione, Context
        Case eError: ErrorTest Test, Infos, IndiceSezione, Context
        Case Else: ErrorTest Test, Infos, IndiceSezione, Context
    End Select
    
End Sub 'Task()




Private Sub Init_Test(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC)


    Init Test, Infos, IndiceSezione, Context

    Context.bFirstTime = True
    Context.FileTest = FileText(Get_Nome_Assoluto_File(Test.Parameter(NomeFileIngresso), IndiceSezione, Infos.Label1.Caption))
    Infos.RichTextBox1.Text = ""
    Infos.RichTextBox1.BackColor = vbWhite
    
    
    Context.PREPROCESSOR.IndiceDefine = 0
    Context.PREPROCESSOR.IndiceIfDef = 0
    
    
    ReDim Context.PREPROCESSOR.DEFINE(0)
    ReDim Context.PREPROCESSOR.IFDEF(0)
    
    If (Init_PASSI(Test, Infos, IndiceSezione, Context) = True) Then
        Infos.RichTextBox1.Text = " ERROR INIT PASSI " + Context.FileTest + " " + err.Description
        Context.State = eError
    End If
    
    
    
End Sub
'------
' Init
'------
' Questa funzione inizializza la memoria del frame
Private Sub Init(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC)
    On Error GoTo err:
    
    Dim tempodati As Long
    
    tempodati = vbGreen
    
    If (UBound(Test.Parameter) < Bloccante) Then
        Context.Return = eTEST_BUSY
    ElseIf (Test.Parameter(Bloccante) = "NON BLOCCANTE") Then
        Context.Return = eTEST_BUSY_NON_BLOCCANTE
    Else
        Context.Return = eTEST_BUSY 'eTEST_BUSY_NON_BLOCCANTE
    End If
    
    
     
    Context.State = eStart
    
'    Context.bFirstTime = True
'    Context.FileTest = FileText(Get_Nome_Assoluto_File(Test.Parameter(NomeFileIngresso),IndiceSezione,infos.Label1.Caption))
'    Infos.RichTextBox1.Text = ""
'    Infos.RichTextBox1.BackColor = vbWhite
'
'
'    Context.PREPROCESSOR.IndiceDefine = 0
'    Context.PREPROCESSOR.IndiceIfDef = 0
'
'
'    ReDim Context.PREPROCESSOR.DEFINE(0)
'    ReDim Context.PREPROCESSOR.IFDEF(0)
'
'    If (Init_PASSI(Test, Infos, IndiceSezione, Context) = True) Then
'        Infos.RichTextBox1.Text = " ERROR INIT PASSI " + Context.FileTest + " " + err.Description
'        Context.State = eError
'    End If
'
    Exit Sub
    
err:
    Infos.RichTextBox1.Text = " ERROR INIT " + err.Description + " " + Context.FileTest
End Sub 'Init()


'-------
' Start
'-------
Private Sub Start(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC)
    
    Context.State = eRun
    
    SEQUENZA_PLC_INIT_ALU Context.ALU
    SEQUENZA_PLC_INIT_PC Context.ALU, Context.PASSI
    
    
    Context.ALU.IndiceSezione = IndiceSezione
    
    
    
    
    
    ' LOG
    '-----
    'AggiungiLogCh IndiceSezione, "| MODULO: SEQUENZA_PLC | Inizio ---------------------- "

End Sub 'Start()


'------
' Run()
'------
Private Sub Run(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC)
    On Error GoTo err:
    Dim Error As String
    Dim Timer As Long
    
    Timer = GetTickCountEvo + 100
    OldTimer = GetTickCountEvo - 1
    
    Do While (Timer > GetTickCountEvo And Context.State = eRun)
    
        If (OldTimer <= GetTickCountEvo) Then
            
            OldTimer = GetTickCountEvo + 2
            
            PPDO_Manager
    
        If (Context.ALU.PC <= UBound(Context.PASSI)) Then
            
            
            
            Select Case Context.PASSI(Context.ALU.PC).NOME
            
            
            Case "SET"
                Error = ISTR_SET(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            Case "RET"
            
                Error = ISTR_RET(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            Case "CALL"
            
                Error = ISTR_CALL(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "POP"
            
                Error = ISTR_POP(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
    
            
            Case "POPS"
            
                Error = ISTR_POPS(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            Case "PUSHS"
            
                Error = ISTR_PUSHS(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            

            
            Case "PUSH"
            
                Error = ISTR_PUSH(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "ADDS"
            
                Error = ISTR_ADDS(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
                   
            
            
            Case "MOVS"
            
                Error = ISTR_MOVS(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "WRITE"
            
                Error = ISTR_WRITE(Context.PASSI(Context.ALU.PC), Context.ALU, Infos)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            Case "APPEND"
            
                Error = ISTR_APPEND(Context.PASSI(Context.ALU.PC), Context.ALU, Infos)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "PRINT"
            
                Error = ISTR_PRINT(Context.PASSI(Context.ALU.PC), Context.ALU, Infos)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "NOT"
            
                Error = ISTR_NOT(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            
            Case "OR"
            
                Error = ISTR_OR(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
                        
            
            
            Case "AND"
            
                Error = ISTR_AND(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            
            Case "DIV"
            
                Error = ISTR_DIV(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "MUL"
                Error = ISTR_MUL(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            
            Case "ADD":
                Error = ISTR_ADD(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            Case "SUB"
                Error = ISTR_SUB(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            Case "CMP":
            
                Error = ISTR_CMP(Context.PASSI(Context.ALU.PC), Context.ALU)
                
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
            
            Case "MOV":
                Error = ISTR_MOV(Context.PASSI(Context.ALU.PC), Context.ALU)
                
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
                
                
            Case "JMPLE"
                Error = ISTR_JMPC(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                
                If (Error <> "OK") Then
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                    Context.State = eError
                End If
            
                
            
            Case "JMPNE"
                Error = ISTR_JMPC(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                
                If (Error <> "OK") Then
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                    Context.State = eError
                End If
                        
            
            Case "JMPE"
                
                Error = ISTR_JMPC(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
            
                
            
            Case "JMPL"
                
                Error = ISTR_JMPC(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                
                If (Error <> "OK") Then
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                    Context.State = eError
                End If
            
                
            
            Case "JMPG"
                
                Error = ISTR_JMPC(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                
                If (Error <> "OK") Then
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                    Context.State = eError
                End If
            
                
    
            Case "JMPGE"
                
                Error = ISTR_JMPC(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
                        
                
            Case "JMP"
                
                Error = ISTR_JMP(Context.PASSI(Context.ALU.PC), Context.ALU, Context.PASSI)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
                        
            Case "END":
                Context.State = eEnd
            Case "ERROR"
                Context.State = eError
            Case "LABEL":
                Context.ALU.PC = Context.ALU.PC + 1
            
            
            Case "FUNCTION"
                
                Error = ISTR_FUNCTION(Context.PASSI(Context.ALU.PC), Context.ALU)
                If (Error <> "OK") Then
                    Context.State = eError
                    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                End If
                        
                        
            Case Else
                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Context.PASSI(Context.ALU.PC).RIGA + "-> " + Error
                
                Context.State = eError
            End Select
        Else
            Context.State = eError
        End If
        End If
        
    Loop
        
Exit Sub
err:
    Context.State = eError
    Debug.Print err.Description
    
End Sub 'Run()


'------
' Fine
'------
Private Sub Fine(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC)

    If (Infos.RichTextBox1.BackColor = vbRed) Then
        Context.Return = eTEST_ERROR
    Else
        Context.Return = eTEST_OK
    End If
    
End Sub 'Fine()


'------------
' ErrorTest()
'------------
Private Sub ErrorTest(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, ByRef IndiceSezione As Integer, Context As TagSEQUENZA_PLC)
    
    Infos.RichTextBox1.BackColor = vbRed
    
    If (Context.Return <> eTEST_ERROR) Then
        Infos.RichTextBox1 = " ! ! ! ! ! ERRORE SUL TEST ! ! ! ! ! " + vbLf + Infos.RichTextBox1.Text
    End If
    
    Context.Return = eTEST_ERROR

End Sub 'ErrorTest()



Private Function FileText(FileName As String) As String
On Error GoTo err:
    Dim handle As Integer
    handle = FreeFile
    
    Open FileName For Input As #handle
    
    FileText = Input$(LOF(handle), handle)
    
err:
    Close #handle

End Function


Private Function Init_PASSI(ByRef Test As TagSingoloTest, Infos As TagInfosSingleTest, IndiceSezione As Integer, Context As TagSEQUENZA_PLC) As Boolean
    On Error GoTo err:
    Dim Righe As Variant
    Dim i As Long
    Dim Commento As Long
    Dim NomeFinzione As String
    
    Dim NOME As String
    Dim Valore As String
    Dim PASSO As TagIstruzione
    Dim Stato_Create_Define As String
    Dim Stato_Create_Ifdef As String
    
    Context.FileTest = PREPROCESSOR_INCLUDE(Context.FileTest, IndiceSezione)
    
    
    Righe = Split(Context.FileTest, vbCrLf)
    
    
    ' Scansiono il file per creare la lista delle define
    i = 0
    Do While (i <= UBound(Righe))
        
        Stato_Create_Define = PREPROCESSOR_CREATE_IFDEF(Righe(i), Context.PREPROCESSOR)
        
        If (Stato_Create_Define <> "OK") Then
            Infos.RichTextBox1.Text = "ERRORE " + Stato_Create_Define
            Context.State = eError
            Exit Function
        End If
        
        Stato_Create_Define = PREPROCESSOR_CREATE_DEFINE(Righe(i), Context.PREPROCESSOR)
        
        If (Stato_Create_Define <> "OK") Then
            Infos.RichTextBox1.Text = "ERRORE " + Stato_Create_Define
            Context.State = eError
            Exit Function
        End If
        
        Stato_Create_Define = PREPROCESSOR_CREATE_UNDEF(Righe(i), Context.PREPROCESSOR)
        
        If (Stato_Create_Define <> "OK") Then
            Infos.RichTextBox1.Text = "ERRORE " + Stato_Create_Define
            Context.State = eError
            Exit Function
        End If
        
        PREPROCESSOR_SOSTITUISCI_DEFINE Righe(i), Context.PREPROCESSOR
        
        i = i + 1
    Loop
    
    Context.FileTest = Join(Righe, vbCrLf)
    
    Righe = Split(Context.FileTest, vbLf)
    
    
    i = 0
    Do While (i <= UBound(Righe))
            
            ' Applico le define del preprocessor
            'PREPROCESSOR_SOSTITUISCI_DEFINE Righe(i), Context.PREPROCESSOR
            
            ' Decodifico la parola
            If (DECODE_RIGA_TO_PASSO(Righe(i), PASSO, NomeFinzione) = False) Then
                
                If (Context.bFirstTime = True) Then
                
                    ReDim Context.PASSI(0)
                    Context.bFirstTime = False
                    Context.PASSI(0) = PASSO
                    
                Else
                
                    ReDim Preserve Context.PASSI(UBound(Context.PASSI) + 1)
                    Context.PASSI(UBound(Context.PASSI)) = PASSO
                    
                End If
                
            End If
        
        i = i + 1
    Loop
    Exit Function
err:
    Init_PASSI = True
    dubug.Print err.Description
End Function




Private Function PREPROCESSOR_INCLUDE(ByRef ContenutoFile As String, IndiceSezione As Integer) As String

    Dim i As Long
    Dim Righe As Variant
    Dim include As String
    Dim FileTest As String
    Dim TempVariant As Variant
    Dim TempString As String
    Dim ContenutoFileDaIncludere As String
    
    Dim prova As Variant
    Dim NomeFileDaIncludere As String
    Dim ContenutoFileUscita As String
    
    include = "#include "
    
    Righe = Split(ContenutoFile, vbCr)
    
    ContenutoFileUscita = ContenutoFile
    
    Do While (i < UBound(Righe))
        
        If (InStr(1, Righe(i), include) > 0) Then
            TempString = Replace(Righe(i), include, "")
            
            NomeFileDaIncludere = GET_PAROLA_TRA_APICI(TempString)
            NomeFileDaIncludere = Replace(NomeFileDaIncludere, Chr(34), "")
            
            ContenutoFileDaIncludere = FileText(Get_Nome_Assoluto_File(NomeFileDaIncludere, IndiceSezione, Infos.Label1.Caption))
            ContenutoFileDaIncludere = PREPROCESSOR_INCLUDE(ContenutoFileDaIncludere, IndiceSezione)
                        
           TempVariant = Split(ContenutoFileUscita, Righe(i))
            
            ContenutoFileUscita = TempVariant(0) + ContenutoFileDaIncludere + vbCrLf + TempVariant(1)
            
        End If
        
        i = i + 1
    Loop
    
    PREPROCESSOR_INCLUDE = ContenutoFileUscita

End Function






Private Function PREPROCESSOR_CREATE_IFDEF(ByRef RIGA As Variant, PREPROCESSOR As TagPreprocessor) As String

    Dim InizioDefine As Long
    Dim TempString As String
    Dim DEFINE As String
    Dim coppia As Variant
    Dim i As Long
    
    Dim Parola As String
    Dim Valore As String
    
    
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    
    
    DEFINE = "#endif"
    PREPROCESSOR_CREATE_IFDEF = "OK"
    
    RIGA = Rimuovi_Commento_Da_Stringa(RIGA, "//")      ' Tolgo i commenti
    
    InizioDefine = InStr(1, RIGA, DEFINE)
    
    If (InizioDefine > 0) Then
    
        
        TempString = Mid(RIGA, InizioDefine + Len(DEFINE))
                                                                     
        If (PREPROCESSOR.IndiceIfDef = 0) Then
            PREPROCESSOR_CREATE_IFDEF = "ERRORE IFDEF ANNIDATI"
        Else
            PREPROCESSOR.IndiceIfDef = PREPROCESSOR.IndiceIfDef - 1
            
        End If
        RIGA = ""
        Exit Function
    End If
    
    
    If (PREPROCESSOR.IndiceIfDef > 0) Then
        If (PREPROCESSOR.IFDEF(PREPROCESSOR.IndiceIfDef - 1) = False) Then
              RIGA = ""
            Exit Function
        End If
    End If

    
    
    DEFINE = "#ifndef "
    PREPROCESSOR_CREATE_IFDEF = "OK"
    InizioDefine = InStr(1, RIGA, DEFINE)
    
    If (InizioDefine > 0) Then
        TempString = Mid(RIGA, InizioDefine + Len(DEFINE))
        
        coppia = Split(TempString, " ")
        
        If (UBound(coppia) = 0) Then
        
            Parola = Trim(coppia(0))
            
            Parola = Replace(Parola, vbCr, "")
            Parola = Replace(Parola, vbLf, "")
            Parola = Replace(Parola, vbCrLf, "")
            Parola = Replace(Parola, Chr(9), "")
            
            ' Controllo se la define è stata già inserita nel vettore
            i = 0
            Do While (bExit = False)
                If (PREPROCESSOR.DEFINE(i).Parola = Parola) Then
                    bExit = True
                    bFind = False
                ElseIf (i >= UBound(PREPROCESSOR.DEFINE)) Then
                    bExit = True
                    bFind = True
                Else
                    i = i + 1
                End If
            Loop
                        
            RIGA = ""
            
            PREPROCESSOR.IFDEF(PREPROCESSOR.IndiceIfDef) = bFind
            PREPROCESSOR.IndiceIfDef = PREPROCESSOR.IndiceIfDef + 1
                
        End If
    End If
        
        
        
    DEFINE = "#ifdef "
    PREPROCESSOR_CREATE_IFDEF = "OK"
    InizioDefine = InStr(1, RIGA, DEFINE)
    
    If (InizioDefine > 0) Then
        TempString = Mid(RIGA, InizioDefine + Len(DEFINE))
        
        coppia = Split(TempString, " ")
        
        If (UBound(coppia) = 0) Then
        
            Parola = Trim(coppia(0))
            
            Parola = Replace(Parola, vbCr, "")
            Parola = Replace(Parola, vbLf, "")
            Parola = Replace(Parola, vbCrLf, "")
            Parola = Replace(Parola, Chr(9), "")
            
            ' Controllo se la define è stata già inserita nel vettore
            i = 0
            Do While (bExit = False)
                If (PREPROCESSOR.DEFINE(i).Parola = Parola) Then
                    bExit = True
                    bFind = True
                ElseIf (i >= UBound(PREPROCESSOR.DEFINE)) Then
                    bExit = True
                Else
                    i = i + 1
                End If
            Loop
                        
            RIGA = ""
            PREPROCESSOR.IFDEF(PREPROCESSOR.IndiceIfDef) = bFind
            PREPROCESSOR.IndiceIfDef = PREPROCESSOR.IndiceIfDef + 1
                
        End If
    End If
        
        
        
        
End Function


Private Function RemoveItem(ByVal intItem As Integer, intSrc() As TagDefine) As TagDefine()
  Dim intIndex As Integer
  Dim intDest() As TagDefine
  Dim intLBound As Integer, intUBound As Integer
  'find the boundaries of the source array
  intLBound = LBound(intSrc)
  intUBound = UBound(intSrc)
  'set boundaries for the resulting array
  ReDim intDest(intLBound To intUBound - 1) As TagDefine
  'copy items which remain
  For intIndex = intLBound To intItem - 1
    intDest(intIndex) = intSrc(intIndex)
  Next intIndex
  'skip the removed item
  'and copy the remaining items, with destination index-1
  For intIndex = intItem + 1 To intUBound
    intDest(intIndex - 1) = intSrc(intIndex)
  Next intIndex
  'return the result
  RemoveItem = intDest
End Function


' Applico undef
Private Function PREPROCESSOR_CREATE_UNDEF(ByRef RIGA As Variant, PREPROCESSOR As TagPreprocessor) As String
    Dim InizioDefine As Long
    Dim TempString As String
    Dim DEFINE As String
    Dim coppia As Variant
    Dim i As Long
    
    Dim Parola As String
    Dim Valore As String
    
    
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    
    DEFINE = "#undef "
    PREPROCESSOR_CREATE_UNDEF = "OK"
    
    RIGA = Rimuovi_Commento_Da_Stringa(RIGA, "//")      ' Tolgo i commenti
    
    InizioDefine = InStr(1, RIGA, DEFINE)
    
    If (InizioDefine > 0) Then
        TempString = Mid(RIGA, InizioDefine + Len(DEFINE))
        TempString = Replace(TempString, Chr(9), " ")
        TempString = RTrim(TempString)
        TempString = LTrim(TempString)
        
        coppia = Split(TempString, " ")
        
        If (UBound(coppia) = 0) Then
        
            Parola = Trim(coppia(0))
            
            
            ' Controllo se la define è stata già inserita nel vettore
            i = 0
            Do While (bExit = False)
                If (PREPROCESSOR.DEFINE(i).Parola = Parola) Then
                    bExit = True
                    bFind = True
                ElseIf (i >= UBound(PREPROCESSOR.DEFINE)) Then
                    bExit = True
                Else
                    i = i + 1
                End If
            Loop
            
            If (bFind = True) Then
                'Parola = Replace(Trim(Parola), Chr(9), "")
                PREPROCESSOR.DEFINE = RemoveItem(i, PREPROCESSOR.DEFINE)
                                
                PREPROCESSOR.IndiceDefine = UBound(PREPROCESSOR.DEFINE)
                RIGA = ""
            Else
                PREPROCESSOR_CREATE_UNDEF = "UNDEF UNKNOW " + Parola
            End If
            
        End If
    End If
End Function


' Applico le define al sistema
Private Function PREPROCESSOR_CREATE_DEFINE(ByRef RIGA As Variant, PREPROCESSOR As TagPreprocessor) As String
    Dim InizioDefine As Long
    Dim TempString As String
    Dim DEFINE As String
    Dim coppia As Variant
    Dim i As Long
    
    Dim Parola As String
    Dim Valore As String
    
    
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    
    DEFINE = "#define "
    PREPROCESSOR_CREATE_DEFINE = "OK"
    
    RIGA = Rimuovi_Commento_Da_Stringa(RIGA, "//")      ' Tolgo i commenti
    
    InizioDefine = InStr(1, RIGA, DEFINE)
    
    If (InizioDefine > 0) Then
        TempString = Mid(RIGA, InizioDefine + Len(DEFINE))
        TempString = Replace(TempString, Chr(9), " ")
        
        coppia = Split(TempString, " ")
        
        If (UBound(coppia) >= 0) Then
        
            Parola = Trim(coppia(0))
            
            
            ' Controllo se la define è stata già inserita nel vettore
            i = 0
            Do While (bExit = False)
                If (PREPROCESSOR.DEFINE(i).Parola = Parola) Then
                    bExit = True
                    bFind = True
                ElseIf (i >= UBound(PREPROCESSOR.DEFINE)) Then
                    bExit = True
                Else
                    i = i + 1
                End If
            Loop
            
            If (bFind = False) Then
                Parola = Replace(Trim(Parola), Chr(9), "")
                PREPROCESSOR.DEFINE(PREPROCESSOR.IndiceDefine).Parola = Parola
                TempString = Replace(TempString, PREPROCESSOR.DEFINE(PREPROCESSOR.IndiceDefine).Parola, "")
                PREPROCESSOR.DEFINE(PREPROCESSOR.IndiceDefine).Valore = Replace(Trim(TempString), vbCr, "")
                PREPROCESSOR.DEFINE(PREPROCESSOR.IndiceDefine).Valore = Replace(Trim(TempString), Chr(9), "")
            
                PREPROCESSOR.IndiceDefine = PREPROCESSOR.IndiceDefine + 1
            
                ReDim Preserve PREPROCESSOR.DEFINE(PREPROCESSOR.IndiceDefine)
                RIGA = ""
            Else
                PREPROCESSOR_CREATE_DEFINE = "DEFINE GIA' DEFINITA " + Parola
            End If
            
        End If
    End If
End Function


Private Function PREPROCESSOR_SOSTITUISCI_DEFINE(ByRef RIGA As Variant, PREPROCESSOR As TagPreprocessor)
    
    Dim i As Long

    
    If (PREPROCESSOR_DEFINE = False) Then
        i = 0
        Do While (i < PREPROCESSOR.IndiceDefine)
            If (InStr(1, RIGA, "#") > 0) Then
            
            Else
                RIGA = Replace(RIGA, PREPROCESSOR.DEFINE(i).Parola, PREPROCESSOR.DEFINE(i).Valore)
            End If
            i = i + 1
        Loop
        
    End If
    
End Function

Private Function GET_PAROLA_TRA_APICI(ByRef INGRESSO As Variant) As String
    Dim PrimoApice As Long
    Dim SecondoApice As Long

    PrimoApice = InStr(1, INGRESSO, Chr(34))
    
    If (PrimoApice > 0) Then
        SecondoApice = InStr(PrimoApice + 1, INGRESSO, Chr(34))
                
        If (SecondoApice > 0) Then
                    
            GET_PAROLA_TRA_APICI = Mid(INGRESSO, PrimoApice, SecondoApice - PrimoApice + 1)
                
        End If
                
    Else
        GET_PAROLA_TRA_APICI = ""
    End If
    
End Function


' Questa funzione analizza una riga e crea il corrispettivo passo
Private Function DECODE_RIGA_TO_PASSO(ByVal RIGA As String, ByRef PASSO_OUT As TagIstruzione, ByRef NomeFinzione As String) As Boolean
    
    Dim strIndex As Long
    Dim bReturn As Boolean
    Dim Debug_riga As String
    
    Dim PrimoApice As Long
    Dim SecondoApice As Long
    
    
    Dim PASSO As TagIstruzione
    Dim PARAMETRI As Variant
    Dim i As Long
    Dim bFind As Boolean
    Debug_riga = RIGA
    
    RIGA = Rimuovi_Commento_Da_Stringa(RIGA, "//")      ' Tolgo i commenti
    RIGA = Replace(RIGA, vbLf, "")
    RIGA = Replace(RIGA, vbTab, "")
    RIGA = Replace(RIGA, vbCr, "")
    RIGA = Replace(RIGA, vbCrLf, "")
    
    
    If (InStr(1, RIGA, "$$FUNCTION$$") > 0) Then
        Debug.Print NomeFinzione
    End If
        
    If (InStr(RIGA, "_ELSEA") > 0) Then
        Debug.Print "SUKA"
    End If
    RIGA = Replace(RIGA, "$$FUNCTION$$", NomeFinzione)
    RIGA = LTrim(RIGA)
    RIGA = RTrim(RIGA)
    strIndex = InStr(1, RIGA, " ")
    ' Cerco il primo spazio nella riga con valore diverso da 0
    If (strIndex > 0) Then
        
        ' Mi salvo il nome ...
        PASSO.NOME = Left(RIGA, strIndex - 1)
        PASSO.RIGA = Debug_riga
        

        
        
        RIGA = Mid(RIGA, strIndex + 1)

        
        PARAMETRI = Split(RIGA, ",")
                
        i = 0
        Do While (i <= UBound(PARAMETRI))
            PARAMETRI(i) = Replace(PARAMETRI(i), vbCr, "")
            
            PrimoApice = InStr(1, PARAMETRI(i), Chr(34))
            
            If (PrimoApice > 0) Then
                SecondoApice = InStr(PrimoApice + 1, PARAMETRI(i), Chr(34))
                
                If (SecondoApice > 0) Then
                    
                    PARAMETRI(i) = Mid(PARAMETRI(i), PrimoApice, SecondoApice - PrimoApice + 1)
                
                End If
                
            Else
                PARAMETRI(i) = Replace(PARAMETRI(i), " ", "")
            End If
            i = i + 1
        Loop
        
        If (PASSO.NOME = "FUNCTION") Then
            NomeFinzione = Replace(PARAMETRI(0), Chr(34), "")
        End If
        
        ReDim PASSO.par(UBound(PARAMETRI))
        
        ' COPIO I PARAMETRI IN USCITA
        i = 0
        Do While (i <= UBound(PARAMETRI))
            PASSO.par(i) = Replace(PARAMETRI(i), ":", "")
            i = i + 1
        Loop
        
        bReturn = False
        bFind = True
    
    End If
    
    If (bFind = False) Then
        strIndex = InStr(1, RIGA, ":")
                
        ' Controllo se è una label
        If (strIndex > 0) Then
            
            ' Controllo se la label è un salto relativo all'interno della funzoine
                                    
            
            PASSO.NOME = "LABEL"
            ReDim PASSO.par(0)
            PASSO.par(0) = Left(RIGA, strIndex - 1)
            
            PASSO.par(0) = Trim(PASSO.par(0))               ' Rimuovo gli spazi
            
            bReturn = False
            bFind = True
        End If
    End If
    
    
    If (bFind = False) Then
        
        If (InStr(1, RIGA, "END") > 0) Then
            PASSO.NOME = "END"
            bFind = True
        End If
    End If
    
    If (bFind = False) Then
        
        If (InStr(1, RIGA, "ERROR") > 0) Then
            PASSO.NOME = "ERROR"
            bFind = True
        End If
    End If
    
    If (bFind = False) Then
        
        If (InStr(1, RIGA, "RET") > 0) Then
            PASSO.NOME = "RET"
            bFind = True
        End If
    End If
    

    If (bFind = True) Then
        PASSO_OUT = PASSO
        bReturn = False
    Else
        bReturn = True
    End If
    
    
    DECODE_RIGA_TO_PASSO = bReturn
End Function



Private Function ISTR_JMPC(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu, ByRef PASSI() As TagIstruzione) As String
    Dim i As Long
    Dim bFind As Boolean
    Dim bExit As Boolean
    Dim bJump As Boolean
    
    ISTR_JMPC = "OK"
    
    Select Case PASSO.NOME
    
    Case "JMPLE":           ' minore uguale
        If (ALU.n = True Or ALU.Z = True) Then
            bJump = True
        End If
    Case "JMPE"             ' uguale
        If (ALU.Z = True) Then
            bJump = True
        End If
    Case "JMPNE"             ' non uguale uguale
        If (ALU.Z = False) Then
            bJump = True
        End If
    Case "JMPL"             ' minore
        If (ALU.n = True And ALU.Z = False) Then
            bJump = True
        End If
    Case "JMPG"             ' maggiore
        If (ALU.n = False And ALU.Z = False) Then
            bJump = True
        End If
    Case "JMPGE"            ' maggiore uguale
        If (ALU.n = False Or ALU.Z = True) Then
            bJump = True
        End If
    Case Else
        ISTR_JMPC = "ERRORE ISTRUZIONE NON CONOSCIUTA"
    End Select
    
    
    If (ISTR_JMPC = "OK") Then
        
        If (bJump = True) Then
            If (UBound(PASSO.par) = 0) Then
            
                Do While (bExit = False)
                    
                    If (i > UBound(PASSI)) Then
                        
                        bExit = True
                    ElseIf (IsArrayInitialized(PASSI(i).par) = False) Then
                        i = i + 1
                    ElseIf (PASSI(i).NOME = "LABEL" And PASSI(i).par(0) = PASSO.par(0)) Then
                        
                        bFind = True
                        bExit = True
                        
                        ALU.PC = i
                        ISTR_JMPC = "OK"
                    Else
                        i = i + 1
                    End If
                
                Loop
                
                
                If (bFind = False) Then
                    ISTR_JMPC = "LABEL NON TROVATA"
                End If
            Else
                ISTR_JMPC = "PARAMETRO NON VALIDO"
            End If
        Else
            ALU.PC = ALU.PC + 1
        End If
    End If
    
End Function





Private Function ISTR_FUNCTION(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
On Error GoTo err:
    ALU.CurrentFunction = PASSO.par(0)
    ALU.PC = ALU.PC + 1
    
    ISTR_FUNCTION = "OK"
    Exit Function
err:
    ISTR_FUNCTION = "ERRORE GENERICO IN " + PASSO.RIGA + " " + err.Description
End Function

' ISTRUZIONI SUPPORTATE !!!!!!!!!!
Private Function ISTR_JMP(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu, ByRef PASSI() As TagIstruzione) As String
    Dim i As Long
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    If (PASSO.NOME = "JMP") Then
        
        If (UBound(PASSO.par) = 0) Then
        
            Do While (bExit = False)
                
                If (i > UBound(PASSI)) Then
                    
                    bExit = True
                    
                ElseIf (IsArrayInitialized(PASSI(i).par) = False) Then
                    i = i + 1
                    
                ElseIf (PASSI(i).NOME = "LABEL" And PASSI(i).par(0) = PASSO.par(0)) Then
                    
                    bFind = True
                    bExit = True
                    
                    ALU.PC = i
                    ISTR_JMP = "OK"
                Else
                    i = i + 1
                End If
            
            Loop
            
            If (bFind = False) Then
                ISTR_JMP = "NON HO TROVATO IL JUMP"
            End If
            
        Else
            ISTR_JMP = "PARAMETRO NON VALIDO"
        End If
    Else
        ISTR_JMP = "ERROR NOME ISTRUZIONE ERRATO"
    End If
    
End Function





' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_AND(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "AND") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_AND = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_AND = "OK") Then
                
                ISTR_AND = GET_PARAMETRO_VALUE(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_AND = "OK") Then
                
                    Resultato = Value1 And Value2
                    
                    
                    ISTR_AND = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_AND = "OK") Then
                
                        UpdateStatusRegister Resultato, ALU
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_AND = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_AND = "ERRORE " + err.Description
End Function





Private Function ISTR_MOVS(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As String

    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "MOVS") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 1) Then
            
            ISTR_MOVS = GET_PARAMETRO_STRINGA(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_MOVS = "OK") Then
                
                ISTR_MOVS = SET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Value1)
                ALU.PC = ALU.PC + 1
                Exit Function
                                
            End If
            
            
        Else
            ISTR_MOVS = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_MOVS = "ERRORE " + err.Description
End Function








Private Function ISTR_WRITE(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu, Infos As TagInfosSingleTest) As String
    On Error GoTo err:
    
    Dim Value1 As String

    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "WRITE") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
            
            ISTR_WRITE = GET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Value1)
            
            
            If (ISTR_WRITE = "OK") Then
                
                
                Infos.RichTextBox1.Text = Value1 + vbCr
                ALU.PC = ALU.PC + 1
                Exit Function
                                
            End If
            
            
        Else
            ISTR_WRITE = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_WRITE = "ERRORE " + err.Description
End Function




Private Function ISTR_APPEND(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu, Infos As TagInfosSingleTest) As String
    On Error GoTo err:
    
    Dim Value1 As String

    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "APPEND") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
            
            ISTR_APPEND = GET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Value1)
            
            
            If (ISTR_APPEND = "OK") Then
                
                
                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Value1
                Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
                ALU.PC = ALU.PC + 1
                Exit Function
                                
            End If
            
            
        Else
            ISTR_APPEND = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function
err:

ISTR_APPEND = "ERRORE " + err.Description
End Function




Private Function ISTR_PRINT(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu, Infos As TagInfosSingleTest) As String
    On Error GoTo err:
    
    Dim Value1 As String

    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "PRINT") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
            
            ISTR_PRINT = GET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Value1)
            
            
            If (ISTR_PRINT = "OK") Then
                
                
                Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + Value1 + vbCr
                Infos.RichTextBox1.SelStart = Len(Infos.RichTextBox1.Text)
                ALU.PC = ALU.PC + 1
                Exit Function
                                
            End If
            
            
        Else
            ISTR_PRINT = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_PRINT = "ERRORE " + err.Description
End Function





' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_NOT(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "NOT") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 1) Then
            
            ISTR_NOT = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_NOT = "OK") Then
                
                Resultato = Not Value1
                
                ISTR_NOT = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                
                If (ISTR_NOT = "OK") Then
                    UpdateStatusRegister Resultato, ALU
                    ALU.PC = ALU.PC + 1
                    Exit Function
                End If
                
            End If
            
            
        Else
            ISTR_NOT = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_NOT = "ERRORE " + err.Description
End Function




' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_OR(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "OR") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_OR = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_OR = "OK") Then
                
                ISTR_OR = GET_PARAMETRO_VALUE(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_OR = "OK") Then
                
                    Resultato = Value1 Or Value2
                    
                    
                    ISTR_OR = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_OR = "OK") Then
                
                        UpdateStatusRegister Resultato, ALU
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_OR = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_OR = "ERRORE " + err.Description
End Function











' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_DIV(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "DIV") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_DIV = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_DIV = "OK") Then
                
                ISTR_DIV = GET_PARAMETRO_VALUE(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_DIV = "OK") Then
                
                    Resultato = Value1 / Value2
                    
                    
                    ISTR_DIV = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_DIV = "OK") Then
                
                        UpdateStatusRegister Resultato, ALU
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_DIV = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_DIV = "ERRORE " + err.Description
End Function






' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_MUL(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "MUL") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_MUL = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_MUL = "OK") Then
                
                ISTR_MUL = GET_PARAMETRO_VALUE(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_MUL = "OK") Then
                
                    Resultato = Value1 * Value2
                    
                    
                    ISTR_MUL = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_MUL = "OK") Then
                
                        UpdateStatusRegister Resultato, ALU
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_MUL = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_MUL = "ERRORE " + err.Description
End Function



' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_ADD(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "ADD") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_ADD = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_ADD = "OK") Then
                
                ISTR_ADD = GET_PARAMETRO_VALUE(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_ADD = "OK") Then
                
                    Resultato = Value1 + Value2
                    
                    
                    ISTR_ADD = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_ADD = "OK") Then
                
                        UpdateStatusRegister Resultato, ALU
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_ADD = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_ADD = "ERRORE " + err.Description
End Function








Private Function ISTR_ADDS(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As String
    Dim Value2 As String
    Dim Resultato As String
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "ADDS") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_ADDS = GET_PARAMETRO_STRINGA(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_ADDS = "OK") Then
                
                ISTR_ADDS = GET_PARAMETRO_STRINGA(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_ADDS = "OK") Then
                
                    Resultato = Value1 + Value2
                    
                    
                    ISTR_ADDS = SET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_ADDS = "OK") Then
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_ADDS = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_ADDS = "ERRORE " + err.Description
End Function


' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_SUB(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim Resultato As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "SUB") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 2) Then
            
            ISTR_SUB = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value1)
            
            
            If (ISTR_SUB = "OK") Then
                
                ISTR_SUB = GET_PARAMETRO_VALUE(ALU, PASSO.par(2), Value2)
                
                
                If (ISTR_SUB = "OK") Then
                
                    Resultato = Value1 - Value2
                    
                    
                    ISTR_SUB = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Resultato)
                
                    If (ISTR_SUB = "OK") Then
                
                        UpdateStatusRegister Resultato, ALU
                        ALU.PC = ALU.PC + 1
                        Exit Function
                    
                    End If
                    
                End If
                
            End If
            
            
        Else
            ISTR_SUB = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_SUB = "ERRORE " + err.Description
End Function


' Questa funzione effettua un compare tra il valore dell'accumulatore ed il parametro
Private Function ISTR_CMP(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value1 As Double
    Dim Value2 As Double
    Dim compare As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "CMP") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 1) Then
            
            ISTR_CMP = GET_PARAMETRO_VALUE(ALU, PASSO.par(0), Value1)
            
            If (ISTR_CMP = "OK") Then
                
                ISTR_CMP = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value2)
                
                If (ISTR_CMP = "OK") Then
                
                    compare = Value1 - Value2
                
                    UpdateStatusRegister compare, ALU
                    ALU.PC = ALU.PC + 1
                    Exit Function
                    
                End If
                
            End If
            
            

            
        Else
            ISTR_CMP = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_CMP = "ERRORE " + err.Description
End Function




Private Function ISTR_PUSHS(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    
    On Error GoTo err:
    
    Dim Value As String
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "PUSHS") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
                        
            ISTR_PUSHS = GET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Value)
            
            ALU.STACK_STRINGHE(ALU.SP_STRINGHE) = Value
            
            ALU.SP_STRINGHE = ALU.SP_STRINGHE + 1
            
            If (ALU.SP_STRINGHE > UBound(ALU.STACK_STRINGHE)) Then
                ISTR_PUSHS = "STACK OVERFLOW"
            Else
                ALU.PC = ALU.PC + 1
                Exit Function
            End If
                        
        Else
            ISTR_PUSHS = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_PUSHS = "ERRORE " + err.Description

End Function



Private Function ISTR_PUSH(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    
    On Error GoTo err:
    
    Dim Value As Double
    Dim Pointer As String
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "PUSH") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
            
            ISTR_PUSH = IS_PARAMETRO_POINTER(PASSO.par(0))
            
            If (ISTR_PUSH = "OK") Then
                ISTR_PUSH = GET_PARAMETRO_ADDRESS(ALU, PASSO.par(0), Pointer)
                If (ISTR_PUSH = "OK") Then
                    ALU.STACK_POINTER(ALU.SP_POINTER) = Pointer
                                
                    ALU.SP_POINTER = ALU.SP_POINTER + 1
                    ALU.PC = ALU.PC + 1
                    Exit Function
                End If
            End If
            
            
            ISTR_PUSH = GET_PARAMETRO_VALUE(ALU, PASSO.par(0), Value)
            If (ISTR_PUSH = "OK") Then
                ALU.STACK(ALU.SP) = Value
                ALU.SP = ALU.SP + 1
            
                If (ALU.SP > UBound(ALU.STACK)) Then
                    ISTR_PUSH = "STACK OVERFLOW"
                Else
                    ALU.PC = ALU.PC + 1
                    Exit Function
                End If
            Else
                ISTR_PUSH = PASSO.RIGA + " " + ISTR_PUSH
            End If
                        
        Else
            ISTR_PUSH = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_PUSH = "ERRORE " + err.Description

End Function




Private Function ISTR_RET(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    
    If (PASSO.NOME = "RET") Then
        
        POP_STACK_CALL ALU
        ALU.PC = ALU.LR
        ISTR_RET = "OK"

    Else
        ISTR_RET = "ERROR NOME ISTRUZIONE ERRATO"
    End If
End Function


Private Sub PUSH_STACK_CALL(ByRef ALU As TagAlu)
Dim i As Long

    i = 0
    Do While (i < UBound(ALU.R))
        ALU.STACK_CALL(ALU.SP_CALL).R(i) = ALU.R(i)
        i = i + 1
    Loop
    
    i = 0
    Do While (i < UBound(ALU.s))
        ALU.STACK_CALL(ALU.SP_CALL).s(i) = ALU.s(i)
        i = i + 1
    Loop
    
    
    i = 0
    Do While (i < UBound(ALU.p))
        ALU.STACK_CALL(ALU.SP_CALL).p(i) = ALU.p(i)
        i = i + 1
    Loop
    
    
    ALU.STACK_CALL(ALU.SP_CALL).LR = ALU.LR
    
    ALU.SP_CALL = ALU.SP_CALL + 1
    
    
End Sub


Private Sub POP_STACK_CALL(ByRef ALU As TagAlu)
Dim i As Long

    ALU.SP_CALL = ALU.SP_CALL - 1
    i = 0
    Do While (i < UBound(ALU.R))
        ALU.R(i) = ALU.STACK_CALL(ALU.SP_CALL).R(i)
        i = i + 1
    Loop
    
    i = 0
    Do While (i < UBound(ALU.s))
        ALU.s(i) = ALU.STACK_CALL(ALU.SP_CALL).s(i)
        i = i + 1
    Loop
    
    i = 0
    Do While (i < UBound(ALU.p))
        ALU.p(i) = ALU.STACK_CALL(ALU.SP_CALL).p(i)
        i = i + 1
    Loop
    
    
    ALU.LR = ALU.STACK_CALL(ALU.SP_CALL).LR
    
End Sub


Private Function ISTR_CALL(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu, ByRef PASSI() As TagIstruzione) As String
On Error GoTo err:
    Dim i As Long
    Dim bFind As Boolean
    Dim bExit As Boolean
    
    If (PASSO.NOME = "CALL") Then
        
        If (UBound(PASSO.par) = 0) Then
        
            Do While (bExit = False)
                
                If (i > UBound(PASSI)) Then
                    
                    bExit = True
                    
                ElseIf (IsArrayInitialized(PASSI(i).par) = False) Then
                    i = i + 1
                ElseIf (PASSI(i).NOME = "FUNCTION" And PASSI(i).par(0) = PASSO.par(0)) Then
                        
                        bFind = True
                        bExit = True
                        
                        
                        
                        
                        
                        ALU.LR = ALU.PC + 1
                        ALU.PC = i
                        
                        PUSH_STACK_CALL ALU
                        
                        ISTR_CALL = "OK"
                    
                Else
                    i = i + 1
                End If
            
            Loop
        Else
            ISTR_CALL = "PARAMETRO NON VALIDO"
        End If
    Else
        ISTR_CALL = "ERROR NOME ISTRUZIONE ERRATO"
    End If
    
    If (bFind = False) Then
        ISTR_CALL = "FUNZIONE NON TROVATA"
    End If
Exit Function
err:
    ISTR_CALL = "ERRORE " + err.Description
End Function





Private Function ISTR_POPS(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value As String
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "POPS") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
            
            ALU.SP_STRINGHE = ALU.SP_STRINGHE - 1
            
            If (ALU.SP_STRINGHE < 0) Then
                ISTR_POPS = "STACK OVERFLOW"
                Exit Function
            End If
            
            Value = ALU.STACK_STRINGHE(ALU.SP_STRINGHE)
            
            ISTR_POPS = SET_PARAMETRO_STRINGA(ALU, PASSO.par(0), Value)
            ALU.PC = ALU.PC + 1
            Exit Function
        Else
            ISTR_POPS = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_POPS = "ERRORE " + err.Description

End Function



Private Function ISTR_POP(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    
    On Error GoTo err:
    
    Dim Value As Double
    Dim Pointer As String
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "POP") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 0) Then
            
            
            
            ISTR_POP = IS_PARAMETRO_POINTER(PASSO.par(0))
            
            If (ISTR_POP = "OK") Then
                ALU.SP_POINTER = ALU.SP_POINTER - 1
                ISTR_POP = SET_PARAMETRO_ADDRESS(ALU, PASSO.par(0), ALU.STACK_POINTER(ALU.SP_POINTER))
                ALU.PC = ALU.PC + 1
                Exit Function
            End If
            
            ALU.SP = ALU.SP - 1
            
            If (ALU.SP < 0) Then
                ISTR_POP = "STACK OVERFLOW"
                Exit Function
            End If
            
            Value = ALU.STACK(ALU.SP)
            
            ISTR_POP = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Value)
            ALU.PC = ALU.PC + 1
            Exit Function
        Else
            ISTR_POP = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_POP = "ERRORE " + err.Description

End Function



Private Function ISTR_SET(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    On Error GoTo err:
    
    Dim Value As String
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "SET") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 1) Then
            
            ISTR_SET = GET_PARAMETRO_ADDRESS(ALU, PASSO.par(1), Value)
            
            If (ISTR_SET = "OK") Then
            
                ISTR_SET = SET_PARAMETRO_ADDRESS(ALU, PASSO.par(0), Value)
                
                If (ISTR_SET = "OK") Then
                
                    'UpdateStatusRegister ALU.A, ALU
                    ALU.PC = ALU.PC + 1
                    Exit Function
                    
                End If
                
                Exit Function
            End If
            
        Else
            ISTR_SET = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
    Exit Function
err:
    
    ISTR_SET = "ERRORE " + err.Description
End Function


Private Function ISTR_MOV(ByRef PASSO As TagIstruzione, ByRef ALU As TagAlu) As String
    
    On Error GoTo err:
    
    Dim Value As Double
    
    ' Controllo il nome dell'istruzione
    
    bReturn = "ERRORE"
    
    If (PASSO.NOME = "MOV") Then
        
        
        ' Controllo il numero di operandi
        If (UBound(PASSO.par) = 1) Then
            
            ISTR_MOV = GET_PARAMETRO_VALUE(ALU, PASSO.par(1), Value)
            
            If (ISTR_MOV = "OK") Then
            
                ISTR_MOV = SET_PARAMETRO_VALUE(ALU, PASSO.par(0), Value)
                
                If (ISTR_MOV = "OK") Then
                
                    'UpdateStatusRegister ALU.A, ALU
                    ALU.PC = ALU.PC + 1
                    Exit Function
                    
                End If
                
                Exit Function
            End If
            
        Else
            ISTR_MOV = "ERROR NOME ISTRUZIONE ERRATO"
            
        End If
    End If
Exit Function

err:

ISTR_MOV = "ERRORE " + err.Description

End Function




Private Function SET_PARAMETRO_STRINGA(ByRef ALU As TagAlu, ByRef PARAMETRO As String, ByRef Value As String) As String

    Dim REGISTER As Integer

    
    ' E' un registro
    SET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_INTERNALSTRING(PARAMETRO, REGISTER)
    
    If (SET_PARAMETRO_STRINGA = "OK") Then
    
        ALU.s(REGISTER) = Value
        Exit Function
        
    End If
    

End Function



Private Function READ_POINTER_DATA(ByRef ALU As TagAlu, ByRef POINTER_INDEX As Integer, ByRef Value As Double) As String

    READ_POINTER_DATA = GET_PARAMETRO_VALUE(ALU, ALU.p(POINTER_INDEX), Value)

End Function



Private Function WRITE_POINTER_DATA(ByRef ALU As TagAlu, ByRef POINTER_INDEX As Integer, ByVal Value As Double) As String

    WRITE_POINTER_DATA = SET_PARAMETRO_VALUE(ALU, ALU.p(POINTER_INDEX), Value)

End Function




Private Function SET_PARAMETRO_ADDRESS(ByRef ALU As TagAlu, ByRef PARAMETRO As String, ByRef Value As String) As String
    Dim TP As TagTupla
    Dim REGISTER As Integer
    Dim Pointer As Integer
        
    
    ' E' un puntatore
    SET_PARAMETRO_ADDRESS = CONVERT_PARAMETRO_TO_POINTER(PARAMETRO, Pointer)
    
    If (SET_PARAMETRO_ADDRESS = "OK") Then
        
        ALU.p(Pointer) = Value
    Else
        SET_PARAMETRO_ADDRESS = "ERRORE L'INDIRIZZO DI UNA VARIABILE O REGISTRO PUO' ESSERE IMPOSTATO SOLO SU UN PUNTATORE"
    End If
      
    
End Function



Private Function SET_PARAMETRO_VALUE(ByRef ALU As TagAlu, ByRef PARAMETRO As String, ByRef Value As Double) As String
    Dim TP As TagTupla
    Dim REGISTER As Integer
    Dim Pointer As Integer

    
    ' E' una Tupla
    SET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_TUPLA(PARAMETRO, TP)
    
    If (SET_PARAMETRO_VALUE = "OK") Then
        
        SetCalibratedValue_INTERNAL ALU, TP.ID, TP.TYPE, TP.Addr, Value
        Exit Function
    
    End If
    
    
    ' E' un registro
    SET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_REGISTER(PARAMETRO, REGISTER)
    
    If (SET_PARAMETRO_VALUE = "OK") Then
    
        ALU.R(REGISTER) = Value
        Exit Function
        
    End If
    
    
    ' E' un puntatore
    SET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_POINTER(PARAMETRO, Pointer)
    
    If (SET_PARAMETRO_VALUE = "OK") Then
        
        If (WRITE_POINTER_DATA(ALU, Pointer, Value) = "OK") Then
            Exit Function
        End If
    End If
      
    
End Function

    
    
Private Function GET_PARAMETRO_STRINGA(ByRef ALU As TagAlu, ByRef PARAMETRO As String, ByRef Value As String) As String
On Error GoTo err:
    Dim STRINGA As String
    Dim TP As TagTupla
    Dim NUMBER As Double
    Dim Pointer As Integer
    
    Dim INDEX_STRING As Integer
        ' E' una stringa
    GET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_STRINGA(PARAMETRO, STRINGA)
    
    If (GET_PARAMETRO_STRINGA = "OK") Then
    
        Value = STRINGA
        Exit Function
        
    End If


    ' E' una stringa interna
    GET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_INTERNALSTRING(PARAMETRO, INDEX_STRING)
    
    If (GET_PARAMETRO_STRINGA = "OK") Then
    
        Value = ALU.s(CInt(INDEX_STRING))
        Exit Function
        
    End If


    ' E' una tupla
    GET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_TUPLA(PARAMETRO, TP)
    
    If (GET_PARAMETRO_STRINGA = "OK") Then
        
        Value = CStr(GetCalibratedValue_INTERNAL(ALU, TP.ID, TP.TYPE, TP.Addr))
        Exit Function
        
    End If



    ' E' una register
    GET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_REGISTER(PARAMETRO, INDEX_STRING)
    
    If (GET_PARAMETRO_STRINGA = "OK") Then
        
        Value = CStr(ALU.R(CInt(INDEX_STRING)))
        Exit Function
        
    End If


    ' E' un numero
    GET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_NUMBER(PARAMETRO, NUMBER)
    
    If (GET_PARAMETRO_STRINGA = "OK") Then
        
        Value = CStr(NUMBER)
        
        Exit Function
        
    End If
    
    ' E' una pointer
    GET_PARAMETRO_STRINGA = CONVERT_PARAMETRO_TO_POINTER(PARAMETRO, Pointer)
    
    If (GET_PARAMETRO_STRINGA = "OK") Then
        GET_PARAMETRO_STRINGA = GET_PARAMETRO_STRINGA(ALU, ALU.p(Pointer), Value)
        
        Exit Function
        
    End If

err:
    GET_PARAMETRO_STRINGA = "ERRORE " + err.Description
End Function
    
    


Private Function GET_PARAMETRO_ADDRESS(ByRef ALU As TagAlu, ByRef PARAMETRO As String, ByRef Value As String) As String
On Error GoTo err:
    
    Dim TP As TagTupla
    Dim REGISTER As Integer
    Dim NUMBER As Double
    Dim Pointer As Integer
    
    ' E' una Tupla
    GET_PARAMETRO_ADDRESS = CONVERT_PARAMETRO_TO_TUPLA(PARAMETRO, TP)
    
    If (GET_PARAMETRO_ADDRESS = "OK") Then
        Value = PARAMETRO
        Exit Function
    
    End If
    
    
    ' E' un registro
    GET_PARAMETRO_ADDRESS = CONVERT_PARAMETRO_TO_REGISTER(PARAMETRO, REGISTER)
    
    If (GET_PARAMETRO_ADDRESS = "OK") Then
        Value = PARAMETRO
        Exit Function
        
    End If
    
    
    ' E' un numero
    GET_PARAMETRO_ADDRESS = CONVERT_PARAMETRO_TO_NUMBER(PARAMETRO, NUMBER)
    
    If (GET_PARAMETRO_ADDRESS = "OK") Then
    
        GET_PARAMETRO_ADDRESS = "ERRORE UN NUMERO NON PUO' ESSERE UTILIZZATO COME INDIRIZZO"
        Exit Function
        
    End If


    ' E' un puntatore
    GET_PARAMETRO_ADDRESS = CONVERT_PARAMETRO_TO_POINTER(PARAMETRO, Pointer)
    
    If (GET_PARAMETRO_ADDRESS = "OK") Then
            GET_PARAMETRO_ADDRESS = GET_PARAMETRO_ADDRESS(ALU, ALU.p(Pointer), Value)
            Exit Function
    End If
    
    Exit Function
    
err:

    GET_PARAMETRO_ADDRESS = "ERRORE GENERICO " + err.Description
    
End Function

Private Function IS_PARAMETRO_POINTER(ByRef PARAMETRO As String) As String
    Dim Pointer As Integer
    ' E' un puntatore
    IS_PARAMETRO_POINTER = CONVERT_PARAMETRO_TO_POINTER(PARAMETRO, Pointer)
    
End Function


Private Function GET_PARAMETRO_VALUE(ByRef ALU As TagAlu, ByRef PARAMETRO As String, ByRef Value As Double) As String
On Error GoTo err:
    
    Dim TP As TagTupla
    Dim REGISTER As Integer
    Dim Pointer As Integer
    Dim NUMBER As Double
    
    
    ' E' una Tupla
    GET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_TUPLA(PARAMETRO, TP)
    
    If (GET_PARAMETRO_VALUE = "OK") Then
        
        Value = GetCalibratedValue_INTERNAL(ALU, TP.ID, TP.TYPE, TP.Addr)
        Exit Function
    
    End If
    
    
    ' E' un registro
    GET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_REGISTER(PARAMETRO, REGISTER)
    
    If (GET_PARAMETRO_VALUE = "OK") Then
    
        Value = ALU.R(REGISTER)
        Exit Function
        
    End If
    
    
    ' E' un numero
    GET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_NUMBER(PARAMETRO, NUMBER)
    
    If (GET_PARAMETRO_VALUE = "OK") Then
    
        Value = NUMBER
        Exit Function
        
    End If


    ' E' un puntatore
    GET_PARAMETRO_VALUE = CONVERT_PARAMETRO_TO_POINTER(PARAMETRO, Pointer)
    
    If (GET_PARAMETRO_VALUE = "OK") Then
    
        If (READ_POINTER_DATA(ALU, Pointer, Value) = "OK") Then
            Exit Function
        End If
    End If

    
    
err:

    GET_PARAMETRO_VALUE = "ERRORE GENERICO DEL PARAMETRO " + CStr(PARAMETRO) + " IN " + CStr(ALU.PC) + err.Description
    
End Function




Private Function CONVERT_PARAMETRO_TO_NUMBER(ByRef PARAMETRO As String, ByRef NUMBER As Double) As String
    On Error GoTo err:
    Dim bReturn As String
    Dim TempString As String
    
    bReturn = "OK"
    TempString = PARAMETRO
    
    If (InStr(1, TempString, "@") > 0) Then
        TempString = Replace(TempString, "@", "")
        If (IsNumeric(TempString) = True) Then
            NUMBER = CDbl(Val(TempString))
        Else
            bReturn = "ERRORE IL PARAMETRO NON E' UN NUMERO"
        End If
        
    Else
        bReturn = "ERRORE IL PARAMETRO NON E' UN NUMERO"
    End If
    
    CONVERT_PARAMETRO_TO_NUMBER = bReturn
    
    Exit Function
err:
    CONVERT_PARAMETRO_TO_NUMBER = "ERROR " + err.Description

End Function



Private Function CONVERT_PARAMETRO_TO_INTERNALSTRING(ByRef PARAMETRO As String, ByRef INTERNAL_INDEX As Integer) As String
    On Error GoTo err:
    Dim bReturn As String
    Dim TempString As String
    
    bReturn = "OK"
    TempString = PARAMETRO
    
    If (InStr(1, TempString, "S") > 0) Then
        TempString = Replace(TempString, "S", "")
        INTERNAL_INDEX = CInt(TempString)
        
    Else
        bReturn = "ERRORE IL PARAMETRO NON E' UN REGISTRO"
    End If
    
    CONVERT_PARAMETRO_TO_INTERNALSTRING = bReturn
    
    Exit Function
err:
    CONVERT_PARAMETRO_TO_INTERNALSTRING = "ERROR " + err.Description
End Function



Private Function CONVERT_PARAMETRO_TO_STRINGA(ByRef PARAMETRO As String, ByRef STRINGA As String) As String
    On Error GoTo err:
    Dim bReturn As String
    Dim TempString As String
    Dim linefeed As Variant
    Dim i As Long
    Dim bFirstTime As Boolean
    Dim NewString As String
    
    bReturn = "OK"
    TempString = PARAMETRO
    
    If (InStr(1, TempString, Chr(34)) > 0) Then
        TempString = Replace(TempString, Chr(34), "")
        
        linefeed = Split(TempString, "\n")
        
        i = 0
        bFirstTime = True
        Do While (i <= UBound(linefeed))
            
            If (bFirstTime = True) Then
                bFirstTime = False
            Else
                NewString = NewString + vbCrLf
            End If
            
            NewString = NewString + linefeed(i)
            i = i + 1
        Loop
        
        STRINGA = NewString
        
    Else
        bReturn = "ERRORE IL PARAMETRO NON E' UN NUMERO"
    End If
    
    CONVERT_PARAMETRO_TO_STRINGA = bReturn
    
    Exit Function
err:
    CONVERT_PARAMETRO_TO_STRINGA = "ERROR " + err.Description

End Function





Private Function CONVERT_PARAMETRO_TO_POINTER(ByRef PARAMETRO As String, ByRef Pointer As Integer) As String
    On Error GoTo err:
    Dim bReturn As String
    Dim TempString As String
    
    bReturn = "OK"
    TempString = PARAMETRO
    
    If (InStr(1, TempString, "P") > 0) Then
        TempString = Replace(TempString, "P", "")
        Pointer = CInt(TempString)
        
    Else
        bReturn = "ERRORE IL PARAMETRO NON E' UN PUNTATORE"
    End If
    
    CONVERT_PARAMETRO_TO_POINTER = bReturn
    
    Exit Function
err:
    CONVERT_PARAMETRO_TO_POINTER = "ERROR " + err.Description
End Function



Private Function CONVERT_PARAMETRO_TO_REGISTER(ByRef PARAMETRO As String, ByRef REGISTER As Integer) As String
    On Error GoTo err:
    Dim bReturn As String
    Dim TempString As String
    
    bReturn = "OK"
    TempString = PARAMETRO
    
    If (InStr(1, TempString, "R") > 0) Then
        TempString = Replace(TempString, "R", "")
        REGISTER = CInt(TempString)
        
    Else
        bReturn = "ERRORE IL PARAMETRO NON E' UN REGISTRO"
    End If
    
    CONVERT_PARAMETRO_TO_REGISTER = bReturn
    
    Exit Function
err:
    CONVERT_PARAMETRO_TO_REGISTER = "ERROR " + err.Description
End Function


' QUesta funzione converte un parametro in una tupla
Private Function CONVERT_PARAMETRO_TO_TUPLA(ByRef PARAMETRO As String, ByRef TP As TagTupla) As String
    On Error GoTo err:
    Dim TempStr As String
    Dim bReturn As String
    Dim DatiTupla As Variant
    
    
    bReturn = "OK"
    
    TempStr = PARAMETRO
    
    If (InStr(1, TempStr, "(") = 1 And InStr(1, TempStr, ")") = Len(TempStr)) Then
        
        TempStr = Replace(TempStr, "(", "")
        TempStr = Replace(TempStr, ")", "")
        TempStr = Replace(TempStr, " ", "")
        
        DatiTupla = Split(TempStr, ".")
        
        If (UBound(DatiTupla) <> 2) Then
            bReturn = "STRUTTURA DELLA TUPLA ERRATA"
        Else
            TP.ID = CInt(DatiTupla(0))
            TP.TYPE = DatiTupla(1)
            TP.Addr = CInt(DatiTupla(2))
        End If
        
    Else
        bReturn = "MANCANO LE PARENTESI NELLA TUPLA"
        
    End If
    
    
    
    CONVERT_PARAMETRO_TO_TUPLA = bReturn
    Exit Function
err:
    CONVERT_PARAMETRO_TO_TUPLA = "ERRORE GENERICO " + err.Description
End Function



Private Sub UpdateStatusRegister(ByRef Value As Double, ByRef ALU As TagAlu)
    
    If (Value = 0) Then
        ALU.Z = True
    Else
        ALU.Z = False
    End If
    
    If (Value < 0) Then
        ALU.n = True
    Else
        ALU.n = False
    End If

End Sub

Public Function IsArrayInitialized(arr) As Boolean

  Dim rv As Long

  On Error GoTo err:

  rv = UBound(arr)
  IsArrayInitialized = True


Exit Function
err:
    IsArrayInitialized = False
End Function




Public Function GetCalibratedValue_INTERNAL(ByRef ALU As TagAlu, ByVal IdScheda As Integer, ByVal TipoVariabile As String, ByVal AddrVariabile As Integer) As Double

    On Error GoTo err:

    'FORMATO:   IdScheda.TipoVariabile.AddrVariabile

    If (IdScheda = 0 And TipoVariabile = "RICHTEXTBOX") Then
        
        Select Case AddrVariabile
            '(O.RICHTEXTBOX.0)
            Case 0: GetCalibratedValue_INTERNAL = ALU.Infos.RichTextBox1.BackColor
                        
            Case Else
                GetCalibratedValue_INTERNAL = 0
        
        End Select
    ElseIf (IdScheda = 0 And TipoVariabile = "LABEL") Then
        Select Case AddrVariabile
            '(O.LABEL.0)
            Case 0: GetCalibratedValue_INTERNAL = ALU.Infos.Label1.BackColor
                        
            Case Else
                GetCalibratedValue_INTERNAL = 0
        
        End Select
    Else
        
        If (IdScheda = -1) Then
            IdScheda = ALU.IndiceSezione + 1
        End If
        
        
        GetCalibratedValue_INTERNAL = GetCalibratedValue_EVO(IdScheda, TipoVariabile, AddrVariabile)
    End If
        
        
        
    Exit Function

err:
    GetCalibratedValue_INTERNAL = 0
    Debug.Print err.Description
    Exit Function

End Function 'END GetCalibratedValue_INTERNAL()


Public Function SetCalibratedValue_INTERNAL(ByRef ALU As TagAlu, ByVal IdScheda As Integer, ByVal TipoVariabile As String, ByVal AddrVariabile As Integer, ByVal ValueIn As Double) As Double

    On Error GoTo err:

    'FORMATO:   IdScheda.TipoVariabile.AddrVariabile
    SetCalibratedValue_INTERNAL = 0

    If (IdScheda = 0 And TipoVariabile = "RICHTEXTBOX") Then
        
        Select Case AddrVariabile
            '(O.RICHTEXTBOX.0)
            Case 0: ALU.Infos.RichTextBox1.BackColor = ValueIn
                        
            Case Else
            SetCalibratedValue_INTERNAL = 1
        
        End Select
    ElseIf (IdScheda = 0 And TipoVariabile = "LABEL") Then
        Select Case AddrVariabile
            '(O.LABEL.0)
            Case 0: ALU.Infos.Label1.BackColor = ValueIn
                        
            Case Else
                SetCalibratedValue_INTERNAL = 1
        
        End Select
    Else
        
        
        If (IdScheda = -1) Then
            IdScheda = ALU.IndiceSezione + 1
        End If
        
        If (AddrVariabile < 0) Then
            AddrVariabile = ALU.IndiceSezione - AddrVariabile - 1
        End If
        
        SetCalibratedValue_INTERNAL = SetCalibratedValue(IdScheda, TipoVariabile, AddrVariabile, ValueIn)
    End If
        
        
        
    Exit Function

err:
    SetCalibratedValue_INTERNAL = 1
    Debug.Print err.Description
    Exit Function

End Function 'END SetCalibratedValue_INTERNAL()
