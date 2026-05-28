Attribute VB_Name = "Module_LOG_FILE"
Public Function GetFilePath(ByVal Path As String, ByVal codicePRodotto As String, ByVal SerialNumber As String) As String
    Dim FilePath As String
    
    If (Right(Path, 1) <> "\") Then
        Path = Path + "\"
    End If
    
    Path = Path + codicePRodotto
    
    VerificaOCreaCartella Path
    
    If (Right(Path, 1) <> "\") Then
        Path = Path + "\"
    End If
    
    ' Determina il percorso del file di log
    FilePath = Path & SerialNumber & ".log"

    GetFilePath = FilePath
End Function




Public Sub SalvaRisultatiTest(ByVal Path As String, Infos As TagInfosSingleTest, testResults As String)
On Error GoTo err:
    Dim FilePath As String
    
    FilePath = GetFilePath(Path, Infos.codicePRodotto.Caption, Infos.Label1.Caption)
    
    ' Ottiene un numero di file libero
    fileNumber = FreeFile
    
    ' Apre il file in modalità Append (aggiunge alla fine del file esistente)
    Open FilePath For Append As #fileNumber
    
    ' Scrive i risultati dei test nel file di log
    
    Print #fileNumber, "Data: " & Now
    Print #fileNumber, "Nome Test: " + Infos.CWButtonTest.OffText
    Print #fileNumber, "Risultati del Test: "
    Print #fileNumber, String(50, "-") ' Linea di separazione per i test successivi
    Print #fileNumber, testResults
    ' Chiude il file
    Close #fileNumber
    'MsgBox "CLOSE"
    Exit Sub
err:
Close #fileNumber
    MsgBox err.Description
End Sub


Public Sub AggiungiRiga(Path As String, Infos As TagInfosSingleTest, Riga As String)
On Error GoTo err:
    Dim FilePath As String
    Dim fileNumber As Integer
    If (Right(Path, 1) <> "\") Then
        Path = Path + "\"
    End If
    
    Path = Path + Infos.codicePRodotto
    
    VerificaOCreaCartella Path
    
    If (Right(Path, 1) <> "\") Then
        Path = Path + "\"
    End If
    
    ' Determina il percorso del file di log
    FilePath = Path & Infos.Label1.Caption & ".log"
        
    ' Ottiene un numero di file libero
    fileNumber = FreeFile
    
    ' Apre il file in modalità Append (aggiunge alla fine del file esistente)
    Open FilePath For Append As #fileNumber
    
    ' Scrive i risultati dei test nel file di log
    Print #fileNumber, Riga
err:
    ' Chiude il file
    Close #fileNumber

End Sub


Private Sub VerificaOCreaCartella(cartella As String)
    ' Verifica se la cartella esiste
    If Dir(cartella, vbDirectory) = "" Then
        ' Se non esiste, crea la cartella
        On Error GoTo ErroreCreazione
        MkDir cartella
        'MsgBox "Cartella creata con successo: " & cartella
        Exit Sub
ErroreCreazione:
        'MsgBox "Errore nella creazione della cartella: " & cartella, vbCritical
    Else
        'MsgBox "La cartella esiste già: " & cartella
    End If
End Sub





Public Sub Prepare_MsgBox(Infos As TagInfosSingleTest, ByVal str As String, ByVal tipo As Long)
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "-------------------------" + vbLf
    Infos.RichTextBox1.Text = Infos.RichTextBox1.Text + "ERROR - " + str + vbLf
    MsgBox "str", tipo
End Sub

