Attribute VB_Name = "Module_SETTINGS"

' Struttura dati di scambio con il file di settaggio
Public Type TagSettings
    FolderConfigurazioneBancoCollaudo As String
    FolderConfigurazioneModuli As String                    ' Directory dove è presente il file di configurazione del modulo
    FolderConfigurazioneTest As String                      ' Directory dove sono presenti le liste dei test dei vari moduli
    FolderSettaggiProgramma As String                     ' Cartella dei file di calibrazione
    FolderFileCurveLimite As String                         ' Cartella di deposito delle curve limite
    FolderRampeXY As String                                 ' Cartella dove risiedono le curve di comando usate per le acquisizioni
    FolderFileCurveComando As String                        ' Cartella dove risiedono le curve comando CAn vs Volts
    FolderGraphSaved As String                              ' Cartella dove salvare il file delle acquisizioni
    FolderGraphSaved_LastACQ As String                      ' Cartella dove salvare l'ultima acquisizione il file delle acquisizioni
    FolderGraphSaved_AllACQ As String                       ' Cartella "GRAPH_XY_ALL"       --> dove salvare tutti i file delle acquisizioni
    FolderLOG As String                                     ' Cartella dove salvare il log dei file
    FolderGraphSaved_EVO As String                          ' Cartella dove salvare il file delle acquisizioni
    CanUsbName0 As String                                   ' Nome del convertitore CAN usb del banco 0
    CanUsbName1 As String                                   ' Nome del convertitore CAN usb del banco 1
End Type

Public Enum eErrorSettings
    ERROR_Settings
    OK_Settings
End Enum



Public Function ApriSettings(ByRef Context As TagSettings) As eErrorSettings
    Dim fileId As Long
    Dim FileName As String
    Dim LineaDati As String                 ' linea dei dati
    Dim Dati As Variant                     ' Dati
    Dim i As Integer
    
    Dim bReturn As eErrorSettings
    
    bReturn = ERROR_Settings
    
    
    FileName = GetSetting(App.EXEName, "Settings.ini", "Path", App.Path + "\settings.ini")
    
    On Error GoTo err:
    fileId = FreeFile
    'FileName = App.Path + "\settings.ini"

    Open FileName For Input As #fileId
  
    Do While (EOF(fileId) = False)
        Input #fileId, LineaDati
        Dati = Split(LineaDati, "=")                    ' Prelevo i dati
        
        If (UBound(Dati) = 1) Then                      ' Controllo quanti campi ci sono attivi
            Select Case (Dati(0))
                Case "FolderConfigurazioneBancoCollaudo"
                    Context.FolderConfigurazioneBancoCollaudo = Dati(1)
                Case "FolderConfigurazioneTest"
                    Context.FolderConfigurazioneTest = Dati(1)
                Case "FolderConfigurazioneModuli"
                    Context.FolderConfigurazioneModuli = Dati(1)
                Case "FolderSettaggiProgramma"
                    Context.FolderSettaggiProgramma = Dati(1)
                Case "FolderFileCurveLimite"
                    Context.FolderFileCurveLimite = Dati(1)
                Case "FolderRampeXY"
                    Context.FolderRampeXY = Dati(1)
                Case "FolderFileCurveComando"
                    Context.FolderFileCurveComando = Dati(1)
                Case "FolderGraphSaved"
                    Context.FolderGraphSaved = Dati(1)
                    
                    If (Context.FolderGraphSaved_AllACQ = "") Then
                        Context.FolderGraphSaved_AllACQ = Dati(1)
                    End If
                
                Case "FolderGraphSaved_EVO"                         ' Cartella "GRAPH_XY"
                    Context.FolderGraphSaved_EVO = Dati(1)
                    
                    If (Context.FolderGraphSaved_AllACQ = "") Then
                        Context.FolderGraphSaved_AllACQ = Dati(1)
                    End If
                Case "FolderLOG"
                    Context.FolderLOG = Dati(1)
                Case "FolderGraphSaved_LastACQ"
                    Context.FolderGraphSaved_LastACQ = Dati(1)
                Case "CAN1"
                    Context.CanUsbName0 = Dati(1)
                Case "CAN2"
                    Context.CanUsbName1 = Dati(1)
                Case "FolderGraphSaved_AllACQ"                  ' Cartella "GRAPH_XY_ALL"
                    Context.FolderGraphSaved_AllACQ = Dati(1)
            End Select
            
        End If
    Loop
    bReturn = OK_Settings
    
    ' Se non è statao definito la variabile FolderGraphSaved_EVO allora prendo
    ' FolderGraphSaved come valore di default
    If (Context.FolderGraphSaved_EVO = "") Then
        Context.FolderGraphSaved_EVO = Context.FolderGraphSaved
    End If
    
    
    
    ApriSettings = bReturn
    Close #fileId
    Exit Function
err:
    ApriSettings = bReturn
    Close #fileId
End Function


Public Function SaveSettings(ByRef Context As TagSettings) As eErrorSettings
'    Dim fileId As Long
'    Dim FileName As String
'    Dim LineaDati As String                 ' linea dei dati
'    Dim Dati As Variant                     ' Dati
'
'    Dim bReturn As eErrorSettings
'
'    bReturn = ERROR_Settings
'
'    On Error GoTo err:
'    fileId = FreeFile
'    FileName = App.Path + "\settings.ini"
'
'    Open FileName For Output As #fileId
'
'
'
'    Close #fileId
'    Exit Function
'err:
'    SaveSettings = bReturn
'    Close #fileId
End Function



