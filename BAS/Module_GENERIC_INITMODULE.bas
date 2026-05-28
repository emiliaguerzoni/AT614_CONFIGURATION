Attribute VB_Name = "Module_GENERIC_INITMODULE"

Private Declare Function GetTickCount Lib "kernel32" () As Long
Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Declare Function QueryPerformanceFrequency Lib "kernel32" (X As Currency) As Boolean

Private First As Boolean
Private Ctr1 As Currency
Private Freq1 As Currency


Public Enum eAdcMM2218
    eSegnaleFeedbackFED5_1 = 0
    eSegnaleFeedbackFED5_2 = 1
    eSegnaleFeedbackFED5_3 = 2
    eSegnaleFeedbackFED5_4 = 3
    eSegnaleFeedbackFED5_5 = 4
    eTemperatura = 5
    ePressione = 6
End Enum

Public Sub TurnOnAllModule()
    Dim i As Integer
    Do While (i < MAX_SEZIONI)
        'MLTPower(i) = False
        MMS2218.DOUT(0).Value(i) = True
        i = i + 1
    Loop
End Sub


Public Sub TurnOffAllModule()
    Dim i As Integer
    Do While (i < MAX_SEZIONI)
        'MLTPower(i) = False
        MMS2218.DOUT(0).Value(i) = False
        i = i + 1
    Loop
End Sub

Public Sub TurnOnModule(ByVal Indice As Integer)
    'MLTPower(Indice) = True
    MMS2218.DOUT(0).Value(Indice) = True
End Sub

Public Function GetModuleState(ByVal Indice As Integer) As Boolean
    'MLTPower(Indice) = True
    GetModuleState = MMS2218.DOUT(0).Value(Indice)
End Function


Public Sub TurnOffModule(ByVal Indice As Integer)
    'MLTPower(Indice) = False
    MMS2218.DOUT(0).Value(Indice) = False
End Sub

Public Function InitCalibration()
    Dim i As Integer
    Dim NomeComputer As String
    Dim bMsgBox As Boolean
    NomeComputer = TrovaNomeComputer()
    
    CreadDirectory (Settings.FolderSettaggiProgramma + "\" + NomeComputer + "\MMS2218")
    Do While (i < UBound(Calib2218))
        Calib2218(i).FileName = Settings.FolderSettaggiProgramma + "\" + NomeComputer + "\MMS2218\ADC" + CStr(i) + ".csv"
        If (ApriCalibration(Calib2218(i)) = True And bMsgBox = False) Then
            MsgBox "Errore durante apertura file di calibrazione !!!!! " + Calib2218(i).FileName, vbCritical
            bMsgBox = True
        End If
        i = i + 1
    Loop
    

    
    
    SetSourceAddress MMS2218, &HA0
    Configure_Tx_DO_Message MMS2218, 0, 500, 13
    Configure_Tx_AVC_Message MMS2218, 0, 50
    Configure_Tx_AVC_Message MMS2218, 1, 50
    Configure_Tx_AVC_Message MMS2218, 2, 50
    Configure_Tx_AVC_Message MMS2218, 3, 50
    Configure_Tx_AVC_Message MMS2218, 4, 50
    Configure_Tx_AVC_Message MMS2218, 5, 50
    Configure_Tx_AVC_Message MMS2218, 6, 50
    Configure_Tx_AVC_Message MMS2218, 7, 50
    Configure_Tx_AVC_Message MMS2218, 8, 50
    Configure_Tx_AVC_Message MMS2218, 9, 50
    
    
End Function


Public Function CreadDirectory(pathdacreare As String) As Boolean
On Error GoTo err:
    MkDir (pathdacreare)
    Exit Function
err:
    Debug.Print err.Description
End Function


Public Function GetCalibratedValue(ByVal Index As eAdcMM2218) As Double
    Dim MessageIndex As Integer
    Dim AdcIndex As Integer
    Dim i As Integer
    Dim ValueOut As Double
    
    Do While (i < Index)
            
        AdcIndex = AdcIndex + 1
        
        If (AdcIndex >= 4) Then
            MessageIndex = MessageIndex + 1
            AdcIndex = 0
        End If
        
        i = i + 1
    Loop
    
    If (GetCalibrationPoint(Calib2218(Index), MMS2218.AIN16(MessageIndex).Value(AdcIndex), ValueOut) = False) Then
        GetCalibratedValue = ValueOut
    End If
End Function



Public Sub PPDO_Manager()
    Dim MessaggioCan As String
    
    canUsbTask
    
    Do While (GetFrameFromRxBuffer(e2218, MessaggioCan) = eOK)
        DecodeMessage MMS2218, MessaggioCan
    Loop


    ManageTxMessage MMS2218, MessaggioCan
    
    
    If (MessaggioCan <> "") Then
        Output MessaggioCan
        MessaggioCan = ""
    End If
    
    
     ManageTxAVCMessage MMS2218, MessaggioCan
    
    If (MessaggioCan <> "") Then
        Output MessaggioCan
        MessaggioCan = ""
    End If
    
    ManageTX_MLT_Comandi MMS2218, MessaggioCan
    
    If (MessaggioCan <> "") Then
        Output MessaggioCan
        MessaggioCan = ""
    End If
    
End Sub


Public Function SetCalibratedValue(ByVal IdScheda As Integer, ByVal TipoVariabile As String, ByVal AddrVariabile As Integer, ByVal ValueIn As Double) As Double

    Dim MessageIndex As Integer
    Dim AdcIndex As Integer
    Dim i As Integer
    Dim Value As Double
    
    On Error GoTo err:
    
    ' 'SA' = identificativo della scheda sulla rete

       
       
    
    Select Case TipoVariabile
                   
        Case "DOUT":
            If (AddrVariabile < 64) Then
                MMS2218.DOUT(0).Value(AddrVariabile) = ValueIn
            End If
        
            MMS2218.DOUT(0).TimeStamp = GetTickCountEvo
            
        Case "AOUT16":
            
            Do While (i < AddrVariabile)
                AdcIndex = AdcIndex + 1
                If (AdcIndex >= 4) Then
                    MessageIndex = MessageIndex + 1
                    AdcIndex = 0
                End If
                
                i = i + 1
            Loop
            
            MMS2218.AOUT16(MessageIndex).Value(AdcIndex) = ValueIn
            MMS2218.AOUT16(MessageIndex).TimeStamp = GetTickCountEvo
        
        Case "MLT_COMMAND"
            If (IdScheda <= UBound(MMS2218.MLT_COMANDI)) Then
                
                If (AddrVariabile <= UBound(MMS2218.MLT_COMANDI(IdScheda).Comandi)) Then
                    MMS2218.MLT_COMANDI(IdScheda).Comandi(AddrVariabile).Value = ValueIn
                    MMS2218.MLT_COMANDI(IdScheda).Comandi(AddrVariabile).Richiesta = True
                
                ElseIf (AddrVariabile < 100 + UBound(MMS2218.MLT_COMANDI(IdScheda).PARAMETRI)) Then
                    
                    MMS2218.MLT_COMANDI(IdScheda).PARAMETRI(AddrVariabile - 100) = ValueIn
                                
                Else
                    GoTo err
                End If
            
            End If
        
        Case "MLT"
            
            If (IdScheda <= UBound(MMS2218.MLT)) Then
                
                If (AddrVariabile <= UBound(MMS2218.MLT(IdScheda).Dato)) Then
                    MMS2218.MLT(IdScheda).Dato(AddrVariabile).Value = ValueIn
                    MMS2218.MLT(IdScheda).Dato(AddrVariabile).TimeStamp = GetTickCountEvo
                    'MMS2218.INTERNAL(0).TimeStamp = GetTickCount
                Else
                    GoTo err
                End If
            
            End If
            
        Case "INTERNAL":
        
            If (IdScheda = 0) Then
                If (AddrVariabile <= UBound(MMS2218.INTERNAL(0).Value)) Then
                    MMS2218.INTERNAL(0).Value(AddrVariabile) = ValueIn
                    'MMS2218.INTERNAL(0).TimeStamp = GetTickCount
                Else
                    GoTo err
                End If
            ElseIf IdScheda <> MMS2218.SA Then
                GoTo err
            End If
        
        Case "AVC"
            If (IdScheda <= UBound(MMS2218.AVC)) Then
                'MMS2218.AVC(IdScheda).TimeStamp = GetTickCount
                
                If (AddrVariabile <= UBound(MMS2218.AVC(IdScheda).Value)) Then
                    MMS2218.AVC(IdScheda).Value(AddrVariabile) = ValueIn
                    
                    MMS2218.AVC(IdScheda).TimeRequest = GetTickCountEvo
                Else
                    GoTo err
                End If
            
            End If
        
        
        
        Case "AIN16":
        Case "AIN8":
        Case "DIN":
        
        Case Else: GoTo err
    
    End Select
    
    Exit Function

err:
    Debug.Print err.Description
    Exit Function


End Function 'END SetCalibratedValue()

Public Function GetCalibratedTimeStamp_EVO(IdScheda As Integer, TipoVariabile As String, ByVal AddrVariabile As Integer) As Double
    Select Case TipoVariabile
    
        Case "DIN":
                    
            i = 0
            Do While (i < AddrVariabile)
                    
                DinIndex = DinIndex + 1
                
                If (DinIndex >= 64) Then
                    MessageIndex = MessageIndex + 1
                    DinIndex = 0
                End If
                
                i = i + 1
            Loop
                    
            GetCalibratedTimeStamp_EVO = MMS2218.DIN(MessageDinIndex).TimeStamp
            
            
        
        Case "DOUT":
        
            If IdScheda <> MMS2218.SA Then
                GoTo err
            End If
                    
            i = 0
            Do While (i < AddrVariabile)
                    
                DinIndex = DinIndex + 1
                
                If (DinIndex >= 64) Then
                    MessageIndex = MessageIndex + 1
                    DinIndex = 0
                End If
                
                i = i + 1
            Loop
        
            GetCalibratedTimeStamp_EVO = MMS2218.DOUT(MessageDinIndex).TimeStamp
        
            
            
        Case "AIN16":
            If IdScheda <> MMS2218.SA Then
                GoTo err
            End If
            
            
            GetCalibratedTimeStamp_EVO = MMS2218.DOUT(MessageIndex).TimeStamp
                    
                                                            
        
        Case "AIN8":
       
        Case "AOUT16":
            
            If IdScheda <> MMS2218.SA Then
                GoTo err
            End If
            
            
            GetCalibratedTimeStamp_EVO = MMS2218.AOUT16(MessageIndex).TimeStamp

                    
                    
        Case "CE16":
            If (IdScheda < UBound(MMS2218.CE16)) Then
                
                ' Tick di sistema
                GetCalibratedTimeStamp_EVO = MMS2218.CE16(IdScheda).TimeStamp
                
            End If
                    
        Case "MLT":
            If (IdScheda < UBound(MMS2218.MLT)) Then

                ' Tick di sistema
                
                
                
                If (AddrVariabile <= UBound(MMS2218.MLT(IdScheda).Dato)) Then
                    GetCalibratedTimeStamp_EVO = MMS2218.MLT(IdScheda).Dato(AddrVariabile).TimeStamp
                End If
                

            End If
            
        Case "MLT_COMMAND"

        Case "AVC"
        
            If (IdScheda <= UBound(MMS2218.AVC)) Then
                    
                GetCalibratedTimeStamp_EVO = MMS2218.AVC(IdScheda).TimeStamp
            
            End If
        
        
        
        
        
        Case "INTERNAL":
            If (IdScheda = 0) Then
                Select Case (AddrVariabile)
                ' Tick di sistema
                Case 0:
                    GetCalibratedTimeStamp_EVO = GetTickCountEvo
                Case Else
                    
                    If (AddrVariabile <= UBound(MMS2218.INTERNAL(0).Value)) Then
                        GetCalibratedTimeStamp_EVO = MMS2218.INTERNAL(0).Value(AddrVariabile)
                    End If
                    
                End Select
            End If
        Case Else: GoTo err
    
    End Select
    
    Exit Function

err:
    GetCalibratedTimeStamp_EVO = 0
    Debug.Print "GetCalibratedValue_EVO : " + err.Description
    Debug.Print
    'MsgBox "TRAPPOLA ERROR"
    Exit Function


End Function




Public Function GetCalibratedValue_EVO(IdScheda As Integer, TipoVariabile As String, ByVal AddrVariabile As Integer) As Double

    Dim MessageIndex As Integer
    Dim AdcIndex As Integer
    Dim MessageDinIndex As Integer
    Dim DinIndex As Integer
    Dim ValueInBoolean As Double
    
    Dim i As Integer
    Dim ValueOut As Double
    
    On Error GoTo err:
                    

    
    Select Case TipoVariabile
    
        Case "DIN":
                    
            i = 0
            Do While (i < AddrVariabile)
                    
                DinIndex = DinIndex + 1
                
                If (DinIndex >= 64) Then
                    MessageIndex = MessageIndex + 1
                    DinIndex = 0
                End If
                
                i = i + 1
            Loop
        
            If (MMS2218.DIN(MessageDinIndex).Value(DinIndex) = True) Then
                ValueInBoolean = 1
            Else
                ValueInBoolean = 0
            End If
            
            GetCalibratedValue_EVO = ValueInBoolean
            
            
'            If (GetCalibrationPoint(MMS6252.Calib6252DIN(AddrVariabile), ValueInBoolean, ValueOut) = False) Then
'                GetCalibratedValue = ValueOut
'            End If
        
        Case "DOUT":
        
            If IdScheda <> MMS2218.SA Then
                GoTo err
            End If
                    
            i = 0
            Do While (i < AddrVariabile)
                    
                DinIndex = DinIndex + 1
                
                If (DinIndex >= 64) Then
                    MessageIndex = MessageIndex + 1
                    DinIndex = 0
                End If
                
                i = i + 1
            Loop
        
            If (MMS2218.DOUT(MessageDinIndex).Value(DinIndex) = True) Then
                ValueInBoolean = 1
            Else
                ValueInBoolean = 0
            End If
            
'            If (GetCalibrationPoint(MMS2218.Calib6252DOUT(AddrVariabile), ValueInBoolean, ValueOut) = False) Then
'                GetCalibratedValue = ValueOut
'            End If
            GetCalibratedValue_EVO = ValueInBoolean
            
        Case "AIN16":
            If IdScheda <> MMS2218.SA Then
                GoTo err
            End If
            
            i = 0
            Do While (i < AddrVariabile)
                    
                AdcIndex = AdcIndex + 1
                
                If (AdcIndex >= 4) Then
                    MessageIndex = MessageIndex + 1
                    AdcIndex = 0
                End If
                
                i = i + 1
            Loop
            
            GetCalibratedValue_EVO = MMS6252.AIN16(MessageIndex).Value(AdcIndex)
            
                                                               
                                                            '(*** MessageIndex * 4 + AdcIndex)

'            If (GetCalibrationPoint(MMS6252.Calib6252AI16(AddrVariabile), MMS6252.AIN16(MessageIndex).Value(AdcIndex), ValueOut) = False) Then
'                GetCalibratedValue = ValueOut
'            End If
        
        
        Case "AIN8":
       
        Case "AOUT16":
            
            If IdScheda <> MMS2218.SA Then
                GoTo err
            End If
            
            i = 0
            Do While (i < AddrVariabile)
                    
                AdcIndex = AdcIndex + 1
                
                If (AdcIndex >= 4) Then
                    MessageIndex = MessageIndex + 1
                    AdcIndex = 0
                End If
                
                i = i + 1
            Loop
            
            GetCalibratedValue_EVO = MMS2218.AOUT16(MessageIndex).Value(AdcIndex)

            
'            If (GetCalibrationPoint(MMS6252.Calib6252AO(AddrVariabile), MMS2218.AOUT16(MessageIndex).Value(AdcIndex), ValueOut) = False) Then
'                GetCalibratedValue = ValueOut
'            End If
                    
                    
        Case "CE16":
            If (IdScheda < UBound(MMS2218.CE16)) Then
                
                ' Tick di sistema
                    
                If (AddrVariabile <= UBound(MMS2218.CE16(IdScheda).Value)) Then
                    GetCalibratedValue_EVO = MMS2218.CE16(IdScheda).Value(AddrVariabile)
                End If
                
                
            End If
                    
        Case "MLT":
            If (IdScheda < UBound(MMS2218.MLT)) Then

                ' Tick di sistema
                    
                If (AddrVariabile <= UBound(MMS2218.MLT(IdScheda).Dato)) Then
                    GetCalibratedValue_EVO = MMS2218.MLT(IdScheda).Dato(AddrVariabile).Value
                End If
                

            End If
            
        Case "MLT_COMMAND"
            If (IdScheda < UBound(MMS2218.MLT_COMANDI)) Then
                
                If (AddrVariabile < UBound(MMS2218.MLT_COMANDI(IdScheda).Comandi)) Then
                    
                    GetCalibratedValue_EVO = MMS2218.MLT_COMANDI(IdScheda).Comandi(AddrVariabile).Value
                
                ElseIf (AddrVariabile < 100 + UBound(MMS2218.MLT_COMANDI(IdScheda).PARAMETRI)) Then
                    
                    GetCalibratedValue_EVO = MMS2218.MLT_COMANDI(IdScheda).PARAMETRI(AddrVariabile - 100)
                    
                End If
            End If
        
        Case "AVC"
        
            If (IdScheda <= UBound(MMS2218.AVC)) Then
                
                If (AddrVariabile <= UBound(MMS2218.AVC(IdScheda).Value)) Then
                    
                    GetCalibratedValue_EVO = MMS2218.AVC(IdScheda).Value(AddrVariabile)
                    
                Else
                    GoTo err
                End If
            
            End If
        
        
        
        
        
        Case "INTERNAL":
            If (IdScheda = 0) Then
                Select Case (AddrVariabile)
                ' Tick di sistema
                Case 0:
                    GetCalibratedValue_EVO = GetTickCountEvo
                Case Else
                    
                    If (AddrVariabile <= UBound(MMS2218.INTERNAL(0).Value)) Then
                        GetCalibratedValue_EVO = MMS2218.INTERNAL(0).Value(AddrVariabile)
                    End If
                    
                End Select
            End If
        Case Else: GoTo err
    
    End Select
    
    Exit Function

err:
    GetCalibratedValue_EVO = 0
    Debug.Print "GetCalibratedValue_EVO : " + err.Description
    Debug.Print
    'MsgBox "TRAPPOLA ERROR"
    Exit Function

End Function 'END GetCalibratedValue()











' Questa funzione apre il file selezionato
Public Function ApriFileConfigurazioneDistributore(ByRef Context1 As TagTagDistributore, FileName As String)
    Dim fileId As Long

    Dim LineaDati As String                 ' linea dei dati
    Dim Dati As Variant                     ' Dati
    Dim i As Integer
    Dim Index As Integer
    Dim j As Long
    
    
    On Error GoTo err:
    fileId = FreeFile
    'FileName = Settings.FolderConfigurazioneBancoCollaudo + "\" + Context.NomeDistributoreSelezionato


    Do While (i < MAX_SEZIONI)
        Sezione(i).CodiceModulo = ""
        Sezione(i).OldCodiceModulo = ""
        Sezione(i).Visible = False
        Context1.NomeFileCalibrazioneCe16(i) = ""
        
        
        ' Resetto tutti i test
        j = 0
        Do While (j < Sezione(i).CWButtonTest.count)
            Sezione(i).CWButtonTest(j).Value = False
            j = j + 1
        Loop
        
        
        i = i + 1
    Loop
    

    Open FileName For Input As #fileId
    
    ReDim Context1.Custom(0)
  
  
    Do While (EOF(fileId) = False)
        Input #fileId, LineaDati
        Dati = Split(LineaDati, "=")                    ' Prelevo i dati
        
        If (UBound(Dati) = 1) Then                      ' Controllo quanti campi ci sono attivi
            Select Case (Dati(0))
                Case "sezione1"
                    ApriFileTest Settings.FolderConfigurazioneTest + "\" + Dati(1) + ".csv", TestSezione(0)
                    Sezione(0).CodiceModulo = Dati(1)
                Case "sezione2"
                    ApriFileTest Settings.FolderConfigurazioneTest + "\" + Dati(1) + ".csv", TestSezione(1)
                    Sezione(1).CodiceModulo = Dati(1)
                Case "sezione3"
                    ApriFileTest Settings.FolderConfigurazioneTest + "\" + Dati(1) + ".csv", TestSezione(2)
                    Sezione(2).CodiceModulo = Dati(1)
                Case "sezione4"
                    ApriFileTest Settings.FolderConfigurazioneTest + "\" + Dati(1) + ".csv", TestSezione(3)
                    Sezione(3).CodiceModulo = Dati(1)
                Case "sezione5"
                    ApriFileTest Settings.FolderConfigurazioneTest + "\" + Dati(1) + ".csv", TestSezione(4)
                    Sezione(4).CodiceModulo = Dati(1)
                Case "calibrazioneCE16_1"
                    Context1.NomeFileCalibrazioneCe16(0) = Dati(1)
                Case "calibrazioneCE16_2"
                    Context1.NomeFileCalibrazioneCe16(1) = Dati(1)
                Case "calibrazioneCE16_3"
                    Context1.NomeFileCalibrazioneCe16(2) = Dati(1)
                Case "calibrazioneCE16_4"
                    Context1.NomeFileCalibrazioneCe16(3) = Dati(1)
                Case "calibrazioneCE16_5"
                    Context1.NomeFileCalibrazioneCe16(4) = Dati(1)
                Case "calibrazioneCE16_6"
                    Context1.NomeFileCalibrazioneCe16(5) = Dati(1)
                Case "Temperatura_olio_max"
                    Context1.Temperatura_olio_max = Dati(1)
                Case "Temperatura_olio_min"
                    Context1.Temperatura_olio_min = Dati(1)
                Case "Pressure_max"
                    Context1.Pressure_max = Dati(1)
                Case "Pressure_min"
                    Context1.Pressure_min = Dati(1)
                    
                Case "dato":
                    
                    If (CaricaDato(Context1.Custom(UBound(Context1.Custom)), Dati(1)) = False) Then
                        ReDim Preserve Context1.Custom(UBound(Context1.Custom) + 1)
                    End If
                    
            End Select
            
        End If
    Loop
    bReturn = OK_Settings

    Close #fileId
    Exit Function
err:
    Close #fileId
End Function




Private Function CaricaDato(ByRef Context1 As custom_Dati_Distributore, ByVal Valore As String) As Boolean
    On Error GoTo err:
    Dim Dati As Variant
    Dim Tupla As Variant
    
    Valore = Replace(Valore, " ", "")
    Valore = Replace(Valore, "(", "")
    Valore = Replace(Valore, ")", "")
    
    
    Dati = Split(Valore, ":")
    Tupla = Split(Dati(1), ".")
    
       
    Context1.NOME = Dati(0)
    Context1.Tupla.IdScheda = Tupla(0)
    Context1.Tupla.TipoVariabile = Tupla(1)
    Context1.Tupla.AddrVariabile = Tupla(2)
    Exit Function

err:

    CaricaDato = True
    
End Function


Public Sub AllineamentoVariabili()
End Sub


' Questa funzione allinea le variabili interne al modulo alle variabili del protocollo di comunicazione utilizzato
Public Sub AllineaVariabili()
    Dim i As Integer
    
'    Do While (i < MAX_SEZIONI)
'        Select Case i
'        Case 0: MMS2218.DOUT(0).Value(0) = MLTPower(0)
'        Case 1: MMS2218.DOUT(0).Value(1) = MLTPower(1)
'        Case 2: MMS2218.DOUT(0).Value(2) = MLTPower(2)
'        Case 3: MMS2218.DOUT(0).Value(3) = MLTPower(3)
'        Case 4: MMS2218.DOUT(0).Value(4) = MLTPower(4)
'        End Select
'
'        i = i + 1
'    Loop
    
    i = 0
    Do While (i < MAX_SEZIONI)
        Select Case i
        Case 0: MMS2218.DOUT(0).Value(5) = True 'MLTCANRelais(0)
        Case 1: MMS2218.DOUT(0).Value(6) = True 'MLTCANRelais(1)
        Case 2: MMS2218.DOUT(0).Value(7) = True 'MLTCANRelais(2)
        Case 3: MMS2218.DOUT(0).Value(8) = True 'MLTCANRelais(3)
        Case 4: MMS2218.DOUT(0).Value(9) = True 'MLTCANRelais(4)
        End Select
        
        i = i + 1
    Loop
    
End Sub

Public Function Attesa_Plus(ByVal TempoAttesa As Double)

Dim cont As Long
Dim elapsed As Long
Dim MessaggioCan As String

elapsed = 0
TempoAttesa = TempoAttesa * 1000

cont = GetTickCount + TempoAttesa

Do While (GetTickCount <= cont)
    
    canUsbTask
    
    If GetTickCount > elapsed Then
        elapsed = GetTickCount + 10
        
        AllineaVariabili
        'AllineamentoVariabili
        
        Do While (GetFrameFromRxBuffer(e2218, MessaggioCan) = eOK)
            DecodeMessage MMS2218, MessaggioCan
    
        Loop


        ManageTxMessage MMS2218, MessaggioCan
        
        
        If (MessaggioCan <> "") Then
            Output MessaggioCan
            MessaggioCan = ""
        End If
    
    
        ManageTxAVCMessage MMS2218, MessaggioCan
        
        
        If (MessaggioCan <> "") Then
            Output MessaggioCan
            MessaggioCan = ""
        End If
    

        ManageTX_MLT_Comandi MMS2218, MessaggioCan
        
        If (MessaggioCan <> "") Then
            Output MessaggioCan
            MessaggioCan = ""
        End If


        DoEvents
    End If
Loop

End Function 'END "Attesa()"




'
'Public Function GetTickCountEvo() As Long
'    Dim cTimer As Currency
'    Dim Freq As Currency
'
'    QueryPerformanceCounter cTimer
'    QueryPerformanceFrequency Freq
'
'
'    GetTickCountEvo = cTimer / Freq * 1000
'
'End Function

    
    
Public Function GetTickCountEvo() As Long
    

    

    If (First = False) Then
        First = True
        QueryPerformanceFrequency Freq1
        QueryPerformanceCounter Ctr1
    End If
    
    
      QueryPerformanceCounter Ctr2
      
      'Debug.Print "Start Value: "; Format$(Ctr1, "0.0000")
      'Debug.Print "End Value: "; Format$(Ctr2, "0.0000")
      'QueryPerformanceFrequency Freq
      'Debug.Print "QueryPerformanceCounter minimum resolution: 1/" & _
      '            Freq * 10000; " sec"
      'Debug.Print "API Overhead: " + CStr((1000 * (Ctr2 - Ctr1)) / Freq1) + "seconds"
          
    GetTickCountEvo = (1000 * (Ctr2 - Ctr1)) / Freq1
End Function
