Attribute VB_Name = "MMS2218_CAN"
    
Private Declare Function GetTickCount& Lib "kernel32" ()

' Questo modulo gestisce la comunicazione con il dispostivo CAN MMS2218
' Si definiscono 3 tipologie di messaggi:
' 1) PRPDO - Pseduo PTO. Questo messaggio ciclico viene trasmesso ad event time
' 1) PTPDO - Pseduo PTO. Questo messaggio ciclico viene riceuto dalla 2218 ad evento
' 1) CPROTOCOL  Questo messaggio di natura Master slave permette di accedere alle aree dati interne del modulo

' Lista PGN definita:
' TX ingressi digitali ( PRPDO )
Enum ePRPDO_DI_PGN
    PRPDO_DI_0 = &HFF00 '- DI 0-63
    PRPDO_DI_1 = &HFF01 '- DI 64-127
    PRPDO_DI_2 = &HFF02 '- DI 128-191
    PRPDO_DI_3 = &HFF04 '- DI 192-256
    PRPDO_DI_MAX
End Enum

' TX ingressi analogici ( PRPDO ) Valori letti direttamente dal ADC 16 bit
Enum ePRPDO_AI16_PGN
    PRPDO_AI16_0 = &HFF20 '- AI16 0-3
    PRPDO_AI16_1 = &HFF21 '- AI16 4-7
    PRPDO_AI16_2 = &HFF22 '- AI16 8-11
    PRPDO_AI16_3 = &HFF23 '- AI16 12-15
    PRPDO_AI16_4 = &HFF24 '- AI16 16-19
    PRPDO_AI16_5 = &HFF25 '- AI16 20-23
    PRPDO_AI16_6 = &HFF26 '- AI16 24-27
    PRPDO_AI16_7 = &HFF27 '- AI16 28-31
    PRPDO_AI16_MAX
End Enum

' TX ingressi analogici ( PRPDO ) Valori letti direttamente dal ADC 8 bit
Enum ePRPDO_AI8_PGN
    PRPDO_AI8_0 = &HFF20 '- AI8 0-7
    PRPDO_AI8_1 = &HFF21 '- AI8 8-15
    PRPDO_AI8_2 = &HFF22 '- AI8 16-23
    PRPDO_AI8_3 = &HFF23 '- AI8 24-31
    PRPDO_AI8_4 = &HFF24 '- AI8 32-39
    PRPDO_AI8_5 = &HFF25 '- AI8 40-47
    PRPDO_AI8_6 = &HFF26 '- AI8 48-53
    PRPDO_AI8_7 = &HFF27 '- AI8 54-63
    PRPDO_AI8_MAX
End Enum

' RX uscite digitali Da qualsiasi source address ( PTPDO )
Enum ePTPDO_DO_PGN
    PTPDO_DO_0 = &HFF80 '- DO 0-63
    PTPDO_DO_1 = &HFF81 '- DO 64-127
    PTPDO_DO_2 = &HFF82 '- DO 128-191
    PTPDO_DO_3 = &HFF84 '- DO 192-256
    PTPDO_DO_MAX
End Enum
' RX uscite analogiche da qualsiasi source address ( PTPDO ) Valori da scrivere direttamente dal DAC 16 bit
Enum ePTPDO_AO16_PGN
    PTPDO_AO16_0 = &HFF90 '- AO16 0-3
    PTPDO_AO16_1 = &HFF91 '- AO16 4-7
    PTPDO_AO16_2 = &HFF92 '- AO16 8-11
    PTPDO_AO16_3 = &HFF93 '- AO16 12-15
    PTPDO_AO16_4 = &HFF94 '- AO16 16-19
    PTPDO_AO16_5 = &HFF95 '- AO16 20-23
    PTPDO_AO16_6 = &HFF96 '- AO16 24-27
    PTPDO_AO16_7 = &HFF97 '- AO16 28-31
    PTPDO_AO16_MAX
End Enum

' RX uscite analogiche da qualsiasi source address ( PTPDO ) Valori da scrivere direttamente dal DAC 8 bit
Enum ePTPDO_AO8_PGN
    PTPDO_AO8_0 = &HFFA0 '- AO8 0-7
    PTPDO_AO8_1 = &HFFA1 '- AO8 8-15
    PTPDO_AO8_2 = &HFFA2 '- AO8 16-23
    PTPDO_AO8_3 = &HFFA3 '- AO8 24-31
    PTPDO_AO8_4 = &HFFA4 '- AO8 32-39
    PTPDO_AO8_5 = &HFFA5 '- AO8 40-47
    PTPDO_AO8_6 = &HFFA6 '- AO8 48-55
    PTPDO_AO8_7 = &HFFA7 '- AO8 56-63
    PTPDO_AO8_MAX
End Enum


' RX uscite analogiche da qualsiasi source address ( PTPDO ) Valori da scrivere direttamente dal DAC 8 bit
Enum ePTPDO_ESTIMATED_FLOW_PGN
    PTPDO_ESTIMATED_FLOW_00 = &HFE10
    PTPDO_ESTIMATED_FLOW_01 = &HFE11
    PTPDO_ESTIMATED_FLOW_02 = &HFE12
    PTPDO_ESTIMATED_FLOW_03 = &HFE13
    PTPDO_ESTIMATED_FLOW_04 = &HFE14
    PTPDO_ESTIMATED_FLOW_05 = &HFE15
    PTPDO_ESTIMATED_FLOW_06 = &HFE16
    PTPDO_ESTIMATED_FLOW_07 = &HFE17
    PTPDO_ESTIMATED_FLOW_08 = &HFE18
    PTPDO_ESTIMATED_FLOW_09 = &HFE19
    PTPDO_ESTIMATED_FLOW_10 = &HFE1A
    PTPDO_ESTIMATED_FLOW_11 = &HFE1B
    PTPDO_ESTIMATED_FLOW_12 = &HFE1C
    PTPDO_ESTIMATED_FLOW_13 = &HFE1D
    PTPDO_ESTIMATED_FLOW_14 = &HFE1E
    PTPDO_ESTIMATED_FLOW_15 = &HFE1F
    PTPDO_ESTIMATED_FLOW_MAX
End Enum

' 0xFFFE - CPROTOCOL LISTEN PORT
' 0xFFFD - CPROTOCOL CONNECT PORT da qualsiasi source address

' Struttura dati di un messaggio PRPDO DI
Public Type PRPDO_DI
    Value(64) As Boolean             ' Valore dell'ingresso digitale
    TimeStamp As Long               ' Ultimi istante ricevuto del messaggio
End Type

' Struttura dati di un messaggio PRPDO DO
Public Type PTPDO_D0
    Value(64) As Boolean           ' Valore dell'uscita digitale
    OldValue(64) As Boolean        ' Valore OLD dell'uscita digitale
    TimeStamp As Long              ' Ultimi istante ricevuto del messaggio
    Time As Long                   ' Timer di invio del messaggio ciclico
    Quantity As Integer            ' Quantità delle uscite da trasmettere
End Type


Public Type PTPDO_AVC
    Value(2) As Double          ' Memoria libera associata all'indirizzo 0
    Time As Long             ' Ultimi istante ricevuto del messaggio
    TimeStamp As Long              ' Ultimi istante ricevuto del messaggio
    TimeRequest As Long
End Type



' Struttura dati di un messaggio PRPDO AI16
Public Type PRPDO_AI16
    Value(4) As Long              ' Valore dell'ingresso digitale
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type

' Struttura dati di un messaggio PRPDO AI8
Public Type PRPDO_AI8
    Value(8) As Integer              ' Valore dell'ingresso digitale
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type

' Struttura dati di un messaggio PRPDO AI8
Public Type PTPDO_AO8
    Value(8) As Integer              ' Valore dell'ingresso digitale
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type


' Struttura dati di un messaggio PRPDO AI8
Public Type PTPDO_AO16
    Value(4) As Long              ' Valore dell'ingresso digitale
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type

Public Type PTPDO_INTERNAL
    Value(100) As Double          ' Memoria libera associata all'indirizzo 0
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type


Public Type PTPDO_MLT_VALUE
    Value As Double             ' Memoria libera associata all'indirizzo 0
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type

Public Type PTPDO_MLT
    Dato(50) As PTPDO_MLT_VALUE          ' Memoria libera associata all'indirizzo 0
End Type

Public Type PTPDO_MLT_COMANDI_INTERNAL
    Value As Double                 ' variabili da inviare come comando il sistema invia un comando quando il valore è diverso da 0
    Richiesta As Boolean
End Type
    

Public Type PTPDO_MLT_COMANDI
    Comandi(65) As PTPDO_MLT_COMANDI_INTERNAL                 ' variabili da inviare come comando il sistema invia un comando quando il valore è diverso da 0
    PARAMETRI(70) As Double                                   ' Parametri letti
End Type




Public Type PTPDO_CE16
    Value(2) As Double          ' Memoria libera associata all'indirizzo 0
    TimeStamp As Long             ' Ultimi istante ricevuto del messaggio
End Type



' Questa struttura definisce l'area massimo
Public Type TagNodeMMS2218
    CE16(10) As PTPDO_CE16
    MLT(10) As PTPDO_MLT            ' Gestione dei messaggi provenienti dai moduli FDx
    MLT_COMANDI(10) As PTPDO_MLT_COMANDI
    AVC(10) As PTPDO_AVC            ' Gestione dei messaggi standard AVC
    INTERNAL(1) As PTPDO_INTERNAL
    DIN(1) As PRPDO_DI            ' Numero massimo di ingressi digitali presenti sulla MMS2218
    DOUT(1) As PTPDO_D0           ' Numero massimo di uscite digitali presenti sulla MMS2218
    AIN16(2) As PRPDO_AI16        ' Numero massimo di ingressi analogici presenti sulla MMS2218
    AIN8(0) As PRPDO_AI8          ' Numeromassimo di ingressi analogici a 8 bit presenti sulla MMS2218
    AOUT16(2) As PTPDO_AO16       ' Numero massimo di ingressi analogici presenti sulla MMS2218
    SA As Integer                 ' Identificativo della scheda sulla rete
End Type



' Decodifica i messaggi in ingresso
Public Sub DecodeMessage(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    Update_PRPDO_DI Context, Messaggio
    Update_PRPDO_AI16 Context, Messaggio
    Update_ESTIMATED_FLOW_FDx Context, Messaggio
    Update_PROPB01_FD8 Context, Messaggio
End Sub

' Questa funzione gestisce la trasmissione dei messaggi
Public Sub ManageTxMessage(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    Dim i As Integer
    Dim j As Integer
    Dim VarIndex As Integer
    i = 0
    Messaggio = ""
    Do While (i < UBound(Context.DOUT))
        VarIndex = 0
        Do While (VarIndex < UBound(Context.DOUT(i).Value))
            
            VarIndex = VarIndex + 1
            If (Context.DOUT(i).TimeStamp = 0) Then
                Update_PRPDO_DO Context, i, Messaggio
                Exit Sub
            Else
                j = 0
                Do While (j < UBound(Context.DOUT(i).Value))
                    If (Context.DOUT(i).Value(j) <> Context.DOUT(i).OldValue(j)) Then
                        Update_PRPDO_DO Context, i, Messaggio
                    End If
                    
                    j = j + 1
                Loop
                
                If (Context.DOUT(i).Time <> 0) Then
                    If (Context.DOUT(i).TimeStamp + Context.DOUT(i).Time < GetTickCount) Then
                         Update_PRPDO_DO Context, i, Messaggio
                        Exit Sub
                    End If
                End If
            End If
        Loop
        
        
        i = i + 1
    Loop


End Sub


Public Sub ManageTX_MLT_Comandi(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)

    Dim i As Integer
    Dim j As Integer
    Dim VarIndex As Integer
    i = 0
    Messaggio = ""
    Do While (i < UBound(Context.MLT_COMANDI))
        
        j = 0
        Do While (j < UBound(Context.MLT_COMANDI(i).Comandi))
        
            If (Context.MLT_COMANDI(i).Comandi(j).Richiesta = True) Then
                
                Context.MLT_COMANDI(i).Comandi(j).Richiesta = False
                Update_PRPDO_MLT_COMANDI Context, i, j, Messaggio
                
                
                Exit Sub
            End If
            
            j = j + 1
        Loop
                  
        i = i + 1
    Loop


End Sub


' Questa funzione gestisce la trasmissione dei messaggi
Public Sub ManageTxAVCMessage(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    Dim i As Integer
    Dim j As Integer
    Dim VarIndex As Integer
    i = 0
    Messaggio = ""
    Do While (i < UBound(Context.AVC))
        
        If (Context.AVC(i).TimeRequest + 2000 > GetTickCountEvo) Then
        
        
            If (Context.AVC(i).TimeStamp = 0) Then
                Update_PRPDO_AVC Context, i, Messaggio
                Context.AVC(i).TimeStamp = GetTickCountEvo
                Exit Sub
            ElseIf (Context.AVC(i).Time <> 0) Then
                If (Context.AVC(i).TimeStamp + Context.AVC(i).Time < GetTickCountEvo) Then
                     Update_PRPDO_AVC Context, i, Messaggio
                     
                     Context.AVC(i).TimeStamp = GetTickCountEvo
                    Exit Sub
                End If
            End If
        
        End If
        
        i = i + 1
    Loop


End Sub


' Questa funzione gestisce la trasmissione dei messaggi
Public Sub ManageTxAVCMessage_EVO(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    Dim i As Integer
    Dim j As Integer
    Dim VarIndex As Integer
    i = 0
    Messaggio = ""
    Do While (i < UBound(Context.AVC))
                                   
        If (Context.AVC(i).TimeStamp = 0) Then
            Update_PRPDO_AVC_EVO Context, i, Messaggio
            Exit Sub
        ElseIf (Context.AVC(i).Time <> 0) Then
            If (Context.AVC(i).TimeStamp + Context.AVC(i).Time < GetTickCount) Then
                 Update_PRPDO_AVC_EVO Context, i, Messaggio
                Exit Sub
            End If
        End If
          
        
        i = i + 1
    Loop


End Sub

Private Function Update_PRPDO_AVC_EVO(ByRef Context As TagNodeMMS2218, ByRef Index As Integer, ByRef Messaggio As String)
    Dim Quantity As Integer
    Dim i As Integer
    Dim PRPDO_PGN As String
    Dim J1939INDEX As String
    Dim ucByte As Integer
    Dim ucValue As Integer
    Dim bit As Integer
    Dim bSendData As Boolean
    Dim tempStr As String
    Dim PDU As String
    Dim Context_SA As String
    
    PRPDO_PGN = "FE3" + Hex(Index)
    
    If (Len(PRPDO_PGN) < 4) Then
        Do While (Len(PRPDO_PGN) < 4)
            PRPDO_PGN = "0" + PRPDO_PGN
        Loop
    ElseIf (Len(PRPDO_PGN) > 4) Then
        PRPDO_PGN = Right(PRPDO_PGN, 4)
    End If
    
    strPercentual = Hex(Context.AVC(Index).Value(0))
    
    If (Len(strPercentual) < 2) Then
        Do While (Len(strPercentual) < 2)
            strPercentual = "0" + strPercentual
        Loop
    End If
    
    Context_SA = Hex(Context.AVC(Index).Value(2))
    
    Do While (Len(Context_SA) < 2)
        Context_SA = "0" + Context_SA
    Loop
    
    strPercentual = Hex(Context.AVC(Index).Value(0))
    strState = Hex(Context.AVC(Index).Value(1))

        
    
    Messaggio = "T18" + PRPDO_PGN + Context_SA + "8" + strPercentual + "FF" + "F" & strState + "FFFFFFFFFF" + vbCr
    
End Function





Private Function Update_PRPDO_MLT_COMANDI(ByRef Context As TagNodeMMS2218, ByRef Index As Integer, ByRef Command As Integer, ByRef Messaggio As String)
    Dim Quantity As Integer
    Dim PRPDO_PGN As String
    Dim J1939INDEX As String
    Dim PDU As String
    Dim Dlc As String
    Dim tempStr As String
    Dim bSendCommand As Boolean
    Dim Context_SA As String
    
    
    bSendCommand = True
    
    Context_SA = Hex(Index)
    
    
    Select Case Command
    
    Case 0:
        '
        ' Enter configuration Mode
        '
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        PDU = "70" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
        
    
    Case 1:
        '
        ' Exit configuration Mode
        '
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        PDU = "7F" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    
    
    Case 2:
        '
        ' Store configuration Mode
        '
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        PDU = "76" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
            
    
    Case 3:
        '
        ' Calibrazione del sensore
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
                              
        PDU = "78" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    
    Case 4:
        '
        ' NORMAL MODE
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        
        
        tempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        
        Do While (Len(tempStr) < 2)
            tempStr = "0" + tempStr
        Loop
        
        
        PDU = "7C" + "00" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
        
    Case 5:
        '
        ' Modalità TEST
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        PDU = "79" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        

    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
        
    
    Case 6:
        '
        ' Modalità TEST POT
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        PDU = "7A" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        

    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
        
    
    
    Case 7:
        '
        ' Modalità DEBUG 1
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        'TempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        'TempStr = "1"
        
        'Do While (Len(TempStr) < 2)
            'TempStr = "0" + TempStr
        'Loop
        
        
        PDU = "7B" + "01" + "FF" + "FF" + "FF" + "FF" + "FF" + "FF"
        
    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
        
  
Case 8:
        '
        ' Modalità DEBUG 2
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        tempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        
        
        Do While (Len(tempStr) < 2)
            tempStr = "0" + tempStr
        Loop
        
        
        PDU = "7B" + "02" + tempStr + "FF" + "FF" + "FF" + "FF" + "FF"
        
    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
  
  
Case 9:
        '
        ' Modalità DEBUG 4
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        tempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        
        
        Do While (Len(tempStr) < 2)
            tempStr = "0" + tempStr
        Loop
        
        
        PDU = "7B" + "04" + tempStr + "FF" + "FF" + "FF" + "FF" + "FF"
        
    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
  
  
Case 10:
        '
        ' AutoApp Evo ZERO
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        tempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        
        
        Do While (Len(tempStr) < 2)
            tempStr = "0" + tempStr
        Loop
        
        
        PDU = "7C" + "00" + tempStr + "FF" + "FF" + "FF" + "FF" + "FF"
        
    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
  
    
    
Case 11:
        '
        ' AutoApp Evo FW
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        tempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        
        
        Do While (Len(tempStr) < 2)
            tempStr = "0" + tempStr
        Loop
        
        
        PDU = "7C" + "01" + tempStr + "FF" + "FF" + "FF" + "FF" + "FF"
        
    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
        
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
      
  
Case 12:
        '
        ' AutoApp Evo BW
        '
    
        PRPDO_PGN = "EF8" + Context_SA
        J1939INDEX = "T18" + PRPDO_PGN + "22"
        Dlc = "8"
        
        tempStr = Hex(Context.MLT_COMANDI(Index).Comandi(Command).Value)
        
        
        Do While (Len(tempStr) < 2)
            tempStr = "0" + tempStr
        Loop
        
        
        PDU = "7C" + "02" + tempStr + "FF" + "FF" + "FF" + "FF" + "FF"
        
    
        Messaggio = J1939INDEX + Dlc + PDU + Chr(13)
      
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
        
        
Case 20:
    If (ReadFd5_Merlo_I1(Index, Context.MLT_COMANDI(Index).PARAMETRI, 0) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 21:
    If (ReadFd5_Merlo_I2(Index, Context.MLT_COMANDI(Index).PARAMETRI, 3) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 22:
    If (ReadFd5_PID(Index, Context.MLT_COMANDI(Index).PARAMETRI, 6) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If

Case 23:
    If (ReadFd5_PIDR(Index, Context.MLT_COMANDI(Index).PARAMETRI, 9) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If

    
Case 24:
    If (ReadFd5_SPOOL(Index, Context.MLT_COMANDI(Index).PARAMETRI, 12) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 25:
    If (ReadFd5_OPMODE(Index, Context.MLT_COMANDI(Index).PARAMETRI, 18) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 26:
    If (ReadFd5_CorsaBW(Index, Context.MLT_COMANDI(Index).PARAMETRI, 19) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 27:
    If (ReadFd5_CorsaFW(Index, Context.MLT_COMANDI(Index).PARAMETRI, 20) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 28:
    If (ReadFd5_ALARMS(Index, Context.MLT_COMANDI(Index).PARAMETRI, 21) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 29:
    If (ReadFd5_PIN2(Index, Context.MLT_COMANDI(Index).PARAMETRI, 29) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 30:
    If (ReadFd5_evside(Index, Context.MLT_COMANDI(Index).PARAMETRI, 30) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 31:
    If (ReadFd5_SOURCEADDRESS(Index, Context.MLT_COMANDI(Index).PARAMETRI, 31) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 32:
    If (ReadFd5_LIMITATION(Index, Context.MLT_COMANDI(Index).PARAMETRI, 32) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 33:
    If (ReadFd5_flaotside(Index, Context.MLT_COMANDI(Index).PARAMETRI, 34) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 34:
    If (ReadFd5_I_R(Index, Context.MLT_COMANDI(Index).PARAMETRI, 35) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 35:
    If (ReadFd5_DITHER(Index, Context.MLT_COMANDI(Index).PARAMETRI, 37) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 36:
    If (ReadFd5_BITFIELD(Index, Context.MLT_COMANDI(Index).PARAMETRI, 38) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 37:
    If (ReadFd5_RAMP(Index, Context.MLT_COMANDI(Index).PARAMETRI, 39) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If

Case 38:
    If (ReadFd5_STEPI(Index, Context.MLT_COMANDI(Index).PARAMETRI, 41) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 39:
    If (ReadFd5_STEP_KD(Index, Context.MLT_COMANDI(Index).PARAMETRI, 51) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 40:
    If (ReadFd5_STEP_KI(Index, Context.MLT_COMANDI(Index).PARAMETRI, 59) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
    
    
    
    
    
    
    
Case 41:
    If (WriteFd5_Merlo_I1_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 0) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 42:
    If (WriteFd5_Merlo_I2_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 3) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 43:
    If (WriteFd5_PID_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 6) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If

Case 44:
    If (WriteFd5_PIDR_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 9) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If

    
Case 45:
    If (WriteFd5_SPOOL_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 12) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 46:
    If (WriteFd5_OPMODE_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 18) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 47:
    If (WriteFd5_CorsaBW_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 19) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 48:
    If (WriteFd5_CorsaFW_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 20) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 49:
    If (WriteFd5_ALARMS_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 21) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 50:
    If (WriteFd5_PIN2_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 29) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 51:
    If (WriteFd5_evside_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 30) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 52:
    If (WriteFd5_SOURCEADDRESS_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 31) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 53:
    If (WriteFd5_LIMITATION_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 32) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 54:
    If (WriteFd5_flaotside_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 34) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 55:
    If (WriteFd5_I_R_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 35) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 56:
    If (WriteFd5_DITHER_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 37) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 57:
    If (WriteFd5_BITFIELD_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 38) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
Case 58:
    If (WriteFd5_RAMP_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 39) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If

Case 59:
    If (WriteFd5_STEPI_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 41) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 60:
    If (WriteFd5_STEP_KD_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 51) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
Case 61:
    If (WriteFd5_STEP_KI_EVO(Index, Context.MLT_COMANDI(Index).PARAMETRI, 59) = True) Then
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 0
    Else
        Context.MLT_COMANDI(Index).Comandi(Command).Value = 2
    End If
    
    
    
    
    
    
    End Select
    

    
End Function










Private Function Update_PRPDO_AVC(ByRef Context As TagNodeMMS2218, ByRef Index As Integer, ByRef Messaggio As String)
    Dim Quantity As Integer
    Dim i As Integer
    Dim PRPDO_PGN As String
    Dim J1939INDEX As String
    Dim ucByte As Integer
    Dim ucValue As Integer
    Dim bit As Integer
    Dim bSendData As Boolean
    Dim tempStr As String
    Dim PDU As String
    Dim Context_SA As String
    
    PRPDO_PGN = "FE3" + Hex(Index)
    
    If (Len(PRPDO_PGN) < 4) Then
        Do While (Len(PRPDO_PGN) < 4)
            PRPDO_PGN = "0" + PRPDO_PGN
        Loop
    ElseIf (Len(PRPDO_PGN) > 4) Then
        PRPDO_PGN = Right(PRPDO_PGN, 4)
    End If
    
    strPercentual = Hex(Context.AVC(Index).Value(0))
    
    If (Len(strPercentual) < 2) Then
        Do While (Len(strPercentual) < 2)
            strPercentual = "0" + strPercentual
        Loop
    End If
    
    
    Context_SA = Hex(Context.AVC(Index).Value(2))
    If (Len(Context_SA) < 2) Then
        Do While (Len(Context_SA) < 2)
            Context_SA = "0" + Context_SA
        Loop
    End If
    
    
   ' strPercentual = Hex(Context.AVC(Index).Value(0))
    strState = Hex(Context.AVC(Index).Value(1))

        
    
    Messaggio = "T18" + PRPDO_PGN + Context_SA + "8" + strPercentual + "FF" + "F" & strState + "FFFFFFFFFF" + vbCr
    
    'Context.DOUT(Index).TimeStamp = GetTickCount
End Function










' Questa funzione imposta il timer di trasmissione ciclico del messaggi di trasmissione
Public Sub Configure_Tx_AVC_Message(ByRef Context As TagNodeMMS2218, Index As Integer, Time As Integer)
    Context.AVC(Index).Time = Time
End Sub


' Questa funzione imposta il timer di trasmissione ciclico del messaggi di trasmissione
Public Sub Configure_Tx_DO_Message(ByRef Context As TagNodeMMS2218, Index As Integer, Time As Integer, Quantity As Integer)
    Context.DOUT(Index).Time = Time
    Context.DOUT(Index).Quantity = Quantity
End Sub


' Questa funzione inposta il Source address per la comunicazione con
Public Sub SetSourceAddress(ByRef Context As TagNodeMMS2218, RemoteSA As Integer)
    Context.SA = RemoteSA
End Sub


' Questa funzione decodifica il messaggio CAN ricevuto,se proviene dal CE16 aggiorna i valori della struttura
' dati
Private Sub Update_PRPDO_DI(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    
    Dim PRPDO_PGN As ePRPDO_DI_PGN
    
    Dim SA As String
    Dim PDU As String
    Dim J1939INDEX As String

    Dim bExit As Boolean
    
    Dim i As Integer
    Dim Index As Integer

    ' Mi creo la stringa del SA ( comune a tutti i messaggi in ricezione dalla scheda
    SA = Hex(Context.SA)
    Do While (Len(SA) < 2)
        SA = "0" + SA
    Loop
    
    PRPDO_PGN = PRPDO_DI_0               ' Primo indice da decodificare
    i = 0
    
    If (UBound(Context.DIN) = 0) Then       ' controllo se ho almeno un
        bExit = True
    End If
        
    ' Scansiono tutti i messaggi possibili che la scheda mi può inviare relativa agli ingressi digitali
    Do While (bExit = False)
            
        
        If (Len(Hex(PRPDO_PGN)) > 4) Then
            J1939INDEX = Right(Hex(PRPDO_PGN), Len(Hex(PRPDO_PGN)) - 4)
        Else
            J1939INDEX = Hex(PRPDO_PGN)
        End If
        
        
        
        Index = 0
        Do While (Len(J1939INDEX) < 4)
            J1939INDEX = "0" + J1939INDEX
        Loop
            
        J1939INDEX = J1939INDEX + SA                                ' Creo l'indice da ricercare nel messaggio ricevuto
                
        PDU = GetPDU(Messaggio, J1939INDEX)
                
        If (PDU <> "") Then                                     ' String atrovata
            DecodeDigital PDU, Context.DIN(Index).Value             ' Mi salvo i dati
            Context.DIN(Index).TimeStamp = GetTickCount             ' Mi salvo l'instante ricevuto
            
            Index = Index + 1
            PRPDO_PGN = PRPDO_PGN + 1
        
            If (PRPDO_PGN >= PRPDO_DI_MAX) Then                 ' controllo il PGN del messaggio PRPDO_DI_PGN associato ai DI
                bExit = True                                    ' Finito gli indici del vettore
            End If
        
            i = i + 1
        
            If (i >= UBound(Context.DIN)) Then                   ' Controllo la dimensione massima degli indici del vettore della scheda
                bExit = True                                    ' Finito gli indici del vettore
            End If
        Else
            bExit = True
        End If
    Loop
        
End Sub


Private Sub Update_PRPDO_DO(ByRef Context As TagNodeMMS2218, ByRef Index As Integer, ByRef Messaggio As String)
    Dim Quantity As Integer
    Dim i As Integer
    Dim PRPDO_PGN As String
    Dim J1939INDEX As String
    Dim ucByte As Integer
    Dim ucValue As Integer
    Dim bit As Integer
    Dim bSendData As Boolean
    Dim tempStr As String
    Dim PDU As String
    Quantity = Context.DOUT(Index).Quantity
    
    PRPDO_PGN = (Hex(PTPDO_DO_0 + Index))
    
    If (Len(PRPDO_PGN) < 4) Then
        Do While (Len(PRPDO_PGN) < 4)
            PRPDO_PGN = "0" + PRPDO_PGN
        Loop
    ElseIf (Len(PRPDO_PGN) > 4) Then
        PRPDO_PGN = Right(PRPDO_PGN, 4)
    End If
    
    J1939INDEX = "T18" + PRPDO_PGN + Hex(Context.SA)
    
    
    
    
    Do While (i < Quantity)
        If (bit >= 8) Then
            bit = 0
            
            tempStr = Hex(ucValue)
            ucValue = 0
            If (Len(tempStr) < 2) Then
            Do While (Len(tempStr) < 2)
                tempStr = "0" + tempStr
            Loop
            ElseIf (Len(tempStr) > 2) Then
                tempStr = Right(tempStr, 2)
            End If
    
            PDU = PDU + tempStr
            
            ucByte = ucByte + 1
        End If
        
        bSendData = True
        If (Context.DOUT(Index).Value(i) = True) Then
            ucValue = ucValue + 2 ^ bit
        End If
        
        Context.DOUT(Index).OldValue(i) = Context.DOUT(Index).Value(i)
        bit = bit + 1
        i = i + 1
    Loop
    
    If (bSendData = True) Then
        ucByte = ucByte + 1
        tempStr = Hex(ucValue)
        If (Len(tempStr) < 2) Then
            Do While (Len(tempStr) < 2)
                tempStr = "0" + tempStr
            Loop
        ElseIf (Len(tempStr) > 2) Then
            tempStr = Right(tempStr, 2)
        End If

        PDU = PDU + tempStr
    
    End If
    Messaggio = J1939INDEX + CStr(ucByte) + PDU + Chr(13)
    Context.DOUT(idnex).TimeStamp = GetTickCount
End Sub


' Funzione di decodifica del messaggio AI16
Private Sub Update_PRPDO_AI16(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    
    Dim PRPDO_PGN As ePRPDO_AI16_PGN
    Dim i As Integer
    Dim SA As String
    Dim PDU As String
    Dim J1939INDEX As String
    Dim bExit As Boolean
    Dim Index As Integer
    
    ' Mi creo la stringa del SA ( comune a tutti i messaggi in ricezione dalla scheda
    SA = Hex(Context.SA)
    Do While (Len(SA) < 2)
        SA = "0" + SA
    Loop
    
    PRPDO_PGN = PRPDO_AI16_0                ' Primo indice da decodificare
    i = 0
    
    If (UBound(Context.AIN16) = 0) Then        ' controllo se ho almeno un
        bExit = True
    End If
        
    ' Scansiono tutti i messaggi possibili che la scheda mi può inviare relativa agli ingressi digitali
    Do While (bExit = False)
        
        If (Len(Hex(PRPDO_PGN)) > 4) Then
            J1939INDEX = Right(Hex(PRPDO_PGN), Len(Hex(PRPDO_PGN)) - 4)
        Else
            J1939INDEX = Hex(PRPDO_PGN)
        End If
        
        Do While (Len(J1939INDEX) < 4)
            J1939INDEX = "0" + J1939INDEX
        Loop
            
        J1939INDEX = J1939INDEX + SA                                ' Creo l'indice da ricercare nel messaggio ricevuto
        
        PDU = GetPDU(Messaggio, J1939INDEX)
                    
        If (PDU <> "") Then
            DecodeAnalog16 PDU, Context.AIN16(Index).Value           ' Mi salvo i dati
            Context.AIN16(Index).TimeStamp = GetTickCount            ' Mi salvo l'instante ricevuto
            bExit = True
            
        End If
        
        PRPDO_PGN = PRPDO_PGN + 1
        Index = Index + 1
        If (Index >= UBound(Context.AIN16)) Then
            bExit = True
        End If
    Loop
        
End Sub

' Funzione di decodifica del messaggio AI8
Private Sub Update_PRPDO_AI8(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    
    Dim PRPDO_PGN As ePRPDO_AI8_PGN
    
    Dim SA As String
    Dim PDU As String
    Dim J1939INDEX As String
    Dim bExit As Boolean
    Dim Index As Integer
    
    ' Mi creo la stringa del SA ( comune a tutti i messaggi in ricezione dalla scheda
    SA = Hex(Context.SA)
    Do While (Len(SA) < 2)
        SA = "0" + SA
    Loop
    
    PRPDO_PGN = PRPDO_AI8_0                ' Primo indice da decodificare
    i = 0
    
    If (UBound(Context.AIN8) = 0) Then        ' controllo se ho almeno un
        bExit = True
    End If
        
    ' Scansiono tutti i messaggi possibili che la scheda mi può inviare relativa agli ingressi digitali
    Do While (bExit = False)
            
        J1939INDEX = Hex(PRPDO_DI_PGN)
        Index = 0
        Do While (Len(J1939INDEX) < 4)
            J1939INDEX = "0" + J1939INDEX
        Loop
            
        J1939INDEX = J1939INDEX + SA                                ' Creo l'indice da ricercare nel messaggio ricevuto
        
        PDU = GetPDU(Messaggio, J1939INDEX)
                    
        If (PDU <> "") Then
            DecodeAnalog8 PDU, Context.AIN8(Index).Value           ' Mi salvo i dati
            Context.AIN8(Index).TimeStamp = GetTickCount            ' Mi salvo l'instante ricevuto
            
            Index = Index + 1
            PRPDO_PGN = PRPDO_PGN + 1
        
            If (PRPDO_PGN >= PRPDO_AI8_MAX) Then                 ' controllo il PGN del messaggio PRPDO_DI_PGN associato ai DI
                bExit = True                                    ' Finito gli indici del vettore
            End If
        
            i = i + 1
        
            If (i >= UBound(Context.AIN8)) Then                   ' Controllo la dimensione massima degli indici del vettore della scheda
                bExit = True                                    ' Finito gli indici del vettore
            End If
        End If
    Loop
End Sub


Public Sub UpdateDato(ByRef Dato As PTPDO_MLT_VALUE, ByVal Value As Double)
    Dato.Value = Value
    Dato.TimeStamp = GetTickCountEvo
End Sub


Private Sub Update_PROPB01_FD8(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
   
    
    Dim SA As String
    Dim PDU As String
    Dim J1939INDEX As String
    Dim bExit As Boolean
    Dim Index As Integer
    Dim Indice As Long
    
    Dim Debug_module As Integer
    
    i = 0
    
        
    ' Scansiono tutti i messaggi possibili che la scheda mi può inviare relativa agli ingressi digitali
    Do While (bExit = False)
        
        J1939INDEX = "18FF018" + Hex(i)
                
        Index = 0

               
        
        PDU = GetPDU(Messaggio, J1939INDEX)
                    
        If (PDU <> "") Then
            
            Debug_module = DecodeSingleValue(PDU, 24, 8, False, False)
            
            Select Case (Debug_module)
            
            Case 2          ' Normal operation
                UpdateDato Context.MLT(i).Dato(30), DecodeSingleValue(PDU, 0, 8, True, False)           ' Error.p
                UpdateDato Context.MLT(i).Dato(31), DecodeSingleValue(PDU, 8, 8, True, False)           ' Error.d
                UpdateDato Context.MLT(i).Dato(32), DecodeSingleValue(PDU, 16, 8, True, False)          ' correzione_totale
                UpdateDato Context.MLT(i).Dato(33), DecodeSingleValue(PDU, 32, 8, False, False)         ' ILaminazione[PWM_E]
                UpdateDato Context.MLT(i).Dato(34), DecodeSingleValue(PDU, 40, 8, False, False)         ' ILaminazione[PWM_R]
                UpdateDato Context.MLT(i).Dato(35), DecodeSingleValue(PDU, 48, 8, False, False)          ' iLimitImax
                UpdateDato Context.MLT(i).Dato(36), DecodeSingleValue(PDU, 56 + 7 - 6, 1, False, False)  ' reg_fine
                UpdateDato Context.MLT(i).Dato(37), DecodeSingleValue(PDU, 56 + 7 - 5, 1, False, False)  ' reg_der
                UpdateDato Context.MLT(i).Dato(38), DecodeSingleValue(PDU, 56 + 7 - 1, 1, False, False)  ' bContro_molla
                UpdateDato Context.MLT(i).Dato(39), DecodeSingleValue(PDU, 56 + 7 - 0, 1, False, False)  ' neutral
                UpdateDato Context.MLT(i).Dato(40), DecodeSingleValue(PDU, 56 + 7 - 2, 1, False, False)  ' Rientro_neutro
                UpdateDato Context.MLT(i).Dato(41), DecodeSingleValue(PDU, 56 + 7 - 7, 1, False, False)  ' full_speed
                            
                            
            Case 1:
               
                            
            Case 2:
            
                UpdateDato Context.MLT(i).Dato(10), DecodeSingleValue(PDU, 0, 8, False, True)      ' PwmE
                UpdateDato Context.MLT(i).Dato(11), DecodeSingleValue(PDU, 8, 8, False, True)      ' PwmR
                UpdateDato Context.MLT(i).Dato(12), DecodeSingleValue(PDU, 24, 16, False, True)      ' pos_int
                UpdateDato Context.MLT(i).Dato(13), DecodeSingleValue(PDU, 40, 16, False, True)      ' setpoint_int
                
                
                Indice = DecodeSingleValue(PDU, 16 + 2, 5, False, False)  ' pIndiceDato
                Indice = DecodeSingleValue(PDU, 16 + 2, 6, False, False)  ' pIndiceDato
                
                Select Case Indice
                    Case 1: UpdateDato Context.MLT(i).Dato(14), DecodeSingleValue(PDU, 56, 8, False, False)      ' TRANGE
                    Case 2: UpdateDato Context.MLT(i).Dato(15), DecodeSingleValue(PDU, 56, 8, False, False) * 100 + 8000     ' VBAT
                    Case 3: UpdateDato Context.MLT(i).Dato(16), DecodeSingleValue(PDU, 56, 8, False, False)      ' FAULTID
                    Case 4: UpdateDato Context.MLT(i).Dato(17), DecodeSingleValue(PDU, 56, 8, True, False)      ' TEMPERATURE
                    Case 5: UpdateDato Context.MLT(i).Dato(18), DecodeSingleValue(PDU, 56, 8, False, False)      ' STEP_INTEGRALE
                    
                End Select
            

            End Select
            
            bExit = True                                    ' Finito gli indici del vettore
        
        Else
            i = i + 1
            If (i > 15) Then
                bExit = True
            End If
        End If
    Loop
End Sub


' Funzione di decodifica del messaggio AI8
Private Sub Update_ESTIMATED_FLOW_FDx(ByRef Context As TagNodeMMS2218, ByRef Messaggio As String)
    
    Dim PRPDO_PGN As ePTPDO_ESTIMATED_FLOW_PGN
    
    Dim SA As String
    Dim PDU As String
    Dim J1939INDEX As String
    Dim bExit As Boolean
    Dim Index As Integer
    Dim Indice As Long
    
    Dim Debug_module As Integer
    
    'PRPDO_PGN = PTPDO_ESTIMATED_FLOW_00                ' Primo indice da decodificare
    i = 0
    
        
    ' Scansiono tutti i messaggi possibili che la scheda mi può inviare relativa agli ingressi digitali
    Do While (bExit = False)
        
        J1939INDEX = "18FE1" + Hex(i) + "8" + Hex(i)
                
        Index = 0
        Do While (Len(J1939INDEX) < 6)
            J1939INDEX = "0" + J1939INDEX
        Loop
                    
        
        
        PDU = GetPDU(Messaggio, J1939INDEX)
                    
        If (PDU <> "") Then
            
            Debug_module = DecodeSingleValue(PDU, 16, 2, False, False)
            
            Select Case (Debug_module)
            
            Case 0          ' Normal operation
                UpdateDato Context.MLT(i).Dato(0), DecodeSingleValue(PDU, 0, 8, False, False)      ' EstExtendFlow
                UpdateDato Context.MLT(i).Dato(1), DecodeSingleValue(PDU, 8, 8, False, False)      ' EstExtendFlow
                UpdateDato Context.MLT(i).Dato(2), DecodeSingleValue(PDU, 16, 6, False, False)      ' valve_action
                UpdateDato Context.MLT(i).Dato(3), DecodeSingleValue(PDU, 24, 16, False, True)      ' pos_int
                UpdateDato Context.MLT(i).Dato(4), DecodeSingleValue(PDU, 40, 16, False, True)      ' setpoint_int
                UpdateDato Context.MLT(i).Dato(5), DecodeSingleValue(PDU, 56, 8, False, False)      ' FaultId
                            
                            
            Case 1:
                UpdateDato Context.MLT(i).Dato(20), DecodeSingleValue(PDU, 0, 8, False, False)      ' PWMdutyNom[PWM_E]
                UpdateDato Context.MLT(i).Dato(21), DecodeSingleValue(PDU, 8, 8, False, False)      ' PWMdutyNom[PWM_R]
                UpdateDato Context.MLT(i).Dato(22), DecodeSingleValue(PDU, 24, 16, False, True)      ' hs1f
                UpdateDato Context.MLT(i).Dato(23), DecodeSingleValue(PDU, 40, 16, False, True)      ' hs2f
                UpdateDato Context.MLT(i).Dato(24), DecodeSingleValue(PDU, 56, 8, True, False)      ' Trange
                            
                            
                            
            Case 2:
            
                UpdateDato Context.MLT(i).Dato(10), DecodeSingleValue(PDU, 0, 8, False, True)      ' PwmE
                UpdateDato Context.MLT(i).Dato(11), DecodeSingleValue(PDU, 8, 8, False, True)      ' PwmR
                UpdateDato Context.MLT(i).Dato(12), DecodeSingleValue(PDU, 24, 16, False, True)      ' pos_int
                UpdateDato Context.MLT(i).Dato(13), DecodeSingleValue(PDU, 40, 16, False, True)      ' setpoint_int
                
                
                'indice = DecodeSingleValue(PDU, 16 + 2, 5, False, False)  ' pIndiceDato
                Indice = DecodeSingleValue(PDU, 16 + 2, 6, False, False)  ' pIndiceDato
                
                Select Case Indice
                    Case 1: UpdateDato Context.MLT(i).Dato(14), DecodeSingleValue(PDU, 56, 8, True, False)      ' TRANGE
                    Case 2: UpdateDato Context.MLT(i).Dato(15), DecodeSingleValue(PDU, 56, 8, False, False) * 100 + 8000     ' VBAT
                    Case 3: UpdateDato Context.MLT(i).Dato(16), DecodeSingleValue(PDU, 56, 8, False, False)      ' FAULTID
                    Case 4: UpdateDato Context.MLT(i).Dato(17), DecodeSingleValue(PDU, 56, 8, True, False)      ' TEMPERATURE
                    Case 5: UpdateDato Context.MLT(i).Dato(18), DecodeSingleValue(PDU, 56, 8, True, False)      ' STEP_INTEGRALE
                    
                End Select
            

            End Select
            
            bExit = True                                    ' Finito gli indici del vettore
        
        Else
            i = i + 1
            If (i > 15) Then
                bExit = True
            End If
        End If
    Loop
End Sub







' Questa funzione converte i campi di una PDU in valore digitali
Private Function DecodeDigital(ByRef PDU As String, ByRef Digital() As Boolean)
    Dim i As Integer
    Dim Y As Integer
    Dim DinIndex As Integer
    Dim Length As Integer
    Dim Value As Integer
    Dim PDU_BYTE As String
    On Error GoTo err:
    
    
    
    Length = Len(PDU)
    
    
    
    Do While (i < Length)
        PDU_BYTE = Mid(PDU, i + 1, 2)       ' Prelevo 2 caratteri dalla PDU
        Value = CInt("&H" + PDU_BYTE)            ' lo converto in un numero binario
        
        Y = 0                               ' converto il numero binario in un vettore boleano
        Do While (Y < 8)
        
            If (Value Mod 2) Then
                Digital(DinIndex) = True
                Value = Value - 1
            Else
                Digital(DinIndex) = False
            End If
            
            Value = Value / 2
            DinIndex = DinIndex + 1         ' mi sposto nel vettore dei dati booleani
            'j = j + 1                       ' mi sposto nel bit del byre
            Y = Y + 1
        Loop
        
        i = i + 2                           ' mi sposto di 2 caratteri nella pdu
    Loop
    
    Exit Function
err:
    Debug.Print "DecodeDigital = " + err.Description
End Function

' Questa funzione converte i campi di una PDU in valore digitali
Private Function DecodeAnalog16(ByRef PDU As String, ByRef Analog() As Long)
    Dim i As Integer
    Dim Y As Integer
    Dim Index As Integer
    Dim Length As Integer
    Dim Value As Integer
    Dim PDU_WORD As String
    
    On Error GoTo err:
    
    Length = Len(PDU)
    
    Do While (i < Length)
        PDU_WORD = Mid(PDU, i + 1, 4)       ' Prelevo 4 caratteri dalla PDU
        Value = CLng("&H" + PDU_WORD)       ' lo converto in un numero binario
        Analog(Index) = Value               ' Mi salvo il dato
        Index = Index + 1                   ' mi sposto nel vettore dei dati booleani
        i = i + 4                           ' mi sposto di 4 caratteri nella pdu
    Loop
    
    Exit Function
err:
    Debug.Print "DecodeDigital = " + err.Description
End Function

' Questa funzione converte i campi di una PDU in valore digitali
Private Function DecodeAnalog8(ByRef PDU As String, ByRef Analog() As Integer)
    Dim i As Integer
    Dim Y As Integer
    Dim Index As Integer
    Dim Length As Integer
    Dim Value As Integer
    Dim PDU_BYTE As String
    Length = Len(PDU)
    
    Do While (i < Length)
        PDU_BYTE = Mid(PDU, i + 1, 2)       ' Prelevo 4 caratteri dalla PDU
        Value = CInt(PDU_BYTE)              ' lo converto in un numero binario
        Analog(Index) = Value               ' Mi salvo il dato
        Index = Index + 1                   ' mi sposto nel vettore dei dati booleani
        i = i + 2                           ' mi sposto di 4 caratteri nella pdu
    Loop
End Function

Public Function HexToBin(ByVal hexString As String) As String
    Dim i As Integer, binString As String
    
    For i = 1 To Len(hexString)
        Select Case Mid(hexString, i, 1)
            Case "0"
                binString = binString & "0000"
            Case "1"
                binString = binString & "0001"
            Case "2"
                binString = binString & "0010"
            Case "3"
                binString = binString & "0011"
            Case "4"
                binString = binString & "0100"
            Case "5"
                binString = binString & "0101"
            Case "6"
                binString = binString & "0110"
            Case "7"
                binString = binString & "0111"
            Case "8"
                binString = binString & "1000"
            Case "9"
                binString = binString & "1001"
            Case "A", "a"
                binString = binString & "1010"
            Case "B", "b"
                binString = binString & "1011"
            Case "C", "c"
                binString = binString & "1100"
            Case "D", "d"
                binString = binString & "1101"
            Case "E", "e"
                binString = binString & "1110"
            Case "F", "f"
                binString = binString & "1111"
        End Select
    Next i
    
    HexToBin = binString
End Function


Public Function Bin2Dec(ByVal NUMERO As String) As Long
    Dim Temp As Integer
    For Temp = 1 To Len(NUMERO)
        Bin2Dec = Bin2Dec + (Mid(NUMERO, Temp, 1) * 2) ^ (Len(NUMERO) - Temp)
    Next Temp
    If Right(NUMERO, 1) = "0" Then Bin2Dec = Bin2Dec - 1
End Function




Function InvertByteOrder(binaryString As String) As String
    Dim byteCount As Integer
    Dim byteStart As Integer
    Dim byteLength As Integer
    Dim byteArray() As String
    
    ' Controlla la lunghezza della stringa in input
    Do While (Len(binaryString) Mod 8 <> 0)
        binaryString = "0" + binaryString

    Loop
    
    ' Divide la stringa in byte di 8 bit
    byteCount = Len(binaryString) \ 8
    ReDim byteArray(byteCount - 1) ' Inizializza l'array
    For i = 0 To byteCount - 1
        byteStart = i * 8 + 1
        byteLength = 8
        byteArray(i) = Mid(binaryString, byteStart, byteLength)
    Next i
    
    ' Inverte l'ordine degli elementi dell'array
    For i = 0 To byteCount \ 2 - 1
        Temp = byteArray(i)
        byteArray(i) = byteArray(byteCount - i - 1)
        byteArray(byteCount - i - 1) = Temp
    Next i
    
    ' Unisce gli elementi dell'array in una stringa
    InvertByteOrder = Join(byteArray, "")
End Function




' Questa funzione converte i campi di una PDU in valore digitali
Private Function DecodeSingleValue(ByRef PDU As String, ByVal BitStart As Integer, ByVal BitLength As Integer, ByVal Negativo As Boolean, ByVal BigEndian As Boolean) As Double
    Dim i As Integer
    
    Dim Index As Integer
    Dim binaryString As String
    Dim negatedString As String
    Dim Value As Long
    Dim PDU_VALUE As String
    Dim PDU_BINARY As String
    
    PDU_VALUE = "&h" + PDU
    
    PDU_BINARY = HexToBin(PDU)
    
    Length = Len(PDU_BINARY)
    
    
    If (Length >= BitStart + BitLength) Then
    
        binaryString = Mid(PDU_BINARY, BitStart + 1, BitLength)       ' Prelevo 4 caratteri dalla PDU
        
        Do While (Len(binaryString) < 8 * byteLength)
            binaryString = "0" + binaryString
        Loop
        
        If (BigEndian = True) Then
            binaryString = InvertByteOrder(binaryString)
        End If
        
        If (Negativo = True) Then
            If (Left(binaryString, 1) = "1") Then
                        
                For i = 1 To Len(binaryString)
                    If Mid(binaryString, i, 1) = "0" Then
                        negatedString = negatedString & "1"
                    Else
                        negatedString = negatedString & "0"
                    End If
                Next i
                
                Value = -Bin2Dec(negatedString) + 1
                               
                
            Else
                 Value = Bin2Dec(binaryString)
            End If
        Else
        
            Value = Bin2Dec(binaryString)
        
        End If
    End If
    DecodeSingleValue = Value
End Function

' Questa funzione ricerca l'index all'interno del messaggio ricevuto e ritorna il campo PDU
' Se non ci fosse corrispondeza ritorna una stringa nulla
Private Function GetPDU(ByRef Messaggio As String, J1939INDEX As String) As String
On Error GoTo err:
    Dim Dlc As String
    Dim Start As Long
    Dim Length As Long
    Dim FindIndex As Long
    
    FindIndex = InStr(1, Messaggio, J1939INDEX)                 ' Cerco la stringa nel messaggio
    
    If (FindIndex > 0) Then                                     ' String atrovata
    
        Start = FindIndex + Len(J1939INDEX)
        Length = 1
        Dlc = Mid(Messaggio, Start, Length)                     ' Prelevo il campo dlc
    
        Start = Start + 1
        Length = Len(Messaggio) - Start
        'PDU = Mid(Messaggio, Start, Length)                      ' Prelevo la PDU
        PDU = Mid(Messaggio, Start, Dlc * 2 + 1)                      ' Prelevo la PDU
        
        PDU = Replace(PDU, Chr(13), "")
    Else
        PDU = ""
    End If
    
    GetPDU = PDU
err:
End Function




