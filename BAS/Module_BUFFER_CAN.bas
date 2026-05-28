Attribute VB_Name = "Module_BUFFER_CAN"

Enum eRXErrorDriver
    eOK
    eError
End Enum

Public Enum eDriverRxMachineState
    eSearchT
    eSearchCr
End Enum

'mettere un Id per ogni processo/funzione concorrente che usa il CAN (adesso c'è "eCapture" come Id per la funzione "CaptureMlt()")
'(NOTA: processi tra di loro non concorrenti possono avere uno stesso Id)
Public Enum eTaskRxDriverId
    eCaptureTask
    ePropA
    eCE16               ' Acquisizione messaggi CE16
    e2218               ' Acquisizione messaggi MMS2218
    eSezione1
    eSezione2
    eSezione3
    eSezione4
    eSezione5
    Calib1
    Calib2
    Calib3
    Calib4
    Calib5
   ' eChangeNodeId
    'eStoreNodeId
    eMaxTaskRxDriver
End Enum


Const MAX_FRAME_RX_BUFFER = 500


Dim CirleBufferCan(MAX_FRAME_RX_BUFFER) As String
Dim bOverrun(MAX_FRAME_RX_BUFFER) As Boolean
Dim StartCirleBufferCan As Integer
Dim EndCirleBufferCan(eMaxTaskRxDriver) As Integer
Dim DriverRxStr As String
Dim DriverRxMachineState As eDriverRxMachineState




Public Sub DriverCanUsbRxBuffer(ByRef str As String)

Dim i As Integer
Dim TempStr As String

i = 0
Do While (i < Len(str))
    TempStr = Mid(str, 1 + i, 1)                        ' prelevo un carattere alla volta
    
    If (TempStr = "T") Then                             ' controllo lo start frame in questo caso T
        DriverRxMachineState = eSearchCr                ' Cambio lo stato della macchina, devo cercare la fine del pacchetto ora
        DriverRxStr = "T"                               ' mi salvo lo start frame nel buffer temporaneo
    Else
        If (DriverRxMachineState = eSearchCr) Then      ' ho trovato lo start frame, devo aspettare il cr
            If (TempStr = Chr(13)) Then                 ' il carattere selezionato è il cr ?
                DriverRxStr = DriverRxStr + Chr(13)     ' salvo il cr nel buffer temporaneo
                DriverRxMachineState = eSearchT         ' cambio lo stato macchina della ricezione dei messaggi
                UpdateLowerRxBuffer DriverRxStr         ' salvo il pacchetto ricevuto nel buffer circolare
            Else
                DriverRxStr = DriverRxStr + TempStr     ' ho ricevuto lo start frame,ma non l'end frame... continuo a salvare i caratteri nel buffer temporaneo
            End If
        End If
    End If
    i = i + 1                                           ' punto al prossimo carattere
Loop

End Sub


Public Sub UpdateLowerRxBuffer(ByVal DriverRxStr As String)
    Dim i As Integer
    CirleBufferCan(StartCirleBufferCan) = DriverRxStr
    
    StartCirleBufferCan = StartCirleBufferCan + 1
    
    If (StartCirleBufferCan >= UBound(CirleBufferCan)) Then
        StartCirleBufferCan = 0
    End If
    
    Do While (i < eMaxTaskRxDriver)
        If (EndCirleBufferCan(i) = StartCirleBufferCan) Then
            bOverrun(i) = True
            
            If (i = eCE16) Then
                Debug.Assert True
            End If
        End If
        i = i + 1
    Loop
End Sub

' questa funzione preleva un frame dal buffer circolare in funzione dell'ide del task selezionato.
' ogni task ha un indice proprio di EndCirleBufferCan, questo permette al task di avere sempre a disposizione
' tutti i frame ricevuti nonostante sia in concorrenza con gli altri task sulla ricezione dei messaggi
Public Function GetFrameFromRxBuffer(ByVal TaskId As Integer, ByRef RxBuffer As String) As eRXErrorDriver

If (TaskId >= UBound(EndCirleBufferCan)) Then
        GetFrameFromRxBuffer = eError
Else
    If (EndCirleBufferCan(TaskId) <> StartCirleBufferCan) Then
        RxBuffer = CirleBufferCan(EndCirleBufferCan(TaskId))
        EndCirleBufferCan(TaskId) = EndCirleBufferCan(TaskId) + 1
        If (EndCirleBufferCan(TaskId) >= UBound(CirleBufferCan)) Then
            EndCirleBufferCan(TaskId) = 0
        End If
        
        GetFrameFromRxBuffer = eOK
    Else
        GetFrameFromRxBuffer = eError
    End If
End If

End Function

Public Function FlushRxBuffer(ByVal TaskId As Integer) As eRXErrorDriver
    If (TaskId >= UBound(EndCirleBufferCan)) Then
            FlushRxBuffer = eError
    Else
        EndCirleBufferCan(TaskId) = StartCirleBufferCan
        FlushRxBuffer = eOK
    End If
End Function
