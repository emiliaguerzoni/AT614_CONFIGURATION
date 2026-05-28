Attribute VB_Name = "CANUSB_API"
'
' File:       canusb.bas
'
' Purpose:    VB6 Declarations for CANUSB from LAWICEL AB
'
' Version:    1.0.1, 28:th of January 2006
'
' Author:     Lars Wictorsson
'             LAWICEL AB / SWEDEN
'             http://www.can232.com
'             http://www.canusb.com
'             http://www.lawicel.se
'
' Copyright:  The copyright to the computer program(s) herein is the
'             property of LAWICEL AB, Sweden. The program(s) may be used
'             and/or copied only with the written permission of LAWICEL AB
'             in accordance with the terms and conditions stipulated in
'             the agreement/contract under which the program(s) have been
'             supplied.
'
' Commets:    Needs at least CANUSB driver DLL 0.0.14
'
' History:    2005-06-18  1.0.0   Created (LWI)
'             2006-01-28  1.0.1   Created (LWI)
'
 
'
' Public Const
'
Public Const CAN_MAX_STANDARD_ID = &H7FF
Public Const CAN_MAX_EXTENDED_ID = &H1FFFFFFF
'
' BTR0BTR1 register values for BTR0/BTR1
'
Public Const CAN_BAUD_BTR_1M = "0x00:0x14"          '   1 MBit / s
Public Const CAN_BAUD_BTR_500K = "0x00:1C"          ' 500 kBit / s
Public Const CAN_BAUD_BTR_250K = "0x01:0x1C"        ' 250 kBit / s
Public Const CAN_BAUD_BTR_125K = "0x3:0x1C"         ' 125 kBit / s
Public Const CAN_BAUD_BTR_100K = "0x43:0x2F"        ' 100 kBit / s
Public Const CAN_BAUD_BTR_50K = "0x47:0x2F"         '  50 kBit / s
Public Const CAN_BAUD_BTR_20K = "0x53:0x2F"         '  20 kBit / s
Public Const CAN_BAUD_BTR_10K = "0x67:0x2F"         '  10 kBit / s
Public Const CAN_BAUD_BTR_5K = "0x7F:0x7F"          '   5 kBit / s
'
' Baudrate can also be set with "real" value if set to one of the values below
'
Public Const CAN_BAUD_1M = "1000"                   '   1 MBit / s
Public Const CAN_BAUD_800K = "800"                  ' 800 kBit / s
Public Const CAN_BAUD_500K = "500"                  ' 500 kBit / s
Public Const CAN_BAUD_250K = "250"                  ' 250 kBit / s
Public Const CAN_BAUD_125K = "125"                  ' 125 kBit / s
Public Const CAN_BAUD_100K = "100"                  ' 100 kBit / s
Public Const CAN_BAUD_50K = "50"                    '  50 kBit / s
Public Const CAN_BAUD_20K = "20"                    '  20 kBit / s
Public Const CAN_BAUD_10K = "10"                    '  10 kBit / s
'
' Status bits
'
Public Const CANSTATUS_RECEIVE_FIFO_FULL = &H1
Public Const CANSTATUS_TRANSMIT_FIFO_FULL = &H2
Public Const CANSTATUS_ERROR_WARNING = &H4
Public Const CANSTATUS_DATA_OVERRUN = &H8
Public Const CANSTATUS_ERROR_PASSIVE = &H20
Public Const CANSTATUS_ARBITRATION_LOST = &H40
Public Const CANSTATUS_BUS_ERROR = &H80
'
'  Error return codes
'
Public Const ERROR_CANUSB_OK = 1                    ' All is OK
Public Const ERROR_CANUSB_OPEN_SUBSYSTEM = -2       ' Problems with driver subsystem
Public Const ERROR_CANUSB_COMMAND_SUBSYSTEM = -3    ' Unable to send command to adapter
Public Const ERROR_CANUSB_NOT_OPEN = -4             ' Channel not open
Public Const ERROR_CANUSB_TX_FIFO_FULL = -5         ' Transmit fifo full
Public Const ERROR_CANUSB_INVALID_PARAM = -6        ' Invalid parameter
Public Const ERROR_CANUSB_NO_MESSAGE = -7           ' No message available
'
' Msg Type:
'
Public Const CANMSG_EXTENDED = &H80                 ' Extended Frame
Public Const CANMSG_RTR = &H40                      ' Standart Frame
'
' CAN Frame
'
Public Type canmsg
    ID As Long                                      ' Message id 11/29 Bit
    TimeStamp As Long                               ' timestamp in milliseconds
    flags As Byte                                   ' [extended_id|1][RTR:1][reserver:6]
    len As Byte                                     ' Frame size (0.8)
    data(7) As Byte                                 ' Databytes 0..7
End Type
'
' Alternative CAN Frame
'
Public Type canmsgex
    ID As Long                                      ' Message id 11/29 Bit
    TimeStamp As Long                               ' timestamp in milliseconds
    flags As Byte                                   ' [extended_id|1][RTR:1][reserver:6]
    len As Byte                                     ' Frame size (0.8)
End Type
'
' Open flags
'
Public Const CANUSB_FLAG_TIMESTAMP = &H1            ' Timestamp messages
Public Const CANUSB_FLAG_QUEUE_REPLACE = &H2        ' If input queue is full remove oldest message and insert new message.
'
' Filter mask settings
'
Public Const CANUSB_ACCEPTANCE_CODE_ALL = &H0
Public Const CANUSB_ACCEPTANCE_MASK_ALL = &HFFFFFFFF
'
' Flush Flags
'
Public Const FLUSH_WAIT = &H0
Public Const FLUSH_DONTWAIT = &H1

Public Status As Long

'///////////////////////////////////////////////////////////////////////////////
'// cansub_Open
'//
'//
'// Open CAN interface to device
'//
'// Returs handle to device if open was successfull or zero
'// on falure.
'//
'//
'// szID
'// ====
'// Serial number for adapter or NULL to open the first found.
'//
'//
'// szBitrate
'// =========
'// "10" for 10kbps
'// "20" for 20kbps
'// "50" for 50kbps
'// "100" for 100kbps
'// "250" for 250kbps
'// "500" for 500kbps
'// "800" for 800kbps
'// "1000" for 1Mbps
'//
'// or
'//
'// btr0:btr1 pair  ex. "0x03:0x1c" or 3:28
'//
'// acceptance_code
'// ===============
'// Set to CANUSB_ACCEPTANCE_CODE_ALL to  get all messages.
'//
'// acceptance_mask
'// ===============
'// Set to CANUSB_ACCEPTANCE_MASk_ALL to  get all messages.
'//
'// flags
'// =====
'// CANUSB_FLAG_TIMESTAMP - Timestamp will be set by adapter.

Public Declare Function canusb_Open Lib "canusbdrv" (ByVal szID As String, ByVal szBitrate As String, ByVal acceptance_code As Long, ByVal acceptance_mask As Long, ByVal flags As Long) As Long


'///////////////////////////////////////////////////////////////////////////////
'// canusb_Close
'//
'// Close channel with handle h.
'//
'// Returns <= 0 on failure. >0 on success.

Public Declare Function canusb_Close Lib "canusbdrv" (ByVal h As Long) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_Read
'//
'// Read message from channel with handle h.
'//
'// Returns <= 0 on failure. >0 on success.
'//

Public Declare Function canusb_Read Lib "canusbdrv" (ByVal h As Long, ByRef canmsg As canmsg) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_ReadEx
'//
'// Read message from channel with handle h.
'//
'// This is a version without a data-array in the structure to work with e.g. LabView
'//
'// Returns <= 0 on failure. >0 on success.
'//

'Public Declare Function canusb_ReadEx Lib "canusbdrv" (ByVal h As Long, ByRef canmsgex As canmsgex, ByRef candata As candata) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_ReadFirst
'//
'// Read first message from channel with handle h and id "id" which satisfying flags.
'//
'// Returns <= 0 on failure. >0 on success.
'//

Public Declare Function canusb_ReadFirst Lib "canusbdrv" (ByVal h As Long, ByVal ID As Long, ByVal flags As Byte, ByRef canmsg As canmsg) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_Write
'//
'// Write message to channel with handle h.
'//
'// Returns <= 0 on failure. >0 on success.
'//

Public Declare Function canusb_Write Lib "canusbdrv" (ByVal h As Long, ByRef canmsg As canmsg) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_Status
'//
'// Get Adaper status for channel with handle h.

Public Declare Function canusb_Status Lib "canusbdrv" (ByVal h As Long) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_VersionInfo
'//
'// Get hardware/fi4rmware and driver version for channel with handle h.
'//
'// Returns <= 0 on failure. >0 on success.
'//
'// Format
'//  ??? "Hardware_Major.Hardware_Minor;Firmware_Major.Firmware_Minor;Driver_Major.Driver_Minor"
'//

Public Declare Function canusb_VersionInfo Lib "canusbdrv" (ByVal h As Long, ByVal verinfo As String) As Integer


'///////////////////////////////////////////////////////////////////////////////
'// canusb_Flush
'//
'// Flush output buffer on channel with handle h.
'//
'// Returns <= 0 on failure. >0 on success.
'//
'// If flushflags is set to FLUSH_DONTWAIT the queue is just emptied and
'// there will be no wait for any frames in it to be sent
'//

Public Declare Function canusb_Flush Lib "canusbdrv" (ByVal h As Long, ByVal flushflags As Byte) As Integer








Dim handle As Long
Dim bOpen As Boolean
Dim iDummy As Integer
Dim iDummy100 As Integer
Dim iDummy200 As Integer
Dim CanUsbName As String

Public Sub SetCanusbDevice(Name As String)
    CanUsbName = Name
End Sub

Public Function Flushport() As Boolean
    Dim Status As Integer
    If (bOpen = True) Then
        Status = canusb_Flush(handle, FLUSH_DONTWAIT)
        If (Status <> ERROR_CANUSB_OK) Then
            Flushport = True
        Else
            Flushport = fasòe
        End If
        
    Else
        Flushport = True
    End If
    
End Function

Public Function GetListCanUsb() As String
    Dim Temphandle As Long
    Dim str As String
    Dim i As Integer
    Dim Version As String * 255
    Dim GetVersion As String
    Dim xhandle() As Long
    Dim NumeroDriverCanMounted As Integer
    Temphandle = 1
    
    ReDim Preserve xhandle(NumeroDriverCanMounted)
    
    Do While (Temphandle <> 0)
        Temphandle = canusb_Open(vbNullString, "250", CANUSB_ACCEPTANCE_CODE_ALL, CANUSB_ACCEPTANCE_MASK_ALL, CANUSB_FLAG_TIMESTAMP)
        
        If (Temphandle <> 0) Then
        
    
            If canusb_VersionInfo(Temphandle, Version) = ERROR_CANUSB_OK Then
                GetVersion = Left$(Version, InStr(Version, vbNullChar) - 1)
            End If
        
            GetListCanUsb = GetListCanUsb + GetVersion + vbLf
            
            xhandle(NumeroDriverCanMounted) = Temphandle
            NumeroDriverCanMounted = NumeroDriverCanMounted + 1
            ReDim Preserve xhandle(NumeroDriverCanMounted)
            
            
        End If
        
    Loop
    
    Do While (i < NumeroDriverCanMounted)
        Status = canusb_Flush(xhandle(i), FLUSH_WAIT)
         canusb_Close (xhandle(i))
        i = i + 1
    Loop
    
 
    
End Function

Public Function GetVersion(Optional ByVal handlex As Long = 0) As String

    Dim Version As String * 255
        If (handlex = 0) Then
        If bOpen Then
            retval = canusb_VersionInfo(handle, Version)
            If retval = ERROR_CANUSB_OK Then
                GetVersion = Left$(Version, InStr(Version, vbNullChar) - 1)
            End If
        End If
    Else
            retval = canusb_VersionInfo(handlex, Version)
            If retval = ERROR_CANUSB_OK Then
                GetVersion = Left$(Version, InStr(Version, vbNullChar) - 1)
            End If
    End If
End Function


Public Function OpenPort(bOpenCmd As Boolean, Optional Name As String = "") As Boolean
    ' This variant opens first found Lawicel canusb adapter
    
    Dim Temphandle As Long
    If (bOpenCmd = True) Then
        Temphandle = canusb_Open(Name, "250", CANUSB_ACCEPTANCE_CODE_ALL, CANUSB_ACCEPTANCE_MASK_ALL, CANUSB_FLAG_TIMESTAMP)
'        If (GetVersion(Temphandle) = Name) Then
'            bOpen = True
'            handle = Temphandle
'        End If
        ' This variant opens Lawicel canusb adapter with give serial number
        'handle = canusb_Open("LWNQ06ES", "500", CANUSB_ACCEPTANCE_CODE_ALL, CANUSB_ACCEPTANCE_MASK_ALL, CANUSB_FLAG_TIMESTAMP)
        If (Temphandle <> 0) Then
            bOpen = True
            handle = Temphandle
        End If
    Else
        If (bOpen = True) Then
            Dim Status As Integer
            Status = canusb_Flush(handle, FLUSH_WAIT)
            canusb_Close (handle)
            bOpen = False
        End If
    End If
    OpenPort = bOpen
End Function


'Public Sub OpenPort()
'    MSComm1.CommPort = (ComUsbCan)
'    MSComm1.Settings = "57600,N,8,1" '"38400,N,8,1"
'    MSComm1.PortOpen = True
'    MSComm1.InBufferCount = 0
'    MSComm1.RThreshold = 1
'
'End Sub

Public Function Output_Test(ByRef str As String) As Long
    On Error GoTo err:
   'Me.Visible = False
    Dim Index As String
    'Dim Status As Long
    Dim Dlc As String
    Dim i As Integer
    Dim data(8) As String
    Dim count As Long
    
    If (bOpen = True) Then
        Index = Mid(str, 2, 8)
        Dlc = Mid(str, 10, 1)
        
        Do While (i < CInt(Dlc))
            data(i) = Mid(str, 11 + i * 2, 2)
            i = i + 1
        Loop
        
        
        Dim msg As canmsg
        If (bOpen) Then
            msg.ID = CLng("&H" + Index)
            msg.flags = CANMSG_EXTENDED
            msg.len = CInt(Dlc)
            
            i = 0
            Do While (i < msg.len)
                msg.data(i) = CInt("&H" + data(i))
                i = i + 1
            Loop
riprova:
            count = count + 1
            If (count > 1) Then
                GoTo err:
            End If
            
'            Status = canusb_Status(handle)
'            If (Status <> 0) Then
'                GoTo riprova
'            End If
            retval = canusb_Write(handle, msg)
            If (retval <> ERROR_CANUSB_OK) Then
                GoTo riprova
            End If
        End If
    End If
err:
    Output_Test = retval
End Function




Public Sub Output(ByRef str As String)
    On Error GoTo err:
   'Me.Visible = False
    Dim Index As String
    'Dim Status As Long
    Dim Dlc As String
    Dim i As Integer
    Dim data(8) As String
    Dim count As Long
    
    If (bOpen = True) Then
        Index = Mid(str, 2, 8)
        Dlc = Mid(str, 10, 1)
        
        Do While (i < CInt(Dlc))
            data(i) = Mid(str, 11 + i * 2, 2)
            i = i + 1
        Loop
        
        
        Dim msg As canmsg
        If (bOpen) Then
            msg.ID = CLng("&H" + Index)
            msg.flags = CANMSG_EXTENDED
            msg.len = CInt(Dlc)
            
            i = 0
            Do While (i < msg.len)
                msg.data(i) = CInt("&H" + data(i))
                i = i + 1
            Loop
riprova:
            count = count + 1
            If (count > 1) Then
                GoTo err:
            End If
            
'            Status = canusb_Status(handle)
'            If (Status <> 0) Then
'                GoTo riprova
'            End If
            retval = canusb_Write(handle, msg)
            If (retval <> ERROR_CANUSB_OK) Then
                GoTo riprova
            End If
        End If
    End If
err:
End Sub

Public Sub canUsbTask()
    Dim retval As Integer
    Dim msg As canmsg
    Dim str As String
    Dim i As Integer
riprova:
    If (bOpen = True) Then
        retval = canusb_Read(handle, msg)
        
        If (retval = 1) Then
            str = "T" + Convert(msg.ID, 8)
            str = str + Convert(msg.len, 1)
            
            i = 0
            Do While (i < msg.len)
                str = str + Convert(msg.data(i), 2)
                i = i + 1
            Loop
    
            str = str + Chr(13)
            
            DriverCanUsbRxBuffer str
            
            
            
            GoTo riprova
        End If
    End If
End Sub


Private Function Convert(ByVal data As Long, Length As Integer) As String
    
    Dim TempStr As String
    
    TempStr = Hex(data)
    
    Do While (Len(TempStr) < Length)
        TempStr = "0" + TempStr
    Loop

    Convert = TempStr
End Function


