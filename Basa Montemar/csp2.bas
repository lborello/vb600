Attribute VB_Name = "CSP2"
'*******************************************************************************
' *
' *  Filename:        csp2.bas
' *j
' *  Copyright(c) Symbol Technologies Inc., 2001
' *
' *  Description:     Defines constants and Declares for use with the csp2.dll
' *
' *  Author:          Chris Brock.
' *
' *  Creation Date:   ?/?/??
' *
' *  Derived From:    New File
' *
' *  Edit History:
' *   $Log:   U:/keyfob/archives/winfob/examples/vb/csp2.baV  $
Rem
Rem    Rev 1.2   Feb 01 2002 09:30:00   pangr
Rem Removed commented out code
Rem
Rem    Rev 1.1   Jan 29 2002 16:09:08   pangr
Rem Changed Declare statements of functions that accept/return
Rem non-ASCII strings from type String to Byte arrays.
'*******************************************************************************/

Public Enum CommPorts
    COM1 = 0
    COM2 = 1
    COM3 = 2
    COM4 = 3
    COM5 = 4
    COM6 = 5
    COM7 = 6
    COM8 = 7
    COM9 = 8
    COM10 = 9
    COM11 = 10
    COM12 = 11
    COM13 = 12
    COM14 = 13
    COM15 = 14
    COM16 = 15
End Enum


'// Communications
Declare Function csp2Init Lib "csp2.dll" (ByVal nComPort As Long) As Long
Declare Function csp2Restore Lib "csp2.dll" () As Long
Declare Function csp2WakeUp Lib "csp2.dll" () As Long
Declare Function csp2DataAvailable Lib "csp2.dll" () As Long

'// Basic Functions
Declare Function csp2ReadData Lib "csp2.dll" () As Long
Declare Function csp2ClearData Lib "csp2.dll" () As Long
Declare Function csp2PowerDown Lib "csp2.dll" () As Long
Declare Function csp2GetTime Lib "csp2.dll" (aTimeBuf As Byte) As Long
Declare Function csp2SetTime Lib "csp2.dll" (aTimeBuf As Byte) As Long
Declare Function csp2SetDefaults Lib "csp2.dll" () As Long

'// CSP Data Get
Declare Function csp2GetPacket Lib "csp2.dll" (stPacketData As Byte, ByVal lgBarcodeNumber As Long, ByVal maxLength As Long) As Long
Declare Function csp2GetDeviceId Lib "csp2.dll" (szDeviceId As Byte, ByVal nMaxLength As Long) As Long
Declare Function csp2GetProtocol Lib "csp2.dll" () As Long
Declare Function csp2GetSystemStatus Lib "csp2.dll" () As Long
Declare Function csp2GetSwVersion Lib "csp2.dll" (ByVal szSwVersion As String, ByVal nMaxLength As Long) As Long
Declare Function csp2GetASCIIMode Lib "csp2.dll" () As Long
Declare Function csp2GetRTCMode Lib "csp2.dll" () As Long

'// DLL Configuration
Declare Function csp2SetRetryCount Lib "csp2.dll" (ByVal nRetryCount As Long) As Long
Declare Function csp2GetRetryCount Lib "csp2.dll" () As Long

'// Miscellaneous
Declare Function csp2GetDllVersion Lib "csp2.dll" (ByVal szDllVersion As String, ByVal nMaxLength As Long) As Long
Declare Function csp2TimeStamp2Str Lib "csp2.dll" (Stamp As Byte, ByVal value As String, ByVal nMaxLength As Long) As Long
Declare Function csp2GetCodeType Lib "csp2.dll" (ByVal CodeID As Long, ByVal CodeType As String, ByVal nMaxLength As Long) As Long

'// Advanced functions
Declare Function csp2ReadRawData Lib "csp2.dll" (aBuffer As Byte, ByVal nMaxLength As Long) As Long
Declare Function csp2SetParam Lib "csp2.dll" (ByVal nParam As Long, szString As Byte, ByVal nMaxLength As Long) As Long
Declare Function csp2GetParam Lib "csp2.dll" (ByVal nParam As Long, szString As Byte, ByVal nMaxLength As Long) As Long
Declare Function csp2Interrogate Lib "csp2.dll" () As Long
Declare Function csp2GetCTS Lib "csp2.dll" () As Long
Declare Function csp2SetDTR Lib "csp2.dll" (ByVal nOnOff As Long) As Long
Declare Function csp2SetDebugMode Lib "csp2.dll" (ByVal nOnOff As Long) As Long

'// Returned status values...
Global Const STATUS_OK                 As Long = 0
Global Const COMMUNICATIONS_ERROR      As Long = -1
Global Const BAD_PARAM                 As Long = -2
Global Const SETUP_ERROR               As Long = -3
Global Const INVALID_COMMAND_NUMBER    As Long = -4
Global Const COMMAND_LRC_ERROR         As Long = -7
Global Const RECEIVED_CHARACTER_ERROR  As Long = -8
Global Const GENERAL_ERROR             As Long = -9
Global Const FILE_NOT_FOUND            As Long = 2
Global Const ACCESS_DENIED             As Long = 5

'// Parameter values...
Global Const PARAM_OFF                 As Long = 0
Global Const PARAM_ON                  As Long = 1

Global Const DATA_AVAILABLE            As Long = 1
Global Const DATA_NOT_AVAILABLE        As Long = 0

Global Const DETERMINE_SIZE            As Long = 0
                                                   
