Attribute VB_Name = "modRegistration"
Option Explicit

'Vars for Authentication
Global AuthKey As Boolean
Global AuthString As String

' Reg Key Security Options...
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
' Reg Key ROOT Types...
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const ERROR_SUCCESS = 0
Public Const REG_SZ = 1                         ' Unicode nul terminated string
Public Const REG_DWORD = 4                      ' 32-bit number

Public Const gREGKEYSYSINFOLOC = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Public Const gREGKEYSYSINFO = "System\CurrentControlSet\Control\ComputerName\ComputerName"
Public Const gREGVALSYSINFO = "ComputerName"
Public Const RegKey = "Reg"
Global Register As String

'String to hold Registry Computer Name
Global SysInfoPath As String

Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Private Declare Function GetVolumeInformation Lib "kernel32.dll" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Integer, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Declare Function GetComputerNameAPI Lib "kernel32" Alias "GetComputerNameA" _
     (ByVal lpBuffer As String, nSize As Long) As Long

'Put your encryption string in the
'EncryptName Variable 20 characters
Global Const EncryptName = "Environmentvariables"  '"Putencyptstringhere "

'Put your project name here
'This is an entry in the registry that is created
Const RegPath = "SOFTWARE\Project"

'Data for Date settings
Public GMonth(1 To 12) As String
Public GYear(1 To 22) As String
Public GDuration(1 To 24) As String

Public Sub StartSysInfo()
    
    On Error GoTo SysInfoErr
  
    ' Try To Get System Info Program Path\Name From Registry...
    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
    '
    Else
        GoTo SysInfoErr
    End If
    
    Exit Sub

SysInfoErr:
    MsgBox "System Information Is Unavailable At This Time", vbOKOnly, Mtitle
    
End Sub

Public Function GetKeyValue(KeyRoot As Long, Keyname As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
    Dim i As Long                                           ' Loop Counter
    Dim rc As Long                                          ' Return Code
    Dim hKey As Long                                        ' Handle To An Open Registry Key
    Dim hDepth As Long                                      '
    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
    '------------------------------------------------------------
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    rc = RegOpenKeyEx(KeyRoot, Keyname, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
    
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
    
    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
    KeyValSize = 1024                                       ' Mark Variable Size
    
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
                        
    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
    
    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
    Else                                                    ' WinNT Does NOT Null Terminate String...
        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
    End If
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case KeyValType                                  ' Search Data Types...
    Case REG_SZ                                             ' String Registry Key Data Type
        KeyVal = tmpVal                                     ' Copy String Value
    Case REG_DWORD                                          ' Double Word Registry Key Data Type
        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
        Next
        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
    End Select
    
    GetKeyValue = True                                      ' Return Success
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
    Exit Function                                           ' Exit
    
GetKeyError:      ' Cleanup After An Error Has Occured...
    KeyVal = ""                                             ' Set Return Val To Empty String
    GetKeyValue = False                                     ' Return Failure
    rc = RegCloseKey(hKey)                                  ' Close Registry Key
End Function

Sub InvertIt(pInverStr As String)
    Dim Temp As Integer
    Dim Hold As Integer
    Dim i As Integer
    Dim TempStr As String
        
    TempStr = ""
    For i = 1 To Len(pInverStr)
        Temp = Asc(Mid$(pInverStr, i, 1))
        Hold = 0
Top:
    Select Case Temp
        Case Is > 127
            Hold = Hold + 1
            Temp = Temp - 128
            GoTo Top
        Case Is > 63
            Hold = Hold + 2
            Temp = Temp - 64
            GoTo Top
        Case Is > 31
            Hold = Hold + 4
            Temp = Temp - 32
            GoTo Top
        Case Is > 15
            Hold = Hold + 8
            Temp = Temp - 16
            GoTo Top
        Case Is > 7
            Hold = Hold + 16
            Temp = Temp - 8
            GoTo Top
        Case Is > 3
            Hold = Hold + 32
            Temp = Temp - 4
            GoTo Top
        Case Is > 1
            Hold = Hold + 64
            Temp = Temp - 2
            GoTo Top
        Case Is = 1
            Hold = Hold + 128
            
    End Select
        Temp = 255 Xor Hold
        TempStr = TempStr + Chr(Temp)
    Next i
    
    pInverStr = TempStr
End Sub

Sub EncryptIt(pEncryptStr As String)
    Dim Temp As Integer
    Dim Temp1 As Integer
    Dim Hold As Integer
    Dim i As Integer
    Dim j As Integer
    Dim TempStr As String

    TempStr = ""
    For i = 1 To Len(EncryptName)
        Hold = 0
        Temp = Asc(Mid$(EncryptName, i, 1))
        For j = 1 To Len(pEncryptStr)
            Temp1 = Asc(Mid$(pEncryptStr, j, 1))
            Hold = Temp Xor Temp1
         Next j
        TempStr = TempStr + Chr(Hold)
    Next i
    
    pEncryptStr = TempStr
End Sub

Sub EncipherIt(pEnciperStr As String)
    Dim Temp As Integer
    Dim Hold As String
    Dim i As Integer
    Dim j As Integer
    Dim TempStr As String
    Dim Temp1 As String
    
    TempStr = ""
    For i = 1 To Len(pEnciperStr)
        Temp = Asc(Mid$(pEnciperStr, i, 1))
        Temp1 = Hex(Temp)
        If Len(Temp1) = 1 Then
            Temp1 = "0" & Temp1
        End If
        For j = 1 To 2
            Hold = Mid$(Temp1, j, 1)
            Select Case Hold
                Case "0"
                    TempStr = TempStr + "7"
                Case "1"
                    TempStr = TempStr + "B"
                Case "2"
                    TempStr = TempStr + "F"
                Case "3"
                    TempStr = TempStr + "D"
                Case "4"
                    TempStr = TempStr + "1"
                Case "5"
                    TempStr = TempStr + "9"
                Case "6"
                    TempStr = TempStr + "3"
                Case "7"
                    TempStr = TempStr + "A"
                Case "8"
                    TempStr = TempStr + "6"
                Case "9"
                    TempStr = TempStr + "5"
                Case "A"
                    TempStr = TempStr + "E"
                Case "B"
                    TempStr = TempStr + "8"
                Case "C"
                    TempStr = TempStr + "0"
                Case "D"
                    TempStr = TempStr + "C"
                Case "E"
                    TempStr = TempStr + "2"
                Case "F"
                    TempStr = TempStr + "4"
            End Select
        Next j
    Next i
    pEnciperStr = TempStr
End Sub

Public Sub GetSubKey()

    If Not GetKeyValue(HKEY_LOCAL_MACHINE, RegPath, RegKey, Register) Then
        'Rem Not in registry
        
    End If
    
End Sub

'GetSerialNumber Procedure - Put this in the module or form where it is called.
Function GetSerialNumber(strDrive As String) As Long
    Dim SerialNum As Long
    Dim Res As Long
    Dim Temp1 As String
    Dim Temp2 As String
    'initialise the strings
    Temp1 = String$(255, Chr$(0))
    Temp2 = String$(255, Chr$(0))
    'call the API function
    Res = GetVolumeInformation(strDrive, Temp1, Len(Temp1), SerialNum, 0, 0, Temp2, Len(Temp2))
    GetSerialNumber = SerialNum
    
End Function

Public Function GetHardDiskKey() As String
    'Dimension our variables
    Dim TempStr As String
    Dim RegStr As String
    Dim i As Integer
    Dim SerialNumber As Long
    
    'Get The Computer Name in the registry
    'StartSysInfo
    SerialNumber = GetSerialNumber("C:\")
    SysInfoPath = Str(SerialNumber)
    
    'For encrypting purposes make the length
    'of it no more than 20 character
    If Len(SysInfoPath) > 20 Then
        SysInfoPath = Left$(SysInfoPath, 20)
    End If
    'invert the computer name
    Call InvertIt(SysInfoPath)
    Call EncryptIt(SysInfoPath)
    Call EncipherIt(SysInfoPath)
    GetSubKey
    
    'verify it
    If Len(SysInfoPath) > 20 Then
        SysInfoPath = Left$(SysInfoPath, 20)
    End If
GetHardDiskKey = SysInfoPath
End Function


Public Function GetDecodedKey(ByVal pHardDiskKey As String, pLength As Integer) As String
Dim AuthString As String
Dim intIndex As Integer
Dim TempStr As String
    If Len(pHardDiskKey) > pLength Then
        pHardDiskKey = Left$(pHardDiskKey, pLength)
    End If
    For intIndex = 1 To Len(pHardDiskKey)
        TempStr = Mid$(pHardDiskKey, intIndex, 1)
        Select Case TempStr
            Case "0"
                AuthString = AuthString + "V"
            Case "1"
                AuthString = AuthString + "I"
            Case "2"
                AuthString = AuthString + "K"
            Case "3"
                AuthString = AuthString + "P"
            Case "4"
                AuthString = AuthString + "O"
            Case "5"
                AuthString = AuthString + "Q"
            Case "6"
                AuthString = AuthString + "S"
            Case "7"
                AuthString = AuthString + "H"
            Case "8"
                AuthString = AuthString + "G"
            Case "9"
                AuthString = AuthString + "T"
            Case "A"
                AuthString = AuthString + "U"
            Case "B"
                AuthString = AuthString + "J"
            Case "C"
                AuthString = AuthString + "N"
            Case "D"
                AuthString = AuthString + "L"
            Case "E"
                AuthString = AuthString + "M"
            Case "F"
                AuthString = AuthString + "R"
        End Select
    Next
GetDecodedKey = AuthString
End Function


Private Function DecryptDate(ByVal pStrData As String) As String
Dim intIndex As Integer
Dim lTempstr As String
    For intIndex = 1 To Len(pStrData)
        If intIndex = 3 Or intIndex = 4 Or intIndex = 5 Then
            'Get exact key
            lTempstr = lTempstr & Mid(pStrData, intIndex, 1)
        Else
            Select Case Mid(pStrData, intIndex, 1)
                Case "X"
                    lTempstr = lTempstr & "1"
                Case "C"
                    lTempstr = lTempstr & "2"
                Case "J"
                    lTempstr = lTempstr & "3"
                Case "D"
                    lTempstr = lTempstr & "4"
                Case "A"
                    lTempstr = lTempstr & "5"
                Case "C"
                    lTempstr = lTempstr & "6"
                Case "Y"
                    lTempstr = lTempstr & "7"
                Case "O"
                    lTempstr = lTempstr & "8"
                Case "P"
                    lTempstr = lTempstr & "9"
                Case "E"
                    lTempstr = lTempstr & "0"
                Case Else
                    lTempstr = lTempstr & Mid(pStrData, intIndex, 1)
            End Select
        End If
    Next
DecryptDate = lTempstr
End Function

Public Sub GetDateSettings(pStrData As String, pDate As String, pDuration As Integer)
Dim lTempstr As String
Dim intIndex As Integer
Dim lMonth
Dim lYear

    lTempstr = DecryptDate(pStrData)
    If Len(lTempstr) = 5 Then
        'Parse Date
        pDate = Mid(lTempstr, 1, 2)
        'Parse Month
        For intIndex = 1 To 12
            If GMonth(intIndex) = Mid(lTempstr, 3, 1) Then
                lMonth = intIndex ' MonthName(intIndex)
            End If
        Next
        'Parse Year
        For intIndex = 1 To 22
            If GYear(intIndex) = Mid(lTempstr, 4, 1) Then
                lYear = 2003 + intIndex
            End If
        Next
        'Parse Duration
        For intIndex = 1 To 24
            If GDuration(intIndex) = Mid(lTempstr, 5, 1) Then
                pDuration = intIndex
            End If
        Next
        'Date
        pDate = lMonth & "/" & pDate & "/" & lYear
    End If
End Sub

Public Function GetComputerName() As String
   'Set or retrieve the name of the computer.
   Dim sBuffer As String
   Dim lLen As Long

   'Pad string with spaces
   sBuffer = Space(255 + 1)
   lLen = Len(sBuffer)

   If CBool(GetComputerNameAPI(sBuffer, lLen)) Then
      GetComputerName = Left(sBuffer, lLen)
   Else
      GetComputerName = ""
   End If
End Function
