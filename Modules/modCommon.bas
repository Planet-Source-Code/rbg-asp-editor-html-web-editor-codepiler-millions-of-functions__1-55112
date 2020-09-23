Attribute VB_Name = "modCommon"
Option Explicit

Rem ===================================
Rem PUBLIC ENUMS/TYPES
Rem ===================================
Public Enum ChildArrange
  Cascade = 0
  TileHorizontal
  TileVertical
  Icons
End Enum
Public Enum ModifyTypes
    AddText = 0
    DeleteText = 1
    ReplaceText = 2
    CutText = 3
    PasteText = 4
End Enum
Public Enum TriState
  tsTrue = -1
  tsFalse = 0
  tsnone = -2
End Enum
Public Type CHARRANGE
  cpMin As Long
  cpMax As Long
End Type
Public Type MEMORYSTATUS
  dwLength As Long
  dwMemoryLoad As Long
  dwTotalPhys As Long
  dwAvailPhys As Long
  dwTotalPageFile As Long
  dwAvailPageFile As Long
  dwTotalVirtual As Long
  dwAvailVirtual As Long
End Type
Public Type TEXTRANGE
  chrg As CHARRANGE
  lpstrText As Long    ' /* allocated by caller, zero terminated by RichEdit */
End Type
Public Type POINTAPI
    x As Long
    y As Long
End Type


Rem ===================================
Rem PUBLIC DECLARATIONS
Rem ===================================
Public mIndex As Integer 'For open documents in array
Public Pt          As POINTAPI
Public lngStart1    As Long
Public MoldLocalHost As String
Public Msitedetails As clsSites
Public Mhistories As clsHistories
Public Mfindin As Integer 'Save the findin setting
Public Mfindwhat As String 'Save the findwhat setting
Public Mreplace As String 'Save the replace setting
Public Mcasesensitive As Integer 'Save the case sensitive setting
Public Mwholeword As Integer 'Save the whole word setting
Public Mtitle As String 'Title fo codepiler with version
Public Mcode As Integer 'Registration status code
Public Mreuse As Boolean 'Reuse of splash screen to aboutus screen

Rem ===================================
Rem PUBLIC APIS
Rem ===================================
Public Declare Function SendMessage Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageLong Lib "User32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function OSWinHelp% Lib "User32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SetWindowTheme Lib "uxtheme.dll" (ByVal hwnd As Long, ByVal pszSubAppName As Long, ByVal pszSubIdList As Long) As Long
Public Declare Function GetCaretPos Lib "User32" (lpPoint As POINTAPI) As Long
Public Declare Function SetParent Lib "User32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Sub ReleaseCapture Lib "User32" ()
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Rem ===================================
Rem PUBLIC CONSTANTS
Rem ===================================
Public Const WM_USER = &H400
Public Const EM_HIDESELECTION = WM_USER + 63
Public Const EM_GETLINECOUNT = &HBA        '// Total Line Count
Public Const EM_GETFIRSTVISIBLELINE = &HCE '// First Visible Line
Public Const WM_VSCROLL = &H115            '// Vertical Scrolling
Public Const EM_GETTEXTRANGE = (WM_USER + 75)
Public Const CB_SETDROPPEDWIDTH = &H160&
Public Const EM_SETTARGETDEVICE = (WM_USER + 72)
Public Const NAME_COLUMN = 0
Public Const TYPE_COLUMN = 1
Public Const SIZE_COLUMN = 2
Public Const DATE_COLUMN = 3
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const STATUS_PENDING = &H103&  'new
Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOW = 5
Public Const vbKeyLessThan = 60
Public Const EM_LINEFROMCHAR = &HC9
Public Const EM_LINEINDEX = &HBB
Public Const htMaxEntityVal = 63
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const HWND_NOTOPMOST = -2
Public Const quote = """"

Rem ===================================
Rem PUBLIC FUNCTIONS
Rem ===================================

Public Function S101_Make_Dir(ByVal pName As String)
Rem ---------------------------------
Rem Make folder
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
Dim lTmp As String
Dim lPath As Variant
Dim li As Integer
On Error Resume Next
  Set lFso = New Scripting.FileSystemObject
  If lFso.FolderExists(lPath) Then Exit Function
  lPath = Split(pName, "\")
  For li = LBound(lPath) To UBound(lPath)
    lTmp = lTmp & lPath(li) & "\"
    If Not lFso.FolderExists(lTmp) Then Call lFso.CreateFolder(lTmp)
  Next
  Err.Clear
End Function

Public Function S102_File_Exists(ByVal pFilename As String) As Boolean
Rem ---------------------------------
Rem Check for file existence
Rem ---------------------------------

On Error GoTo S102_Err
  If FileLen(pFilename) > 0 Then
    S102_File_Exists = True
  Else
    S102_File_Exists = False
  End If
  GoTo S102_Out
S102_Err:
  S102_File_Exists = False
S102_Out:
End Function

Public Function S103_FormatRGBString(val As Long) As String
Rem -------------------------------
Rem Format the long color to string
Rem -------------------------------

Dim Color As String
Dim pad As Long
Dim r As String
Dim g As String
Dim b As String

  On Error Resume Next
  Color = Hex(val)
  'determine how many zeros to pad in front of converted value
  pad = 6 - Len(Color)
  
  If pad Then
      Color = String(pad, "0") & Color
  End If
      
  'Extract the rgb components
  r = Right(Color, 2)
  g = Mid(Color, 3, 2)
  b = Left(Color, 2)
  
  ' Swab r and b position, color dialog returns
  ' bgr instead of rgb
  Color = "#" & r & g & b
  
  S103_FormatRGBString = Color
End Function

Public Function S104_Rename(ByVal pPath As String, ByVal pNew As String, Optional ByVal pFolder As Boolean) As Boolean
'
'Rename the folder/file
'
Dim lFso As New FileSystemObject
  On Error GoTo Cerr
  If pFolder Then
    lFso.GetFolder(pPath).Name = pNew
  Else
    lFso.GetFile(pPath).Name = pNew
  End If
  S104_Rename = True
  Exit Function
Cerr:
  S104_Rename = False
End Function

Public Function S105_Delete(ByVal pPath As String, Optional ByVal pFolder As Boolean) As Boolean
'
'Rename the folder/file
'
Dim lFso As New FileSystemObject
  On Error GoTo Cerr
  If pFolder Then
    lFso.DeleteFolder pPath, True
  Else
    lFso.DeleteFile pPath, True
  End If
  S105_Delete = True
  Exit Function
Cerr:
  S105_Delete = False
End Function

Public Sub FormDrag(TheForm As Object)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub

Public Function GetVirtualPath(ByVal pVirtualDir As String, ByVal pSourcePath As String, Optional ByVal pPagePath As String) As String
Dim lStr As String
Dim lTmp As String
Dim Lcnt As Integer
Dim li As Integer
  On Error Resume Next
  If InStr(1, pSourcePath, pVirtualDir, vbTextCompare) > 0 Then
    lStr = pSourcePath
    lStr = Replace(pSourcePath, pVirtualDir, "", , , vbTextCompare)
    If pPagePath <> "" Then 'Get the ..path
      lTmp = pPagePath
      lTmp = Replace(pPagePath, pVirtualDir, "")
      Lcnt = UBound(Split(lTmp, "\"))
'      Lstr = Replace(Lstr, "\", "/")
      lTmp = ""
      For li = 1 To Lcnt - 1
        lTmp = lTmp & "..\"
      Next
      lStr = lTmp & lStr
      If Left(lStr, 1) = "\" Then lStr = Mid(lStr, 2)
    End If
    lStr = Replace(lStr, "//", "/")
    GetVirtualPath = lStr
  End If
End Function

Public Function S106_Read_File(ByVal pFilename As String) As String
Rem ---------------------------------
Rem Read the file content
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
Dim lTmp As String
Dim lPath As Variant
Dim li As Integer
  On Error GoTo Cerr
  Set lFso = New Scripting.FileSystemObject
  If lFso.FileExists(pFilename) Then
    S106_Read_File = lFso.OpenTextFile(pFilename).ReadAll
  End If
  Exit Function
Cerr:
  S106_Read_File = ""
End Function

Public Function S107_Folder_Exist(ByVal pName As String) As Boolean
Rem ---------------------------------
Rem Check folder
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
  On Error GoTo Cerr
  Set lFso = New Scripting.FileSystemObject
  S107_Folder_Exist = lFso.FolderExists(pName)
  Exit Function
Cerr:
  S107_Folder_Exist = False
End Function

Public Function S108_Copy_File(ByVal pSource As String, ByVal pDest As String, Optional ByVal pMove As Boolean) As Boolean
Rem ---------------------------------
Rem Copy file
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
Dim Lok As Boolean
  On Error GoTo Cerr
  Set lFso = New Scripting.FileSystemObject
  If S102_File_Exists(pDest) Then
    If MsgBox("Destination already exists. Do you want replace?", vbQuestion + vbOKCancel, Mtitle) = vbOK Then
      Lok = True
    Else
      Lok = False
    End If
  Else
    Lok = True
  End If
  If Lok Then
    If pMove Then
      lFso.MoveFile pSource, pDest
    Else
      lFso.CopyFile pSource, pDest, True
    End If
  End If
  S108_Copy_File = True
  Exit Function
Cerr:
  S108_Copy_File = False
End Function

Public Function S109_Copy_Folder(ByVal pSource As String, ByVal pDest As String, Optional ByVal pMove As Boolean) As Boolean
Rem ---------------------------------
Rem Copy Folder
Rem ---------------------------------

Dim lFso As Scripting.FileSystemObject
Dim Lok As Boolean
  On Error GoTo Cerr
  Set lFso = New Scripting.FileSystemObject
  If S102_File_Exists(pDest) Then
    If MsgBox("Destination already exists. Do you want replace?", vbQuestion + vbOKCancel, Mtitle) = vbOK Then
      Lok = True
    Else
      Lok = False
    End If
  Else
    Lok = True
  End If
  If Lok Then
    If pMove Then
      lFso.MoveFolder pSource, pDest
    Else
      lFso.CopyFolder pSource, pDest, True
    End If
  End If
  S109_Copy_Folder = True
  Exit Function
Cerr:
  S109_Copy_Folder = False
End Function

Public Sub S110_WriteLog(ByVal pDescription As String, Optional ByVal pDelete As Boolean)
'
'Write logs for tracing purpose
'
Dim lFn As Integer
  If pDelete Then S105_Delete App.Path & "\Tracing.dat"
  lFn = FreeFile
  Open App.Path & "\Tracing.dat" For Append As #lFn
  Print #lFn, "----------------------------"
  Print #lFn, "Time: " & Now
  Print #lFn, "Desc: " & pDescription
  Print #lFn, "----------------------------"
  Print #lFn,
  Close #lFn
End Sub

Public Function fnNumKey(ByVal KAscii As Integer) As Integer
   If (KAscii > 47 And KAscii < 58) Or Chr(KAscii) = "-" Or Chr(KAscii) = "." Or KAscii = 8 Or KAscii = 13 Then
        fnNumKey = KAscii
   Else
        fnNumKey = 0
   End If
End Function

Public Sub OpenWebsite(strWebsite As String)
    If ShellExecute(&O0, "Open", strWebsite, vbNullString, vbNullString, 1) < 33 Then
    End If
    DoEvents
End Sub

