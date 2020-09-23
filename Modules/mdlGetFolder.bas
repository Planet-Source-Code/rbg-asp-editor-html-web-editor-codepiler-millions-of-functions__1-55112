Attribute VB_Name = "mdlGetFolder"
Option Explicit

Private Type SHITEMID
    SHItem As Long
    itemID() As Byte
End Type
Private Type ITEMIDLIST
    shellID As SHITEMID
End Type

Const SF_DESKTOP = &H0
Const SF_PROGRAMS = &H2
Const SF_MYDOCS = &H5
Const SF_FAVORITES = &H6     ' 98+
Const SF_STARTUP = &H7
Const SF_RECENT = &H8
Const SF_SENDTO = &H9
Const SF_STARTMENU = &HB
Const SF_MYMUSIC = &HD       ' Me+
Const SF_DESKTOP2 = &H10
Const SF_NETHOOD = &H13
Const SF_FONTS = &H14
Const SF_SHELLNEW = &H15
Const SF_STARTUP2 = &H18
Const SF_ALLUSERSDESK = &H19
Const SF_APPDATA = &H1A
Const SF_PRINTHOOD = &H1B
Const SF_APPDATA2 = &H1C
Const SF_TEMPINETFILES = &H20
Const SF_COOKIES = &H21
Const SF_HISTORY = &H22
Const SF_ALLUSERSAPPDATA = &H23
Const SF_WINDOWS = &H24
Const SF_WINSYSTEM = &H25
Const SF_PROGFILES = &H26
Const SF_MYPICS = &H27       ' Me+
Const SF_USERDIR = &H28
Const SF_WINSYSTEM2 = &H29
Const SF_COMMON = &H2B


Enum FOLDERTYPE
  ftSYSTEM
  ftWINDOWSFOLDER
  ftPROGRAMS
  ftMYDOCUMENTS
  ftFAVORITES
  ftSTARTUP
  ftRECENT
  ftSENDTO
  ftSTARTMENU
  ftMYMUSIC
  ftDESKTOP
  ftNETHOOD
  ftFONTSFOLDER
  ftTEMPLATES
  ftALLUSERS_STARTUP
  ftALLUSERS_DESKTOP
  ftAPPLICATION_DATA
  ftPRINT_HOOD
  ftLOCALSETTING_APP_DATA
  ftTEMPORARY_INTERNET_FILES
  ftCOOKIES
  ftHISTORY
  ftALLUSERS_APPDATA
  ftWINSYSTEM
  ftWINDOWS
  ftPROGRAM_FILES
  ftMY_PICTURES
  ftUSER_DIRECTORY
  ftWINDOWS_SYSTEM
  ftCOMMON_FILES
End Enum


'TO GET THE SYSTEM FOLDER
'**********************************************
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'**********************************************

'TO GET THE WINDOWS DIRECTORY
'***************************************
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
'***************************************

Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwnd As Long, ByVal folderid As Long, shidl As ITEMIDLIST) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal shidl As Long, ByVal shPath As String) As Long



Private Function getSpecialFolder(whichFolder As Long, ByVal hwnd As Long) As String
    Dim Path As String * 256
    Dim myid As ITEMIDLIST
    Dim rval As Long


    rval = SHGetSpecialFolderLocation(hwnd, whichFolder, myid)

    
    If rval = 0 Then ' If success
      rval = SHGetPathFromIDList(ByVal myid.shellID.SHItem, ByVal Path)
        If rval Then ' If True
        getSpecialFolder = Left(Path, InStr(Path, Chr(0)) - 1)
        End If
    End If
    
End Function


Public Function GetFolder(pFolderType As FOLDERTYPE, ByVal hwnd As Long)
Select Case pFolderType
' TO GET THE SYSTEM FOLDER
'*************************************************
Case 0
Dim SYSTEMFOLDER As String * 256
GetSystemDirectory SYSTEMFOLDER, 256
GetFolder = Left(SYSTEMFOLDER, InStr(SYSTEMFOLDER, Chr(0)) - 1)
'*************************************************


'TO GET THE WINDOWS FOLDER
'*************************************************
Case 1
Dim WINDOWSFOLDER As String * 256
GetWindowsDirectory WINDOWSFOLDER, 256
GetFolder = Left(WINDOWSFOLDER, InStr(WINDOWSFOLDER, Chr(0)) - 1)
'*************************************************



'FROM THIS ITEM TO THE END OF THE CASE SELECT USING
' THE DECLARTION FUNCTION SHGetSpecialFolderLocation
' YOU CAN SEE THAT THE "WINDOWS" FOLDER CAN BE GOT _
  USING 2 DECLATION .
Case 2
GetFolder = getSpecialFolder(&H2, hwnd) 'PROGRAMS
Case 3
GetFolder = getSpecialFolder(&H5, hwnd) 'MY DOCUMENTS
Case 4
GetFolder = getSpecialFolder(&H6, hwnd) 'FAVORITES
Case 5
GetFolder = getSpecialFolder(&H7, hwnd) 'STARTUP
Case 6
GetFolder = getSpecialFolder(&H8, hwnd) 'RECENT
Case 7
GetFolder = getSpecialFolder(&H9, hwnd) 'SEND TO
Case 8
GetFolder = getSpecialFolder(&HB, hwnd) 'START MENU
Case 9
GetFolder = getSpecialFolder(&HD, hwnd) ' MY MUSIC
Case 10
GetFolder = getSpecialFolder(&H10, hwnd) 'DESKTOP
Case 11
GetFolder = getSpecialFolder(&H13, hwnd) 'NETHOOD
Case 12
GetFolder = getSpecialFolder(&H14, hwnd) 'FONTS
Case 13
GetFolder = getSpecialFolder(&H15, hwnd) 'Templates
Case 14
GetFolder = getSpecialFolder(&H18, hwnd) 'ALL USERS START UP
Case 15
GetFolder = getSpecialFolder(&H19, hwnd) 'ALL USERS DESKTOP
Case 16
GetFolder = getSpecialFolder(&H1A, hwnd) 'APPLICATION DATA
Case 17
GetFolder = getSpecialFolder(&H1B, hwnd) 'PRINT HOOD
Case 18
GetFolder = getSpecialFolder(&H1C, hwnd) 'LOCAL SEETING APP'S DATA
Case 19
GetFolder = getSpecialFolder(&H20, hwnd) ' TEMPORARY INTERNET FILES
Case 20
GetFolder = getSpecialFolder(&H21, hwnd) 'COOKIES
Case 21
GetFolder = getSpecialFolder(&H22, hwnd) 'HISTORY
Case 22
GetFolder = getSpecialFolder(&H23, hwnd) 'APP DATA FOR ALL USERS
Case 23
GetFolder = getSpecialFolder(&H24, hwnd) 'WINDOWS
Case 24
GetFolder = getSpecialFolder(&H25, hwnd) 'WINSYSTEM
Case 25
GetFolder = getSpecialFolder(&H26, hwnd) 'PROGRAM FILES
Case 26
GetFolder = getSpecialFolder(&H27, hwnd) 'MY PICTURES
Case 27
GetFolder = getSpecialFolder(&H28, hwnd) 'USER DIRECTORY
Case 28
GetFolder = getSpecialFolder(&H29, hwnd) 'WINDOWS SYSTEM
Case 29
GetFolder = getSpecialFolder(&H2B, hwnd) 'COMMON FILES


End Select

End Function
