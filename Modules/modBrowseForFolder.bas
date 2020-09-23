Attribute VB_Name = "modBrowseForFolder"
'Browse For Folder API Call Version 2.0, By Max Raskin 01 June 2000

'Required Module: modCheck (modCheck.bas)
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

'New Stuff in version 2.0:
'More helpfull comments
'StartUp Special Folder (optional) - browseforfolder with a special folder only
'Enum for the Flags of the BrowseForFolder API function
Enum BrowseForFolderFlags
    ReturnFileSystemFoldersOnly = &H1
    DontGoBelowDomain = &H2
    IncludeStatusText = &H4
    BrowseForComputer = &H1000
    BrowseForPrinter = &H2000
    BrowseIncludeFiles = &H4000
    IncludeTextBox = &H10
    ReturnFileSystemAncestors = &H8
End Enum

'BrowseInfo is a type used with the SHBrowseForFolder API call
Private Type BrowseInfo
     hOwner As Long
     pidlroot As Long
     pszDisplayName As Long
     lpszTitle As Long
     ulFlags As Long
     lpfnCallback As Long
     lParam As Long
     iImage As Long
End Type

Dim pidlroot As Long
'SHBrowseForFolder - Gets the Browse For Folder Dialog
'SHGetPathFromIDList - Converts ID List (pidl) to String

Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Private Declare Function SHGetFolderLocation Lib "shell32" (hwnd As Long, nFolder As Long, hToken As Long, dwReserved As Long, ppidl As Long) As Long

'lstrcat API function appends a string to another - that means that some API functions
'need their string in the numeric way like this does, so its kind of converts strings
'to numbers
Private Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long



Public Function BrowseForFolder(hwnd As Long, Optional Title As String, Optional Flags As BrowseForFolderFlags, Optional StartUpSpecialFolder As Folders) As String
Dim fs As New FileSystemObject
    'Variables for use:
     Dim iNull As Integer
     Dim IDList As Long
     Dim Result As Long
     Dim Path As String
     Dim bi As BrowseInfo
     Dim Ret As String
     Dim lFolder As Folder
     If Flags = 0 Then Flags = BIF_RETURNONLYFSDIRS
     
    'Type Settings
     With bi
'        Ret = fs.FolderExists("d:\srimax")  'Check if the special folder exists
'        If Ret = "True" Then Set lFolder = fs.GetFolder("d:\srimax")
'        If Ret <> "" Then .pidlroot = fs.GetFolder("d:\srimax")   'If there is any valid ID use it
        .hOwner = hWndOwner 'Set Owner Window
        .ulFlags = Flags 'Set flags (if any)
       .lpszTitle = lstrcat(Title, Chr(0)) 'Append title string to a long value
     End With

    'Execute the BrowseForFolder shell API and display the dialog
     IDList = SHBrowseForFolder(bi) 'Return ID List (selected path in a long value)
     
    'Get the info out of the dialog
     If IDList Then
        Path = String$(300, 0)
        Result = SHGetPathFromIDList(IDList, Path) 'Convert ID list to a string
        iNull = InStr(Path, vbNullChar) 'Get the position of the null character
        If iNull Then Path = Left$(Path, iNull - 1) 'Remove the null character
     End If

    'If Cancel button was clicked, error occured or Non File System Folder was selected then Path = ""
     BrowseForFolder = Path
     'Set fs = Nothing
End Function

