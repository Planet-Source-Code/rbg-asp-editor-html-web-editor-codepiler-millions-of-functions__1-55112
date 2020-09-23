Attribute VB_Name = "modXPTheme"
'This module will also work with the
'"Microsoft Windows Common Controls 5.0 (SP2)" component,
'Which comes with VS 6 Enterprise (Not sure about other versions).

'This code creates the xml manifest for using windows
'xp theme controls. The InitXP sub should be called
'in the Form_Initialize event of each form, befor
'the form is shown. This only works when the program
'is an executible, not in the IDE. Also, The First Time
'The exe is run, it will not work. Creating a sub
'program that is run first may do the trick . . .

'Note: Command buttons and option buttons do not seem
'to display properly when placed inside of a frame control.
'Frame controls INSIDE of another frame control also have
'the same effect

'Although I have not tested it on other systems, there
'should be no effect . . .

'Declare all variables
Option Explicit

'Upgrades controls
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long

Public Sub InitXP()
        
    Dim strPath As String   'Path to the manifest file
    Dim strData As String   'XML data for the manifest file
    Dim FF As Integer       'FreeFile handle
    
    'Skip errors
    On Error Resume Next
    
    'Get the path to the exe. The manifest has the same name, with the manifest extension
    strPath = App.Path & IIf(Right(App.Path, 1) = "\", vbNullString, "\")
    strPath = strPath & App.EXEName & IIf(LCase(Right(App.EXEName, 4)) = ".exe", ".manifest", ".exe.manifest")
    
    'Check for the existance of the file, and skip to initializing if found
    If Dir(strPath, vbReadOnly Or vbSystem Or vbHidden) <> vbNullString Then GoTo InitControls
        
    'Set up the xml data
    strData = "<?xml version=" & Chr(34) & "1.0" & Chr(34) & " encoding=" & Chr(34) & "UTF-8" & Chr(34) & " standalone=" & Chr(34) & "yes" & Chr(34) & "?>" & vbCrLf
    strData = strData & "<assembly xmlns=" & Chr(34) & "urn:schemas-microsoft-com:asm.v1" & Chr(34) & " manifestVersion=" & Chr(34) & "1.0" & Chr(34) & ">" & vbCrLf
    strData = strData & "     <assemblyIdentity version=" & Chr(34) & "1.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " name=" & Chr(34) & "HybridDesign.WindowsXP.Example" & Chr(34) & " type=" & Chr(34) & "win32" & Chr(34) & " />" & vbCrLf
    strData = strData & "     <description>Windows XP Theme.</description>" & vbCrLf
    strData = strData & "     <dependency>" & vbCrLf
    strData = strData & "          <dependentAssembly>" & vbCrLf
    strData = strData & "               <assemblyIdentity type=" & Chr(34) & "win32" & Chr(34) & " name=" & Chr(34) & "Microsoft.Windows.Common-Controls" & Chr(34) & " version=" & Chr(34) & "6.0.0.0" & Chr(34) & " processorArchitecture=" & Chr(34) & "X86" & Chr(34) & " publicKeyToken=" & Chr(34) & "6595b64144ccf1df" & Chr(34) & " language=" & Chr(34) & "*" & Chr(34) & " />" & vbCrLf
    strData = strData & "          </dependentAssembly>" & vbCrLf
    strData = strData & "     </dependency>" & vbCrLf
    strData = strData & "</assembly>"
            
    'Open the file and print the xml data
    FF = FreeFile
    Open strPath For Output As #FF
        Print #FF, strData
    Close #FF
    
    'Set the atrributes of the file
    SetAttr strPath, vbHidden Or vbSystem Or vbReadOnly Or vbArchive
        
InitControls:
    
    'Call the api to initialize the XP theme
    Call InitCommonControls
    
End Sub
