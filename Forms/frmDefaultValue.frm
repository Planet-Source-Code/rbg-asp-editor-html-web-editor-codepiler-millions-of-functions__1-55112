VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmDefaultValue 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Default Value"
   ClientHeight    =   4485
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5235
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDefaultValue.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOkD 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   300
      Left            =   2535
      TabIndex        =   3
      Top             =   3930
      Width           =   930
   End
   Begin VB.CommandButton cmdCancelD 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3615
      TabIndex        =   4
      Top             =   3930
      Width           =   930
   End
   Begin VB.ComboBox cboForms1 
      Height          =   315
      Left            =   1515
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   825
      Width           =   3015
   End
   Begin VB.TextBox txtDefaultValue 
      Height          =   315
      Left            =   1935
      TabIndex        =   2
      Top             =   3330
      Width           =   735
   End
   Begin MSComctlLib.ListView lsvList 
      Height          =   1545
      Left            =   810
      TabIndex        =   1
      Top             =   1635
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2725
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5909
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   390
      Picture         =   "frmDefaultValue.frx":000C
      Top             =   120
      Width           =   360
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   285
      X2              =   8000
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Line Line1 
      X1              =   330
      X2              =   8000
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Value"
      Height          =   195
      Left            =   870
      TabIndex        =   8
      Top             =   210
      Width           =   960
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Listbox:"
      Height          =   195
      Left            =   810
      TabIndex        =   7
      Top             =   1335
      Width           =   570
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Default Value:"
      Height          =   195
      Left            =   810
      TabIndex        =   6
      Top             =   3390
      Width           =   1020
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forms:"
      Height          =   195
      Left            =   810
      TabIndex        =   5
      Top             =   885
      Width           =   495
   End
End
Attribute VB_Name = "frmDefaultValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFormsCollection As New Collection
Private mDefaultValues() As String

Private Sub cboForms1_Click()
Dim lInputs As String
Dim Litem As ListItem
Dim lNames As Variant
Dim lValue As String
Dim lValues As String
Dim lFillvalues As Variant
Dim li As Integer
  On Error Resume Next
  'Save the values if change
  If val(cboForms1.Tag) <> cboForms1.ListIndex Then
    lValues = ""
    lValue = ""
    For li = 1 To lsvList.ListItems.Count
      If lsvList.ListItems(li).Tag <> "" Then
        lValue = cboForms1.Tag & ",'" & lsvList.ListItems(li).Text & "'," & lsvList.ListItems(li).Tag
      Else
        lValue = "^,^,^"
      End If
      lValues = lValues & lValue & ","
    Next
    If Right(lValues, 1) = "," Then lValues = Left(lValues, Len(lValues) - 1)
    mDefaultValues(val(cboForms1.Tag)) = lValues
    cboForms1.Tag = cboForms1.ListIndex
    txtDefaultValue.Text = ""
  End If
  'load the selected form
  lInputs = mFormsCollection.Item(cboForms1.ListIndex + 1)
  lInputs = Split(lInputs, "||")(1)
  lFillvalues = Split(mDefaultValues(cboForms1.ListIndex), ",")
  lsvList.ListItems.Clear
  If lInputs <> "" Then
    lNames = Split(lInputs, "^")
    For li = LBound(lNames) To UBound(lNames)
      If lNames(li) <> "" Then
        Set Litem = lsvList.ListItems.Add(, , lNames(li))
        Litem.Tag = IIf(lFillvalues((li * 3) + 2) <> "^", Replace(lFillvalues((li * 3) + 2), "'", ""), "")
      End If
    Next
    If lsvList.ListItems.Count > 0 Then
      cmdOkD.Enabled = True
      'lsvlist_ItemClick lsvList.ListItems(1)
    Else
      cmdOkD.Enabled = False
    End If
    If lsvList.ListItems.Count > 0 Then
      lsvList.Tag = 0
      lsvList_ItemClick lsvList.ListItems(1)
    End If
  End If
End Sub

Private Sub cmdCancelD_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
  GetFormsList
  Screen.MousePointer = vbDefault
End Sub

Private Sub cmdOkD_Click()
Dim lValue As String
Dim lValues As String
Dim li As Integer
  lValues = ""
  lValue = ""
  For li = 1 To lsvList.ListItems.Count
    If lsvList.ListItems(li).Tag <> "" Then
      lValue = cboForms1.Tag & ",'" & lsvList.ListItems(li).Text & "'," & lsvList.ListItems(li).Tag
    Else
      lValue = "^,^,^"
    End If
    lValues = lValues & lValue & ","
  Next
  If Right(lValues, 1) = "," Then lValues = Left(lValues, Len(lValues) - 1)
  mDefaultValues(val(cboForms1.Tag)) = lValues
  cboForms1.Tag = cboForms1.ListIndex
  txtDefaultValue.Text = ""
  
  GenerateDefaultValue
  Unload Me
End Sub

Private Sub lsvList_ItemClick(ByVal Item As MSComctlLib.ListItem)
Dim lValue As String
Dim lValues As Variant
Dim Litem As ListItem
Dim li As Long
  If Not Item Is Nothing Then
    If lsvList.Tag <> Item.Index And val(lsvList.Tag) > 0 Then
      lsvList.ListItems(val(lsvList.Tag)).Tag = txtDefaultValue.Text
      lsvList.Tag = Item.Index
      txtDefaultValue.Text = Item.Tag
    ElseIf lsvList.Tag <> Item.Index And val(lsvList.Tag) = 0 Then
      txtDefaultValue.Text = Item.Tag
      lsvList.Tag = Item.Index
    End If
  End If
End Sub

Private Sub txtDefaultValue_KeyUp(KeyCode As Integer, Shift As Integer)
  lsvList.ListItems(val(lsvList.Tag)).Tag = txtDefaultValue.Text
End Sub

Private Function GetFormsList()
'
'Load the forms in the pages
'
Dim lStr As String
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim lPos3 As Long
Dim lFormStr As String
Dim lFormTag As String
Dim lInputs As String
Dim lCount As Integer
Dim lStart As Long
Dim lName As String
  On Error Resume Next
  Screen.MousePointer = vbHourglass
  lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
  Lpos1 = 1
  Set mFormsCollection = New Collection
  Do
    lName = ""
    Lpos1 = InStr(lStart + 1, LCase(lStr), "<form ")
    If Lpos1 <= 0 Then Lpos1 = InStr(lStart + 1, LCase(lStr), "<form>")
    lStart = Lpos1
    If Lpos1 > 0 Then
      Lpos2 = InStr(lStart, LCase(lStr), "</form>")
      If Lpos2 > 0 Then
        lFormStr = Mid(lStr, lStart, Lpos2 - lStart)
        Lpos2 = InStr(lStart, LCase(lStr), ">")
        If Lpos2 > 0 Then
          lFormTag = Mid(lStr, lStart, Lpos2 - lStart)
          lFormTag = TrancateInput(lFormTag)
          If InStr(1, lFormTag, "name=", vbTextCompare) > 0 Then
            lPos3 = InStr(1, lFormTag, "name=", vbTextCompare) + 5
            Lpos2 = InStr(lPos3, lFormTag, " ")
            If Lpos2 = 0 Then Lpos2 = Len(lFormTag) + 1
            If Lpos2 > 0 Then
              lName = Mid(lFormTag, lPos3, Lpos2 - lPos3)
              lName = Replace(lName, """", "")
            End If
          End If
        End If
        cboForms1.AddItem IIf(lName = "", "form[" & lCount & "]", lName)
        lInputs = GetInputsList(lFormStr, lCount)
        mFormsCollection.Add lInputs
        lCount = lCount + 1
      Else
        Lpos1 = -1
      End If
    End If
  Loop Until lStart <= 0
  If cboForms1.ListCount > 0 Then
    ReDim mDefaultValues(cboForms1.ListCount - 1)
    cboForms1.Tag = 0
    cboForms1.ListIndex = 0
  End If
  Screen.MousePointer = vbDefault
End Function

Private Function GetInputsList(ByVal pForm As String, ByVal pFormIndex As Integer) As String
'
'Load all inputs (type=text)
'
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim lPos3 As Long
Dim lNames As String
Dim lName As String
Dim lValue As String
Dim lList As String
Dim lInput As String
Dim lForm As String
Dim lSelect As String
  Lpos1 = 1
  'lForm = pForm
  pForm = Replace(pForm, "<select", "<input type=list", , , vbTextCompare)
  Do
    Lpos1 = InStr(Lpos1 + 1, LCase(pForm), "<input ")
    If Lpos1 > 0 Then
      Lpos2 = InStr(Lpos1, pForm, ">")
      If Lpos2 > 0 Then
        lInput = Mid(pForm, Lpos1, Lpos2 - Lpos1)
        If lInput <> "" Then
          lInput = TrancateInput(lInput)
          If InStr(LCase(lInput), "type=") > 0 Then
            lPos3 = InStr(LCase(lInput), "type=") + 5
            Lpos2 = InStr(lPos3, LCase(lInput), " ")
            If Lpos2 = 0 Then Lpos2 = Len(lInput) + 1
            If Lpos2 > 0 Then
              lName = Mid(lInput, lPos3, Lpos2 - lPos3)
              lName = Replace(lName, """", "")
              If LCase(lName) = "text" Or LCase(lName) = "list" Then
                If InStr(LCase(lInput), "name=") > 0 Then
                  lPos3 = InStr(LCase(lInput), "name=") + 5
                  Lpos2 = InStr(lPos3, LCase(lInput), " ")
                  If Lpos2 = 0 Then Lpos2 = Len(lInput) + 1
                  If Lpos2 > 0 Then
                    lValue = Mid(lInput, lPos3, Lpos2 - lPos3)
                    lValue = Replace(lValue, """", "")
                    lNames = lNames & lValue & "^"
                    If lName = "list" Then
                      lList = lList & lValue & "^"
                    End If
                  End If
                End If
              End If
            End If
          End If
        End If
      End If
    Else
      Lpos1 = -1
    End If
  Loop Until Lpos1 <= 0
  If InStr(lNames, "^") > 0 Then lNames = Left(lNames, Len(lNames) - 1)
  If InStr(lList, "^") > 0 Then lList = Left(lList, Len(lList) - 1)
  GetInputsList = lNames & "||" & lList
End Function

Private Function GenerateDefaultValue()
'
'Generate the script to set default value
'
Dim lValidation As String
Dim lFunction As String
Dim lArguments As String
Dim lStr As String
Dim lScript As String
Dim lCloseScript As String
Dim li As Long
  lValidation = vbCrLf & vbCrLf & vbTab & _
                "function CP_defaultValue(){" & vbCrLf & vbTab & vbTab & _
                "var i,ctl,args=CP_defaultValue.arguments;" & vbCrLf & vbTab & vbTab & _
                "for (i=0; i<(args.length-2); i+=3) {ctl=document.forms[args[i]][args[i+1]];" & vbCrLf & vbTab & vbTab & _
                "if (ctl){ if (ctl.options(args[i+2])) ctl.options(args[i+2]).selected=true;} }" & vbCrLf & vbTab & _
                "}" & vbCrLf
  lScript = vbCrLf & "<script language=""JavaScript"">" & vbCrLf & vbCrLf & vbTab & "<!-- Code Piler Generated" & vbCrLf
  lCloseScript = vbCrLf & vbTab & "-->" & vbCrLf & "</script>" & vbCrLf
    'generate the arguments
    lArguments = ""
    For li = LBound(mDefaultValues) To UBound(mDefaultValues)
      lArguments = lArguments & mDefaultValues(li) & ","
    Next
    lArguments = Replace(lArguments, "^,^,^,", "")
    lArguments = Replace(lArguments, ",,", ",")
    If Right(lArguments, 1) = "," Then lArguments = Left(lArguments, Len(lArguments) - 1)
    'generate the function calling
    lFunction = "CP_defaultValue(" & lArguments & ");"
    'find the position to insert the script
    lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
    If InStr(1, lStr, "function cp_defaultvalue()", vbTextCompare) = 0 Then
      If InStr(1, lStr, "<!-- code piler generated", vbTextCompare) > 0 Then
          frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, lStr, "<!-- code piler generated", vbTextCompare) + 25
          frmEditor.RTB(frmEditor.GetActiveRTB).Paste lValidation
      ElseIf InStr(1, lStr, "<head>", vbTextCompare) > 0 Then
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<head>") + 6
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lValidation & lCloseScript
      ElseIf InStr(1, lStr, "<html>", vbTextCompare) > 0 Then
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<html>") + 6
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lValidation & lCloseScript
      Else
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = 1
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lValidation & lCloseScript
      End If
    End If
    'include the function calling
    InsertFunctionCall "body", 1, "CP_defaultValue(", lFunction, "onLoad"
End Function

Private Function InsertFunctionCall(ByVal pInsertTag As String, ByVal pNth As Integer, ByVal pFunctionName As String, ByVal pFunctionCall As String, ByVal pEvent As String)
'
'Insert the function call into the body tag
'
Dim lStr As String
Dim lStart As Long
Dim lPos As Long
Dim lTagPos As Long
Dim lLine As String
Dim lStatus As Integer
  lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
  lTagPos = FindTag(pInsertTag, pNth)
  If InStr(lTagPos, lStr, vbCrLf) > 0 Then
    lLine = Mid(lStr, lTagPos, InStr(lTagPos, lStr, vbCrLf) - lTagPos)
  Else
    lLine = Mid(lStr, lTagPos)
  End If
  lPos = lTagPos + FindEventCall(lLine, pEvent, lStatus) - 1
  RemoveFunctionCall pFunctionName, lTagPos
  frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = lPos - 1
  Select Case lStatus
  Case 1, 2 'Just include
    frmEditor.RTB(frmEditor.GetActiveRTB).SelText = pFunctionCall
  Case 3 'Add ""
    frmEditor.RTB(frmEditor.GetActiveRTB).SelText = "=""" & pFunctionCall & """"
  Case 4 'Insert with event
    frmEditor.RTB(frmEditor.GetActiveRTB).SelText = " " & pEvent & " = """ & pFunctionCall & """"
  End Select
End Function

Private Function FindTag(ByVal pTag As String, Optional ByVal pNth As Integer = 1) As Long
'
'Find the tag position
'
Dim li As Integer
Dim lStart As Long
Dim lPos As Long
Dim lStr As String
Dim lTagstart As String
Dim lTagclose As String
Dim lTagseparator As String
  lStart = 1
  lTagstart = "<"
  lTagclose = ">"
  lTagseparator = " "
  lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
  For li = 1 To pNth
    lPos = InStr(lStart, lStr, lTagstart & pTag & lTagseparator, vbTextCompare)
    If lPos <= 0 Then lPos = InStr(lStart, lStr, lTagstart & pTag & lTagclose, vbTextCompare)
    lStart = lPos + 1
  Next
  FindTag = lStart - 1
End Function

Private Function FindEventCall(ByVal pTagLine As String, ByVal pEvent As String, ByRef pStatus As Integer) As Long
'
'Find the event call position for insert function
'pStatus for as below
' 1-Insert after "
' 2-Insert after =
' 3-Insert after event (ie invalid eventcall) and also include =
' 4-Insert as fresh event
'
Dim li As Long
Dim lChar As String
Dim lStart As Long
Dim lPos As Long
Dim lTmp As String
  lTmp = pTagLine
  lPos = InStr(1, pTagLine, pEvent, vbTextCompare)
  lStart = lPos
  If lPos > 0 Then 'if event is present
    lPos = InStr(lPos, pTagLine, "=")
    If lPos > 0 Then
      lStart = lPos + 1
      For li = lStart To Len(pTagLine)
        lChar = Mid(pTagLine, li, 1)
        If lChar <> " " Then
          If lChar = """" Then
            FindEventCall = li + 1
            pStatus = 1
          Else
            FindEventCall = lPos + IIf(lStart < li, 2, 1)
            pStatus = 2
          End If
          Exit For
        End If
      Next
    Else 'if incomplete tag event is present
      FindEventCall = lStart + Len(pEvent)
      pStatus = 3
    End If
  Else 'if no event present
    pTagLine = RTrim(pTagLine)
    lPos = InStr(1, pTagLine, ">")
    FindEventCall = lPos
    pStatus = 4
  End If
End Function

Private Function RemoveFunctionCall(ByVal pFunctionName As String, ByVal pStart As Long)
'
'if function call already exists,remove it
'
Dim Lpos1 As Long
Dim Lpos2 As Long
Dim lStr As String
  lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
  Lpos1 = InStr(pStart, lStr, pFunctionName, vbTextCompare)
  If Lpos1 > 0 Then
    Lpos2 = InStr(Lpos1, lStr, ");")
    If Lpos2 > 0 Then
      Lpos2 = Lpos2 + 2
      frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = Lpos1 - 1
      frmEditor.RTB(frmEditor.GetActiveRTB).SelLength = Lpos2 - Lpos1
      frmEditor.RTB(frmEditor.GetActiveRTB).SelText = ""
    End If
  End If
End Function

Private Function TrancateInput(ByVal pInput As String) As String
'
'trancate the input line for spaces to get the text type/ and name
'
Dim lString As String
Dim lChar As String
Dim li As Integer
  If pInput <> "" Then
    pInput = Replace(pInput, vbCrLf, "")
    pInput = Replace(pInput, vbTab, "")
    'to trancate the unwanted spaces
    Do
      pInput = Replace(pInput, "  ", " ")
    Loop Until InStr(pInput, "  ") = 0
    pInput = Replace(pInput, " = ", "=")
    pInput = Replace(pInput, "= ", "=")
    pInput = Replace(pInput, " =", "=")
    TrancateInput = pInput
  End If
End Function
