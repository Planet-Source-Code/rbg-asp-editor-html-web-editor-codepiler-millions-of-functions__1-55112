VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmFormValidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form Validation"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5355
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
   Icon            =   "frmFormValidation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDisplayName 
      Height          =   315
      Left            =   2085
      TabIndex        =   2
      Top             =   3060
      Width           =   2610
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   300
      Left            =   3750
      TabIndex        =   11
      Top             =   5310
      Width           =   930
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   300
      Left            =   2655
      TabIndex        =   10
      Top             =   5310
      Width           =   930
   End
   Begin VB.TextBox txtTo 
      Height          =   315
      Left            =   4245
      TabIndex        =   9
      Top             =   4695
      Width           =   435
   End
   Begin VB.TextBox txtFrom 
      Height          =   315
      Left            =   3480
      TabIndex        =   8
      Top             =   4695
      Width           =   435
   End
   Begin VB.OptionButton optRange 
      Caption         =   "Number from"
      Height          =   300
      Left            =   2085
      TabIndex        =   7
      Top             =   4695
      Width           =   1305
   End
   Begin VB.OptionButton optNumber 
      Caption         =   "Number"
      Height          =   300
      Left            =   960
      TabIndex        =   6
      Top             =   4695
      Width           =   1050
   End
   Begin VB.OptionButton optEmail 
      Caption         =   "Email"
      Height          =   300
      Left            =   2085
      TabIndex        =   5
      Top             =   4290
      Width           =   1050
   End
   Begin VB.OptionButton optAnything 
      Caption         =   "Any thing"
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   4290
      Width           =   1050
   End
   Begin VB.CheckBox chkRequired 
      Caption         =   "Required"
      Height          =   240
      Left            =   1545
      TabIndex        =   3
      Top             =   3600
      Width           =   990
   End
   Begin VB.ComboBox cboForms 
      Height          =   315
      Left            =   1710
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   810
      Width           =   3015
   End
   Begin MSComctlLib.ListView lsvFields 
      Height          =   1395
      Left            =   960
      TabIndex        =   1
      Top             =   1500
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2461
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
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   5909
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "displayname"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "required"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "type"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "range"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   525
      Picture         =   "frmFormValidation.frx":000C
      Top             =   165
      Width           =   255
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Form Validation"
      Height          =   195
      Left            =   855
      TabIndex        =   18
      Top             =   210
      Width           =   1095
   End
   Begin VB.Line Line1 
      X1              =   420
      X2              =   8090
      Y1              =   585
      Y2              =   585
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   375
      X2              =   8090
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Display Name:"
      Height          =   195
      Left            =   960
      TabIndex        =   17
      Top             =   3120
      Width           =   1020
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "to"
      Height          =   195
      Left            =   4020
      TabIndex        =   16
      Top             =   4755
      Width           =   150
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type:"
      Height          =   195
      Left            =   960
      TabIndex        =   15
      Top             =   3975
      Width           =   420
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Value:"
      Height          =   195
      Left            =   960
      TabIndex        =   14
      Top             =   3600
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fields:"
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   1230
      Width           =   465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Forms:"
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   870
      Width           =   495
   End
End
Attribute VB_Name = "frmFormValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mFormsCollection As New Collection
Private mAllow As Boolean

Private Sub cboForms_Click()
Dim lInputs As String
Dim Litem As ListItem
Dim lNames As Variant
Dim li As Integer
  lInputs = mFormsCollection.Item(cboForms.ListIndex + 1)
  lInputs = Split(lInputs, "||")(0)
  lsvFields.ListItems.Clear
  If lInputs <> "" Then
    lNames = Split(lInputs, "^")
    For li = LBound(lNames) To UBound(lNames)
      If lNames(li) <> "" Then
        Set Litem = lsvFields.ListItems.Add(, , lNames(li))
        Litem.ListSubItems.Add , , lNames(li)
      End If
    Next
    If lsvFields.ListItems.Count > 0 Then
      cmdOk.Enabled = True
      lsvFields_ItemClick lsvFields.ListItems(1)
    Else
      cmdOk.Enabled = False
    End If
  End If
End Sub



Private Sub chkRequired_Click()
  If Not lsvFields.SelectedItem Is Nothing Then
    If mAllow Then lsvFields.SelectedItem.SubItems(2) = IIf(chkRequired.Value = 1, "r", "")
  End If
End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub cmdOk_Click()
  GenerateFormValidation
  Unload Me
End Sub


Private Sub Form_Load()
  Screen.MousePointer = vbHourglass
  GetFormsList
  Screen.MousePointer = vbDefault
End Sub

Private Sub lsvFields_ItemClick(ByVal Item As MSComctlLib.ListItem)
  If Not Item Is Nothing Then
    ClearValues
    mAllow = False
    txtDisplayName.Text = Item.SubItems(1)
    chkRequired.Value = IIf(Item.SubItems(2) = "r", 1, 0)
    Select Case LCase(Item.SubItems(3))
    Case "n"
      optNumber.Value = True
    Case "e"
      optEmail.Value = True
    Case "r"
      optRange.Value = True
    Case Else
      optAnything.Value = True
    End Select
    If Item.SubItems(4) <> "" Then
      txtFrom.Text = Split(Item.SubItems(4), ":")(0)
      txtTo.Text = Split(Item.SubItems(4), ":")(1)
    End If
    mAllow = True
  End If
End Sub



Private Sub optAnything_Click()
  If Not lsvFields.SelectedItem Is Nothing Then
    If mAllow Then
      lsvFields.SelectedItem.SubItems(3) = ""
      lsvFields.SelectedItem.SubItems(4) = ""
    End If
  End If
End Sub

Private Sub optEmail_Click()
  If Not lsvFields.SelectedItem Is Nothing Then
    If mAllow Then
      lsvFields.SelectedItem.SubItems(3) = "e"
      lsvFields.SelectedItem.SubItems(4) = ""
    End If
  End If
End Sub

Private Sub optNumber_Click()
  If Not lsvFields.SelectedItem Is Nothing Then
    If mAllow Then
      lsvFields.SelectedItem.SubItems(3) = "n"
      lsvFields.SelectedItem.SubItems(4) = ""
    End If
  End If
End Sub

Private Sub optRange_Click()
  If Not lsvFields.SelectedItem Is Nothing Then
    If mAllow Then lsvFields.SelectedItem.SubItems(3) = "r"
  End If
End Sub



Private Sub txtDisplayName_Change()
  If Not lsvFields.SelectedItem Is Nothing Then
    If mAllow Then lsvFields.SelectedItem.SubItems(1) = txtDisplayName.Text
  End If
End Sub

Private Sub txtFrom_Change()
  If Not lsvFields.SelectedItem Is Nothing Then
    If optRange.Value = True Then
      If mAllow Then lsvFields.SelectedItem.SubItems(4) = txtFrom.Text & ":" & txtTo.Text
    End If
  End If
End Sub

Private Sub txtTo_Change()
  If Not lsvFields.SelectedItem Is Nothing Then
    If optRange.Value = True Then
      If mAllow Then lsvFields.SelectedItem.SubItems(4) = txtFrom.Text & ":" & txtTo.Text
    End If
  End If
End Sub

'
'User Functions
'
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
        lFormStr = mID(lStr, lStart, Lpos2 - lStart)
        Lpos2 = InStr(lStart, LCase(lStr), ">")
        If Lpos2 > 0 Then
          lFormTag = mID(lStr, lStart, Lpos2 - lStart)
          lFormTag = TrancateInput(lFormTag)
          If InStr(1, lFormTag, "name=", vbTextCompare) > 0 Then
            lPos3 = InStr(1, lFormTag, "name=", vbTextCompare) + 5
            Lpos2 = InStr(lPos3, lFormTag, " ")
            If Lpos2 = 0 Then Lpos2 = Len(lFormTag) + 1
            If Lpos2 > 0 Then
              lName = mID(lFormTag, lPos3, Lpos2 - lPos3)
              lName = Replace(lName, """", "")
            End If
          End If
        End If
        cboForms.AddItem IIf(lName = "", "form[" & lCount & "]", lName)
        lInputs = GetInputsList(lFormStr, lCount)
        mFormsCollection.Add lInputs
        lCount = lCount + 1
      Else
        Lpos1 = -1
      End If
    End If
  Loop Until lStart <= 0
  If cboForms.ListCount > 0 Then cboForms.ListIndex = 0
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
        lInput = mID(pForm, Lpos1, Lpos2 - Lpos1)
        If lInput <> "" Then
          lInput = TrancateInput(lInput)
          If InStr(LCase(lInput), "type=") > 0 Then
            lPos3 = InStr(LCase(lInput), "type=") + 5
            Lpos2 = InStr(lPos3, LCase(lInput), " ")
            If Lpos2 = 0 Then Lpos2 = Len(lInput) + 1
            If Lpos2 > 0 Then
              lName = mID(lInput, lPos3, Lpos2 - lPos3)
              lName = Replace(lName, """", "")
              If LCase(lName) = "text" Or LCase(lName) = "list" Then
                If InStr(LCase(lInput), "name=") > 0 Then
                  lPos3 = InStr(LCase(lInput), "name=") + 5
                  Lpos2 = InStr(lPos3, LCase(lInput), " ")
                  If Lpos2 = 0 Then Lpos2 = Len(lInput) + 1
                  If Lpos2 > 0 Then
                    lValue = mID(lInput, lPos3, Lpos2 - lPos3)
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

Private Function ClearValues()
'
'Clear all values before move/enter the values
'
  mAllow = False
  txtDisplayName.Text = ""
  chkRequired.Value = 0
  optAnything.Value = True
  txtFrom.Text = ""
  txtTo.Text = ""
  mAllow = True
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

Private Function GenerateFormValidation()
'
'Generate the script for form validation
'
Dim lValidation As String
Dim lFunction As String
Dim lArguments As String
Dim lStr As String
Dim lScript As String
Dim lCloseScript As String
Dim li As Long
  lValidation = vbCrLf & vbTab & "function cP_Formvalidation()" & vbCrLf & vbTab & "{ " & vbCrLf & vbTab & vbTab & "var li,field,pos,dispname;" & vbCrLf & _
                vbTab & vbTab & "var val,type,err,args;" & vbCrLf & vbTab & vbTab & "args=cP_Formvalidation.arguments;" & vbCrLf & vbTab & vbTab & "err='';" & vbCrLf & _
                vbTab & vbTab & "for (li=1; li<(args.length-2); li+=3) " & vbCrLf & vbTab & vbTab & "{ " & vbCrLf & vbTab & vbTab & vbTab & "type=args[li+2];  " & vbCrLf & _
                vbTab & vbTab & vbTab & "field=document.forms[args[0]][args[li]];" & vbCrLf & vbTab & vbTab & vbTab & "if (field) " & vbCrLf & vbTab & vbTab & vbTab & "{ " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & "dispname=field.name; " & vbCrLf & vbTab & vbTab & vbTab & vbTab & "if (args[li+1]!="""") dispname=args[li+1]; " & vbCrLf & vbTab & vbTab & vbTab & vbTab & "if ((val=field.value)!="""") " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & "{ " & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & "switch(type.substring(0,2))" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & "{" & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "case 're':" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "pos=value.indexOf('@'); " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "if (pos<1 || pos==(val.length-1)) err+='""' + dispname + '"" should be an e-mail address.\n'; " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "break;" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "case 'rn':" & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "if (isNaN(val)) err += '""' + dispname + '"" should be a number.\n'; " & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "break;" & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "case 'rr':" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "if (isNaN(val)) err += '""' + dispname + '"" should be a number.\n';" & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "pos=type.indexOf(':'); " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "if (val<type.substring(2,pos) || type.substring(pos+1)<val) err += '""' + dispname + '"" should be a number between ' + type.substring(2,pos) + ' and ' + type.substring(pos+1) +'.\n'; " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & "break;" & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & "}" & vbCrLf & vbTab & vbTab & vbTab & vbTab & "} " & vbCrLf & _
                vbTab & vbTab & vbTab & vbTab & "else " & vbCrLf & vbTab & vbTab & vbTab & vbTab & vbTab & "if (type.charAt(0) == 'r') err += '""' + dispname + '"" is required.\n'; " & vbCrLf & vbTab & vbTab & vbTab & "} " & vbCrLf & _
                vbTab & vbTab & "}" & vbCrLf & vbTab & vbTab & "if (err) alert(err);" & vbCrLf & vbTab & vbTab & "document.CP_returnValue = (err == '');" & vbCrLf & vbTab & "}" & vbCrLf
                
  lScript = vbCrLf & "<script language=""JavaScript"">" & vbCrLf & vbCrLf & vbTab & "<!-- Code Piler Generated" & vbCrLf
  lCloseScript = vbCrLf & vbTab & "-->" & vbCrLf & "</script>" & vbCrLf
  If lsvFields.ListItems.Count > 0 Then
    'generate the arguments
    lArguments = ""
    For li = 1 To lsvFields.ListItems.Count
      If lsvFields.ListItems(li).SubItems(2) <> "" Or lsvFields.ListItems(li).SubItems(3) <> "" Or lsvFields.ListItems(li).SubItems(4) <> "" Then
        lArguments = lArguments & "'" & lsvFields.ListItems(li).Text & "','" & IIf(lsvFields.ListItems(li).SubItems(1) = lsvFields.ListItems(li).Text, "", lsvFields.ListItems(li).SubItems(1)) & "','" & lsvFields.ListItems(li).SubItems(2) & lsvFields.ListItems(li).SubItems(3) & lsvFields.ListItems(li).SubItems(4) & "',"
      End If
    Next
    If Right(lArguments, 1) = "," Then lArguments = Left(lArguments, Len(lArguments) - 1)
    'generate the function calling
    lFunction = "cP_Formvalidation(" & cboForms.ListIndex & "," & lArguments & ");return document.CP_returnValue;"
    'find the position to insert the script
    lStr = frmEditor.RTB(frmEditor.GetActiveRTB).Text
    If InStr(1, lStr, "function cp_formvalidation()", vbTextCompare) = 0 Then
      If InStr(1, lStr, "<!-- code piler generated", vbTextCompare) > 0 Then
          frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, lStr, "<!-- code piler generated", vbTextCompare) + 25
          frmEditor.RTB(frmEditor.GetActiveRTB).Paste lValidation
      ElseIf InStr(1, LCase(lStr), "<head>") > 0 Then
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<head>") + 6
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lValidation & lCloseScript
      ElseIf InStr(1, LCase(lStr), "<html>") > 0 Then
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = InStr(1, LCase(lStr), "<head>") + 6
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lValidation & lCloseScript
      Else
        frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = 1
        frmEditor.RTB(frmEditor.GetActiveRTB).Paste lScript & lValidation & lCloseScript
      End If
    End If
    'include the function calling
    InsertFunctionCall "form", cboForms.ListIndex + 1, "cP_Formvalidation(", lFunction, "onSubmit"
  End If
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
    lLine = mID(lStr, lTagPos, InStr(lTagPos, lStr, vbCrLf) - lTagPos)
  Else
    lLine = mID(lStr, lTagPos)
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
        lChar = mID(pTagLine, li, 1)
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
    Lpos2 = InStr(Lpos1, lStr, "returnValue;")
    If Lpos2 > 0 Then
      Lpos2 = Lpos2 + 12
      frmEditor.RTB(frmEditor.GetActiveRTB).SelStart = Lpos1 - 1
      frmEditor.RTB(frmEditor.GetActiveRTB).SelLength = Lpos2 - Lpos1
      frmEditor.RTB(frmEditor.GetActiveRTB).SelText = ""
    End If
  End If
End Function

