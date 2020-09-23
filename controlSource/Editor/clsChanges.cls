VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsChanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem -----------------------------
Rem Collection of changes in rtb for Undo/Redo
Rem -----------------------------

Rem -----------------------------
Rem          Properties
Rem -----------------------------
Private WithEvents RTB As Richtextbox
Attribute RTB.VB_VarHelpID = -1
Private mLock As Boolean
Private mUndoOk As Boolean 'check for start the undo
Private mRedoOk As Boolean 'check for start the redo

Rem -----------------------------
Rem      Private Declrations
Rem -----------------------------
Private Declare Function LockWindowUpdate Lib "User32" (ByVal hwndLock As Long) As Long
Private mUndos As New Collection 'collection of changes
Private mBackspace As Boolean 'to consider the continuous backspacekey as a single
Private mDelkey As Boolean 'to consider the continuous deletekey as a single
Private mDoes As Integer 'count of undos/redos
Private mRecover As Boolean 'check for first undo after change

Rem -----------------------------
Rem          Constants
Rem -----------------------------
Private Const mRange = 51 'Range of collection volume

Rem -----------------------------
Rem          Properties
Rem -----------------------------
Public Property Let Richtextbox(ByVal Value As Richtextbox)
  Set RTB = Value
End Property

Public Property Let LockRTB(ByVal Value As Boolean)
  mLock = Value
End Property

Public Property Let CanUndo(ByVal Value As Boolean)
  mUndoOk = Value
End Property

Public Property Get CanUndo() As Boolean
  CanUndo = mUndoOk
End Property

Public Property Get CanRedo() As Boolean
  CanRedo = mRedoOk
End Property

Private Sub Class_Initialize()
  mRecover = True
End Sub

Rem -----------------------------
Rem           Events
Rem -----------------------------
Public Sub RTB_KeyDown(KeyCode As Integer, Shift As Integer)
Dim lFormat As New clsFormat
  If Not mLock Then
    If KeyCode = vbKeyReturn Then 'Return key is pressed
      Add
      KeyCode = 0
    ElseIf KeyCode = vbKeyBack Then 'Backspace key is presses
      If mBackspace = False Then
        Add
        mBackspace = True
      End If
    ElseIf KeyCode = vbKeyDelete Then 'Delete key is presses
      If mDelkey = False Then
        Add
        mDelkey = True
      End If
    ElseIf KeyCode = vbKeyV And Shift = 2 Then 'Control+V is pressed
      Add
      KeyCode = 0
    ElseIf KeyCode = vbKeyX And Shift = 2 Then 'Control+V is pressed
      Add
    ElseIf KeyCode = vbKeyZ And Shift = 2 Then 'Control+Z is pressed
      Undo
      KeyCode = 0
    ElseIf KeyCode = vbKeyY And Shift = 2 Then 'Control+Y is pressed
      Redo
      KeyCode = 0
    ElseIf Shift = 0 Then 'Other than above
      If mRecover Then Add
      mBackspace = False
      mDelkey = False
      mRecover = False
    End If
    If KeyCode = vbKeyTab Or KeyCode = vbKeyF3 Then
      KeyCode = 0
    End If
  End If
End Sub

Rem -----------------------------
Rem           Methods
Rem -----------------------------

Public Function Undo() As Boolean
'
'Undo the changes
'
Dim lFormat As New clsFormat
  If mRecover = False Then Add 'add last change
  If mUndos.Count > 0 Then
    If mDoes >= mRange - 1 Or mDoes >= mUndos.Count - 1 Then
      mUndoOk = False
      Undo = False
      Exit Function
    End If
    Set lFormat = mUndos(mUndos.Count)
    mUndos.Add lFormat, , 1 'Add on top
    mUndos.Remove mUndos.Count 'Remove the Last
    Set lFormat = mUndos(mUndos.Count)
    LockWindowUpdate RTB.hWnd
    RTB.TextRTF = lFormat.TextRTF
    RTB.SelStart = lFormat.SelStart
    RTB.SelLength = lFormat.SelLength
    LockWindowUpdate 0&
    mDoes = mDoes + 1
    mRecover = True
    mRedoOk = True
    Undo = True
  End If
End Function

Public Function Redo() As Boolean
'
'Redo the undo changes
'
Dim lFormat As New clsFormat
  If mUndos.Count > 0 Then
    If mDoes <= 0 Then
      mRedoOk = False
      Redo = False
      Exit Function
    End If
    Set lFormat = mUndos(1)
    LockWindowUpdate RTB.hWnd
    RTB.TextRTF = lFormat.TextRTF
    RTB.SelStart = lFormat.SelStart
    RTB.SelLength = lFormat.SelLength
    LockWindowUpdate 0&
    mDoes = mDoes - 1
    mRecover = True
    mUndos.Remove 1 'Remove the first
    mUndos.Add lFormat 'Add on bottom
    mUndoOk = True
    Redo = True
  End If
End Function

Private Sub Add()
'
'Add the changes to collections
'
Dim lFormat As New clsFormat
Dim lCount As Integer
  lCount = mUndos.Count + 1
  lFormat.SelStart = RTB.SelStart
  lFormat.TextRTF = RTB.TextRTF
  lFormat.SelLength = RTB.SelLength
  lFormat.Index = lCount
  mUndos.Add lFormat
  mDoes = 0
  mUndoOk = True
  If mRange + 1 <= mUndos.Count Then mUndos.Remove 1 'Remove the top if collection has more than range
End Sub
