VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem --------------------------------
Rem       Collection of history files
Rem       ---------------------------
Rem Format of history details:
Rem  Date^Path
Rem --------------------------------

Rem --------------------------------
Rem       Private Declarations
Rem --------------------------------
Private Mhistories As Scripting.Dictionary

Rem --------------------------------
Rem        Load/Save Methods
Rem --------------------------------
Public Sub Load()
'
'Load the History details from History.dat
'
Dim lHistory As clsHistory
Dim lHistories As Variant
Dim lDetails As Variant
Dim lContent As String
Dim li As Integer
  On Error Resume Next
  Set Mhistories = New Scripting.Dictionary
  If S102_File_Exists(App.Path & "\History.dat") Then
    lContent = S106_Read_File(App.Path & "\History.dat")
    If lContent <> "" Then
      lHistories = Split(lContent, vbCrLf)
      For li = LBound(lHistories) To UBound(lHistories)
        If InStr(lHistories(li), "^") > 0 Then
          lDetails = Split(lHistories(li), "^")
          Set lHistory = New clsHistory
          lHistory.HDate = lDetails(0)
          lHistory.LocalPath = lDetails(1)
          Mhistories.Add lHistory.HDate & lHistory.LocalPath, lHistory
          Set lHistory = Nothing
        End If
      Next
    End If
  End If
End Sub

Public Sub Save()
'
'Save the History details...
'
Dim lHistory As clsHistory
Dim lContent As String
Dim lConnection As String
Dim lDetail As String
Dim li As Variant
Dim lj As Integer
Dim lFN As Integer
  lContent = ""
  If Not Mhistories Is Nothing Then
    For Each li In Mhistories.Keys
      lDetail = ""
      Set lHistory = Mhistories.Item(li)
      If Not lHistory Is Nothing Then
        With lHistory
          'Get other details
          lDetail = .HDate & "^" & .LocalPath
          lContent = lContent & lDetail & vbCrLf
        End With
      End If
    Next
    If S102_File_Exists(App.Path & "\History.dat") Then Kill App.Path & "\History.dat"
    lFN = FreeFile
    Open App.Path & "\History.dat" For Output As #lFN
    Print #lFN, lContent
    Close #lFN
  End If
End Sub

Rem ----------------------------------
Rem        Add/Remove Methods
Rem ----------------------------------

Public Sub Add(ByVal History As clsHistory)
'
'Add the History to the collection
'
  If Not History Is Nothing Then
    If Mhistories.Exists(History.HDate & History.LocalPath) Then
      Mhistories.Remove History.HDate & History.LocalPath
    End If
    Mhistories.Add History.HDate & History.LocalPath, History
  End If
End Sub

Public Sub Remove(Key)
'
'Remove the History
'
  If Mhistories.Exists(Key) Then
    Mhistories.Remove Key
  End If
End Sub

Public Sub RemoveFor(HDate As String)
'
'Remove the History for the date
'
Dim lHistory As clsHistory
Dim LIDS As Variant
Dim li As Long
  LIDS = IDs
  For li = LBound(LIDS) To UBound(LIDS)
    Set lHistory = Mhistories.Item(LIDS(li))
    If Not lHistory Is Nothing Then
      If lHistory.HDate = HDate Then
        Remove lHistory.HDate & lHistory.LocalPath
      End If
    End If
  Next
End Sub

Public Sub RemoveAll()
'
'Remove the History for the date
'
  Mhistories.RemoveAll
End Sub

Rem ----------------------------------
Rem           Properties
Rem ----------------------------------

Public Property Get Count() As Integer
  Count = Mhistories.Count
End Property

Public Function Item(Key) As clsHistory
'
'Get the item
'
  If Mhistories.Exists(Key) Then
    Set Item = Mhistories.Item(Key)
  End If
End Function

Public Function Exists(Key) As Boolean
'
'Check for item existence
'
  Exists = Mhistories.Exists(Key)
End Function

Public Function IDs() As Variant
'
'Get the IDs collection as array
'
  IDs = Mhistories.Keys
End Function



