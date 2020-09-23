VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSites"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem --------------------------------
Rem       Collection of sites
Rem       -------------------
Rem     Format of site details:
Rem  Name^Path^Url^UseFTP^Server^User^Password^Port^Connectionname1~string1^Connectionname2~string2^Connectionname3~string3^...
Rem --------------------------------

Rem --------------------------------
Rem       Private Declarations
Rem --------------------------------
Private mSites As Scripting.Dictionary

Rem --------------------------------
Rem        Load/Save Methods
Rem --------------------------------
Public Sub Load()
'
'Load the site details from site.dat
'
Dim lSite As clsSite
Dim lContent As String
Dim lConnection As New Collection
Dim lSites As Variant
Dim lDetails As Variant
Dim lConnections As Variant
Dim li As Integer
Dim lj As Integer
  On Error Resume Next
  Set mSites = New Scripting.Dictionary
  If S102_File_Exists(App.Path & "\Sites.dat") Then
    lContent = S106_Read_File(App.Path & "\Sites.dat")
    If lContent <> "" Then
      lSites = Split(lContent, vbCrLf)
      For li = LBound(lSites) To UBound(lSites)
        If InStr(lSites(li), "^") > 0 Then
          lDetails = Split(lSites(li), "^")
          Set lSite = New clsSite
          lSite.Name = lDetails(0)
          lSite.LocalPath = lDetails(1)
          lSite.URL = lDetails(2)
          lSite.UseFTP = CBool(val(lDetails(3)))
          lSite.Server = lDetails(4)
          lSite.Username = lDetails(5)
          lSite.Password = lDetails(6)
          lSite.Port = lDetails(7)
          If lDetails(8) <> "" Then
            lConnections = Split(lDetails(8), "^")
            For lj = LBound(lConnections) To UBound(lConnections)
              lConnection.Add lConnections(lj)
            Next
            lSite.ConnectionString = lConnection
          End If
          mSites.Add lSite.Name, lSite
          Set lSite = Nothing
        End If
      Next
    End If
  End If
End Sub

Public Sub Save()
'
'Save the site details...
'
Dim lSite As clsSite
Dim lContent As String
Dim lConnection As String
Dim lDetail As String
Dim li As Variant
Dim lj As Integer
Dim lFN As Integer
  lContent = ""
  If Not mSites Is Nothing Then
    For Each li In mSites.Keys
      lDetail = ""
      Set lSite = mSites.Item(li)
      If Not lSite Is Nothing Then
        With lSite
          'Get connection string
          lConnection = ""
          If Not .ConnectionString Is Nothing Then
            For lj = 1 To .ConnectionString.Count
              lConnection = lConnection & .ConnectionString(lj) & "^"
            Next
            If Right(lConnection, 1) = "^" Then lConnection = Left(lConnection, Len(lConnection) - 1)
          End If
          'Get other details
          lDetail = .Name & "^" & _
                    .LocalPath & "^" & _
                    .URL & "^" & _
                    CInt(.UseFTP) & "^" & _
                    .Server & "^" & _
                    .Username & "^" & _
                    .Password & "^" & _
                    .Port & "^" & _
                    lConnection
          lContent = lContent & lDetail & vbCrLf
        End With
      End If
    Next
    If S102_File_Exists(App.Path & "\Sites.dat") Then Kill App.Path & "\Sites.dat"
    lFN = FreeFile
    Open App.Path & "\Sites.dat" For Output As #lFN
    Print #lFN, lContent
    Close #lFN
  End If
End Sub

Rem ----------------------------------
Rem        Add/Remove Methods
Rem ----------------------------------

Public Function Newsite() As clsSite
'
'New site
'
Dim lSite As New clsSite
Dim lConn As New Collection
  lSite.ConnectionString = lConn
  Set Newsite = lSite
End Function

Public Sub Add(ByVal Site As clsSite)
'
'Add the site to the collection
'
  If Not Site Is Nothing Then
    If mSites.Exists(Site.Name) Then
      mSites.Remove Site.Name
    End If
    mSites.Add Site.Name, Site
  End If
End Sub

Public Sub Remove(Key)
'
'Remove the site
'
  If mSites.Exists(Key) Then
    mSites.Remove Key
  End If
End Sub

Rem ----------------------------------
Rem           Properties
Rem ----------------------------------

Public Property Get Count() As Integer
  Count = mSites.Count
End Property

Public Function Item(Key) As clsSite
'
'Get the item
'
  If mSites.Exists(Key) Then
    Set Item = mSites.Item(Key)
  End If
End Function

Public Function Exists(Key) As Boolean
'
'Check for item existence
'
  Exists = mSites.Exists(Key)
End Function

Public Function IDs() As Variant
'
'Get the IDs collection as array
'
  IDs = mSites.Keys
End Function

