VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem --------------------------------
Rem          History details
Rem --------------------------------

Rem --------------------------------
Rem     Properties Declarations
Rem --------------------------------
Private mDate As String
Private mLocalpath As String

Rem --------------------------------
Rem           Properties
Rem --------------------------------

Public Property Get HDate() As String
  HDate = mDate
End Property

Public Property Let HDate(ByVal Value As String)
  mDate = Value
End Property

Public Property Get LocalPath() As String
  LocalPath = mLocalpath
End Property

Public Property Let LocalPath(ByVal Value As String)
  mLocalpath = Value
End Property
