VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSite"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem --------------------------------
Rem          Site details
Rem --------------------------------

Rem --------------------------------
Rem     Properties Declarations
Rem --------------------------------
Private mSitename As String
Private mLocalpath As String
Private mURL As String
Private mConnectionString As Collection
Private mRemote As Boolean
Private mServer As String
Private mUsername As String
Private mPassword As String
Private mPort As Integer

Rem --------------------------------
Rem           Properties
Rem --------------------------------

Public Property Get Name() As String
  Name = mSitename
End Property

Public Property Let Name(ByVal Value As String)
  mSitename = Value
End Property

Public Property Get LocalPath() As String
  LocalPath = mLocalpath
End Property

Public Property Let LocalPath(ByVal Value As String)
  mLocalpath = Value
End Property

Public Property Get URL() As String
  URL = mURL
End Property

Public Property Let URL(ByVal Value As String)
  mURL = Value
End Property

Public Property Get ConnectionString() As Collection
  Set ConnectionString = mConnectionString
End Property

Public Property Let ConnectionString(ByVal Value As Collection)
  Set mConnectionString = Value
End Property

Public Property Get UseFTP() As Boolean
  UseFTP = mRemote
End Property

Public Property Let UseFTP(ByVal Value As Boolean)
  mRemote = Value
End Property

Public Property Get Server() As String
  Server = mServer
End Property

Public Property Let Server(ByVal Value As String)
  mServer = Value
End Property

Public Property Get Username() As String
  Username = mUsername
End Property

Public Property Let Username(ByVal Value As String)
  mUsername = Value
End Property

Public Property Get Password() As String
  Password = mPassword
End Property

Public Property Let Password(ByVal Value As String)
  mPassword = Value
End Property

Public Property Get Port() As Integer
  Port = mPort
End Property

Public Property Let Port(ByVal Value As Integer)
  mPort = Value
End Property

