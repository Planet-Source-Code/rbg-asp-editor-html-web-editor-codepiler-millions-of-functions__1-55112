VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Rem -----------------------------
Rem     Format of Undo/Redo
Rem     -------------------
Rem SelStart
Rem TextRTF
Rem -----------------------------

Rem -----------------------------
Rem    Properties Declaration
Rem -----------------------------
Private mSelStart As Long
Private mTextRTF As String
Private mSelLength As Long
Private mIndex As Long

Rem -----------------------------
Rem          Properties
Rem -----------------------------

Public Property Get SelStart() As Long
  SelStart = mSelStart
End Property

Public Property Let SelStart(ByVal value As Long)
  mSelStart = value
End Property

Public Property Get TextRTF() As String
  TextRTF = mTextRTF
End Property

Public Property Let TextRTF(ByVal value As String)
  mTextRTF = value
End Property

Public Property Get SelLength() As Long
  SelLength = mSelLength
End Property

Public Property Let SelLength(ByVal value As Long)
  mSelLength = value
End Property

Public Property Get Index() As Long
  Index = mIndex
End Property

Public Property Let Index(ByVal value As Long)
  mIndex = value
End Property
