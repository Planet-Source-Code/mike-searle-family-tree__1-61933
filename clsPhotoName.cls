VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPhotoName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarIndId As Long 'local copy
Private mvarY1 As Long 'local copy
Private mvarY2 As Long 'local copy
Private mvarX1 As Long 'local copy
Private mvarX2 As Long 'local copy
Private mvarNote As String 'local copy

Public Property Let Note(ByVal vData As String)
    mvarNote = vData
End Property


Public Property Get Note() As String
    Note = mvarNote
End Property

Public Property Let X2(ByVal vData As Long)
    mvarX2 = vData
End Property

Public Property Get X2() As Long
    X2 = mvarX2
End Property

Public Property Let X1(ByVal vData As Long)
    mvarX1 = vData
End Property

Public Property Get X1() As Long
    X1 = mvarX1
End Property

Public Property Let Y2(ByVal vData As Long)
    mvarY2 = vData
End Property

Public Property Get Y2() As Long
    Y2 = mvarY2
End Property

Public Property Let Y1(ByVal vData As Long)
    mvarY1 = vData
End Property

Public Property Get Y1() As Long
    Y1 = mvarY1
End Property

Public Property Let IndId(ByVal vData As Long)
    mvarIndId = vData
End Property

Public Property Get IndId() As Long
    IndId = mvarIndId
End Property
