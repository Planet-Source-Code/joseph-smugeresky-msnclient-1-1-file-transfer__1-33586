VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTransfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Auth cookie
Private lngAuthCookie As Long

'File path
Private strFilePath As String

'File name
Private strFileName As String

'File size
Private lngFileSize As Long

'Connection index
Private intIndex As Integer

'Transfer Index
Private intXIndex As Integer

'Current File Index
Private intFIndex As Long

'Transfer Type
Public Enum XFRTYPE
    XFR_IN = 1
    XFR_OUT = 2
End Enum

Private XFerType As XFRTYPE
Public Property Get XCookie() As Long

    XCookie = lngAuthCookie

End Property

Public Property Let XCookie(ByVal Cookie As Long)

    lngAuthCookie = Cookie

End Property

Public Property Get XFilePath() As String

    XFilePath = strFilePath

End Property

Public Property Let XFilePath(ByVal FilePath As String)

    strFilePath = FilePath
    
End Property

Public Property Get XFileName() As String

    XFileName = strFileName

End Property

Public Property Let XFileName(ByVal FileName As String)

    strFileName = FileName

End Property

Public Property Get XFileSize() As Long

    XFileSize = lngFileSize

End Property

Public Property Let XFileSize(ByVal FileSize As Long)

    lngFileSize = FileSize

End Property

Public Property Get XIndex() As Integer

    XIndex = intIndex

End Property

Public Property Let XIndex(ByVal Index As Integer)

    intIndex = Index

End Property

Public Property Get XTFRIndex() As Integer

    XTFRIndex = intXIndex

End Property

Public Property Let XTFRIndex(ByVal Index As Integer)

    intXIndex = Index

End Property




Public Property Get XFileIndex() As Integer

    XFileIndex = intFIndex

End Property

Public Property Let XFileIndex(ByVal Fileindex As Integer)

    intFIndex = Fileindex

End Property



Public Property Get XType() As XFRTYPE

    XType = XFerType

End Property

Public Property Let XType(ByVal XFType As XFRTYPE)

    XFerType = XFType

End Property
