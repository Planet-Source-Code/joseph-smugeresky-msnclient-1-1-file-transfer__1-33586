Attribute VB_Name = "modMSN"
Option Explicit

'Var for editing privacy prefs
Public blnPrivacyEdit As Boolean

Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_EXPLORER = &H80000
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_READONLY = &H1
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0
Private Const OFN_SHOWHELP = &H10

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
    
    

Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long



Function SaveDialog(Form1 As Form, Filter As String, Title As String, InitDir As String) As String
    
    Dim OFN As OPENFILENAME
    Dim A As Long
    OFN.lStructSize = Len(OFN)
    OFN.hwndOwner = Form1.hWnd
    OFN.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"


    For A = 1 To Len(Filter)
        If Mid$(Filter, A, 1) = "|" Then Mid$(Filter, A, 1) = Chr$(0)
    Next
    OFN.lpstrFilter = Filter
    OFN.lpstrFile = Space$(254)
    OFN.nMaxFile = 255
    OFN.lpstrFileTitle = Space$(254)
    OFN.nMaxFileTitle = 255
    OFN.lpstrInitialDir = InitDir
    OFN.lpstrTitle = Title
    OFN.flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT
    A = GetSaveFileName(OFN)


    If (A) Then
        SaveDialog = Trim$(OFN.lpstrFile)
    Else
        SaveDialog = ""
    End If
End Function


Public Function ShowOpen(Frm As Form, Optional StartDirectory As String = "", Optional FileFilter As String = "", Optional WindowTitle As String = "Select A File") As String
    
    Dim OFN As OPENFILENAME, lngResult As Long
    
    OFN.lStructSize = Len(OFN)
    OFN.hwndOwner = Frm.hWnd
    OFN.hInstance = App.hInstance
    OFN.lpstrFilter = FileFilter
    OFN.lpstrFile = Space(255)
    OFN.nMaxFile = 256
    OFN.lpstrFileTitle = Space(255)
    OFN.nMaxFileTitle = 256
    OFN.lpstrInitialDir = StartDirectory
    OFN.lpstrTitle = WindowTitle
    OFN.flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    lngResult = GetOpenFileName(OFN)
   
    If lngResult = 1 Then
        ShowOpen = Trim(OFN.lpstrFile)
    Else
        ShowOpen = ""
    End If
    
End Function
