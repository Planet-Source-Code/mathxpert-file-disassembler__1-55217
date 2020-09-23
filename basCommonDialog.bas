Attribute VB_Name = "basCommonDialog"
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

Private Const OFN_READONLY = &H1
Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_NOCHANGEDIR = &H8
Private Const OFN_SHOWHELP = &H10
Private Const OFN_ENABLEHOOK = &H20
Private Const OFN_ENABLETEMPLATE = &H40
Private Const OFN_ENABLETEMPLATEHANDLE = &H80
Private Const OFN_NOVALIDATE = &H100
Private Const OFN_ALLOWMULTISELECT = &H200
Private Const OFN_EXTENSIONDIFFERENT = &H400
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_FILEMUSTEXIST = &H1000
Private Const OFN_CREATEPROMPT = &H2000
Private Const OFN_SHAREAWARE = &H4000
Private Const OFN_NOREADONLYRETURN = &H8000
Private Const OFN_NOTESTFILECREATE = &H10000
Private Const OFN_NONETWORKBUTTON = &H20000
Private Const OFN_NOLONGNAMES = &H40000
Private Const OFN_EXPLORER = &H80000
Private Const OFN_NODEREFERENCELINKS = &H100000
Private Const OFN_LONGNAMES = &H200000
Private Const OFN_SHAREFALLTHROUGH = 2
Private Const OFN_SHARENOWARN = 1
Private Const OFN_SHAREWARN = 0

Private Declare Function GetSaveFileName Lib "comdlg32.dll" _
  Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) _
  As Long
Private Declare Function GetOpenFileName Lib "comdlg32.dll" _
  Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) _
  As Long

Public Function ShowOpenDlg(F As Form, strFilter As String, _
  strTitel As String, Optional strInitDir As String) As String

  Dim lngOpenFileName As OPENFILENAME
  Dim lngAnt As Long

  With lngOpenFileName
    .lStructSize = Len(lngOpenFileName)
    .hwndOwner = F.hWnd
    .hInstance = App.hInstance
    If Right$(strFilter, 1) <> "|" Then _
      strFilter = strFilter & "|"
    
    For lngAnt = 1 To Len(strFilter)
      If Mid$(strFilter, lngAnt, 1) = "|" Then _
       Mid$(strFilter, lngAnt, 1) = Chr$(0)
    Next

    .lpstrFilter = strFilter
    .lpstrFile = Space$(254)
    .nMaxFile = 255
    .lpstrFileTitle = Space$(254)
    .nMaxFileTitle = 255
    .lpstrInitialDir = IIf(IsMissing(strInitDir), vbNullString, strInitDir)
    .lpstrTitle = strTitel
    .flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST

    lngAnt = GetOpenFileName(lngOpenFileName)
    If (lngAnt) Then
      ShowOpenDlg = Trim$(.lpstrFile)
    Else
      ShowOpenDlg = ""
    End If
  End With
End Function

Public Function ShowSaveDlg(F As Form, strFilter As String, _
  strTitel As String, Optional strInitDir As String, _
  Optional strDefExt As String) As String

  Dim lngOpenFileName As OPENFILENAME
  Dim lngAnt As Long
  Dim lngPnt As Long
  
  With lngOpenFileName
    .lStructSize = Len(lngOpenFileName)
    .hwndOwner = F.hWnd
    .hInstance = App.hInstance
    If Right$(strFilter, 1) <> "|" Then _
      strFilter = strFilter & "|"

    For lngAnt = 1 To Len(strFilter)
      If Mid$(strFilter, lngAnt, 1) = "|" Then _
        Mid$(strFilter, lngAnt, 1) = Chr$(0)
    Next
    
    .lpstrFilter = strFilter
    .lpstrFile = Space$(254)
    .lpstrDefExt = IIf(IsMissing(strDefExt), vbNullString, strDefExt)
    .nMaxFile = 255
    .lpstrFileTitle = Space$(254)
    .nMaxFileTitle = 255
    .lpstrInitialDir = IIf(IsMissing(strInitDir), vbNullString, strInitDir)
    .lpstrTitle = strTitel
    .flags = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or _
      OFN_CREATEPROMPT

    lngAnt = GetSaveFileName(lngOpenFileName)
    If (lngAnt) Then
      ShowSaveDlg = Trim$(.lpstrFile)
    Else
      ShowSaveDlg = ""
    End If
  End With
End Function


