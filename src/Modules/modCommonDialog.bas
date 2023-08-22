Attribute VB_Name = "modCommonDialog"
Option Explicit

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
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Public Const OFN_READONLY = &H1
Public Const OFN_OVERWRITEPROMPT = &H2
Public Const OFN_HIDEREADONLY = &H4
Public Const OFN_NOCHANGEDIR = &H8
Public Const OFN_SHOWHELP = &H10
Public Const OFN_ENABLEHOOK = &H20
Public Const OFN_ENABLETEMPLATE = &H40
Public Const OFN_ENABLETEMPLATEHANDLE = &H80
Public Const OFN_NOVALIDATE = &H100
Public Const OFN_ALLOWMULTISELECT = &H200
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_PATHMUSTEXIST = &H800
Public Const OFN_FILEMUSTEXIST = &H1000
Public Const OFN_CREATEPROMPT = &H2000
Public Const OFN_SHAREAWARE = &H4000
Public Const OFN_NOREADONLYRETURN = &H8000
Public Const OFN_NOTESTFILECREATE = &H10000
Public Const OFN_NONETWORKBUTTON = &H20000
Public Const OFN_NOLONGNAMES = &H40000                      '  force no long names for 4.x modules
Public Const OFN_EXPLORER = &H80000                         '  new look commdlg
Public Const OFN_NODEREFERENCELINKS = &H100000
Public Const OFN_LONGNAMES = &H200000                       '  force long names for 3.x modules

Public Const OFN_SHAREFALLTHROUGH = 2
Public Const OFN_SHARENOWARN = 1
Public Const OFN_SHAREWARN = 0


Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

'<Constant>----------------------------------------------------------------
Public Enum SegmentEnum
  afwFilePath = 1
  afwFileName = 2
  afwFileExtension = 3
  afwPathNameSansExtension = 4
  afwFullFileName = 5
End Enum

'<EndConstant>-------------------------------------------------------------

Function SegmentFileName(ByVal FileName As String, Action As SegmentEnum) As String
    Dim CharPos     As Integer
    Dim LenFileName As Integer
On Error GoTo There
        
    LenFileName = Len(FileName)
    Select Case Action
        Case afwFilePath
            For CharPos = LenFileName To 1 Step -1
            If (Mid(FileName, CharPos, 1) = "\") Then
                SegmentFileName = Left(FileName, CharPos)
                Exit Function
            End If
            Next
        Case afwFileName, afwFullFileName
            For CharPos = LenFileName To 1 Step -1
            If (Mid(FileName, CharPos, 1) = "\") Then
                FileName = Mid(FileName, CharPos + 1)
                Exit For
            End If
            Next
            If (Action = afwFileName) Then
                CharPos = InStr(FileName, ".")
                If (CharPos > 0) Then FileName = Left(FileName, CharPos - 1)
            End If
            SegmentFileName = FileName
        Case afwFileExtension
            CharPos = InStr(FileName, ".")
            If (CharPos > 0) Then FileName = Mid(FileName, CharPos + 1)
            SegmentFileName = FileName
        Case afwPathNameSansExtension
            CharPos = InStr(FileName, ".")
            If (CharPos > 0) Then FileName = Left(FileName, CharPos - 1)
            SegmentFileName = FileName
    End Select

Exit Function
There:
    If Err Then

    End If
    Err.Clear
End Function
Function SaveDialog(vhWnd As Long, Filter As String, Title As String, InitDir As String, Optional Flags As Long = OFN_HIDEREADONLY Or OFN_OVERWRITEPROMPT Or OFN_CREATEPROMPT) As String
Dim ofn As OPENFILENAME
Dim a As Long
On Error GoTo There

  ofn.lStructSize = Len(ofn)
  ofn.hwndOwner = vhWnd
  ofn.hInstance = App.hInstance
  If Right$(Filter, 1) <> "|" Then Filter = Filter + "|"
  For a = 1 To Len(Filter)
    If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
  Next
  ofn.lpstrFilter = Filter
  ofn.lpstrFile = Space$(254)
  ofn.nMaxFile = 255
  ofn.lpstrFileTitle = Space$(254)
  ofn.nMaxFileTitle = 255
  ofn.lpstrInitialDir = InitDir
  ofn.lpstrTitle = Title
  ofn.Flags = Flags
  a = GetSaveFileName(ofn)
  If (a) Then
    SaveDialog = Trim$(ofn.lpstrFile)
  Else
    Err.Raise vbObjectError + 5656, , "Cancel selected"
    SaveDialog = ""
  End If

Exit Function
There:
    If Err Then

    End If
    Err.Clear
End Function
Function OpenDialog(hForm As Long, Filter As String, Title As String, InitDir As String, Optional Flags As Long = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST) As String
    Dim ofn As OPENFILENAME
    Dim a As Long

    ofn.lStructSize = Len(ofn)
    ofn.hwndOwner = hForm
    ofn.hInstance = App.hInstance
    If Right$(Filter, 1) <> "|" Then Filter = Filter & "|"
    For a = 1 To Len(Filter)
        If Mid$(Filter, a, 1) = "|" Then Mid$(Filter, a, 1) = Chr$(0)
    Next
    ofn.lpstrFilter = Filter
    ofn.lpstrFile = Space$(254)
    ofn.nMaxFile = 255
    ofn.lpstrFileTitle = Space$(254)
    ofn.nMaxFileTitle = 255
    ofn.lpstrInitialDir = InitDir
    ofn.lpstrTitle = Title
    'If Not IsMissing(Flags) Then
    '  If TypeName(Flags) = "Long" Then ofn.Flags = Flags
    'Else
    '  ofn.Flags = OFN_HIDEREADONLY Or OFN_FILEMUSTEXIST
    'End If
    ofn.Flags = Flags
    a = GetOpenFileName(ofn)
    If (a) Then
        Dim sss As String
        sss = Trim$(ofn.lpstrFile)
        If Right$(Trim$(sss), 1) = vbNullChar Then
            OpenDialog = Left$(Trim$(sss), Len(sss) - 1)
        Else
            OpenDialog = Trim$(sss)
        End If
    Else
'        Err.Raise vbObjectError + 6, , "Cancel Selected"
        OpenDialog = ""
    End If

End Function
