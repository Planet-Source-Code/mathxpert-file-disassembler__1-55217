VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Disassembler"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Español"
      Height          =   375
      Left            =   2273
      TabIndex        =   8
      ToolTipText     =   "Idioma"
      Top             =   1200
      Width           =   2055
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      ToolTipText     =   "Browse"
      Top             =   120
      Width           =   375
   End
   Begin VB.CommandButton Command2 
      Caption         =   "..."
      Height          =   285
      Left            =   6120
      TabIndex        =   6
      ToolTipText     =   "Save As"
      Top             =   480
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2040
      TabIndex        =   5
      Top             =   480
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Disassemble"
      Height          =   285
      Left            =   5400
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2040
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   6120
      Top             =   0
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Output file:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "File to disassemble:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const APPTITLE = "File Disassembler"
Private Const MSG = " [Disassembling - please wait]"

Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Dim OutFN As String
Dim DisasmFN As String
Dim FileLen1 As Long
Dim FileLen2 As Long
Dim bExit As Boolean
Dim bUsed As Boolean
Dim i As Long

Private Sub ChangeLanguage()
With Command4
    If .Caption = "Español" Then
        .ToolTipText = "Idioma"
        Label1.Caption = "File to disassemble:"
        Label2.Caption = "Output file:"
        Command1.Caption = "Disassemble"
        Command2.ToolTipText = "Save As"
        Command3.ToolTipText = "Browse"
    Else
        .ToolTipText = "Language"
        Label1.Caption = "Archivo para desmontar:"
        Label2.Caption = "Archivo de salida:"
        Command1.Caption = "Desmontar"
        Command2.ToolTipText = "Guardar Como"
        Command3.ToolTipText = "Hojear"
    End If
End With
End Sub

Private Function GetPercentage(Num As Long, TotalNum As Long) As Single
GetPercentage = Int((Num / TotalNum) * 100)
End Function

Private Function GetShortFileName(ByVal FullPath As String) As String
On Error Resume Next
Dim lAns As Long
Dim sAns As String
Dim iLen As Integer

If Dir(FullPath) = "" Then Exit Function

sAns = Space(255)
lAns = GetShortPathName(FullPath, sAns, 255)
GetShortFileName = Left(sAns, lAns)
End Function

Private Sub Command1_Click()
On Error Resume Next

Text1.SetFocus
If Text1 = "" Then Exit Sub
If Text2 = "" Then Exit Sub
If Dir(Text1) = "" Then Exit Sub

DisasmFN = GetShortFileName(Text1)
OutFN = Text2

Open OutFN For Output As #1
Close #1

If bExit Then bExit = False
If bUsed Then bUsed = False
If i > 0 Then i = 0
Timer1.Enabled = True
Shell "link /dump /disasm " & DisasmFN & " /out:" & OutFN, vbHide
End Sub

Private Sub Command2_Click()
Dim tmpOutFN As String
Text2.SetFocus
With Command4
    tmpOutFN = ShowSaveDlg(Me, IIf(.Caption = "English", "Archivo de Texto", "Text File") & " (*.txt)|*.txt", IIf(.Caption = "English", "Guardar Como", "Save As"), Environ("PROGRAMFILES"), "*.txt")
End With
If tmpOutFN <> "" Then Text2 = tmpOutFN
End Sub

Private Sub Command3_Click()
Dim tmpDisasmFN As String
Text1.SetFocus
With Command4
    tmpDisasmFN = ShowOpenDlg(Me, IIf(.Caption = "English", "Tipos Apoyados", "Supported Types") & " (*.exe, *.dll, *.ocx)|*.exe;*.dll;*.ocx", IIf(.Caption = "English", "Elegir archivo", "Choose file"), Environ("PROGRAMFILES"))
End With
If tmpDisasmFN <> "" Then Text1 = tmpDisasmFN
End Sub

Private Sub Command4_Click()
Text1.SetFocus
With Command4
    If .Caption = "English" Then .Caption = "Español" Else .Caption = "English"
End With
ChangeLanguage
End Sub

Private Sub Form_Load()
Command4.Caption = GetSetting(App.Title, "Settings", "OppLang", "Español")
ChangeLanguage
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
SaveSetting App.Title, "Settings", "OppLang", Command4.Caption
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then KeyAscii = 0: Command1_Click
End Sub

Private Sub Timer1_Timer()
FileLen2 = FileLen(OutFN)

If bUsed Then
    If FileLen1 = FileLen2 Then
        i = i + 1
    Else
        If i > 0 Then i = 0
    End If
End If
If i > 200 Then bExit = True: GoTo ExitHnd

If FileLen(OutFN) >= FileLen(DisasmFN) * 9.86 Then
ExitHnd:
    Timer1.Enabled = False
    If Caption = APPTITLE & MSG Then Caption = APPTITLE
    If Not Command1.Enabled Then Command1.Enabled = True
    If ProgressBar1.Value <> 0 Then ProgressBar1.Value = 0
    If bExit Then Exit Sub
Else
    If Caption = APPTITLE Then Caption = APPTITLE & MSG
    If Command1.Enabled Then Command1.Enabled = False
    If ProgressBar1.Value <> GetPercentage(FileLen(OutFN), FileLen(DisasmFN) * 9.86) Then ProgressBar1.Value = GetPercentage(FileLen(OutFN), FileLen(DisasmFN) * 9.86)
End If

FileLen1 = FileLen(OutFN)
If Not bUsed Then bUsed = True
End Sub
