VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   ClientHeight    =   5190
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   5880
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   5880
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTranslate2 
      Caption         =   "interpréter Français en Anglaise"
      BeginProperty Font 
         Name            =   "Paramount"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   12
      Top             =   2640
      Width           =   3255
   End
   Begin VB.TextBox Text3 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   10
      Top             =   1080
      Width           =   615
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   840
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCorrect 
      Caption         =   "Correction"
      Height          =   495
      Left            =   4680
      TabIndex        =   9
      Top             =   0
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   840
      Top             =   3480
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   4680
      TabIndex        =   4
      Top             =   1200
      Width           =   1095
   End
   Begin VB.CommandButton cmdNew 
      Caption         =   "New"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdTranslate 
      Caption         =   "interpréter anglaise en français"
      BeginProperty Font 
         Name            =   "Paramount"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2160
      Width           =   3255
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   3360
      Width           =   5295
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Simplified Arabic"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   945
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1080
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Mots Français"
      Height          =   255
      Left            =   360
      TabIndex        =   13
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "nombre du record"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   735
   End
   Begin VB.Label Banner 
      Height          =   255
      Left            =   480
      TabIndex        =   7
      Top             =   4920
      Width           =   4095
   End
   Begin VB.Label Label3 
      Caption         =   "Mots Arabe"
      Height          =   255
      Left            =   1080
      TabIndex        =   6
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "English/French Translator"
      BeginProperty Font 
         Name            =   "Paramount"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   120
      Width           =   3975
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dictionary As Translate
Dim FileNum As Integer
Dim RecordLen As Long
Dim CurrentRecord As Long
Dim LastRecord As Long
Dim X
Dim RecNum As Long
Dim Y As Long


Private Sub cmdClear_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
End Sub

Private Sub cmdCorrect_Click()
Dim MyNo
MyNo = InputBox("Write The Record No")
If MyNo <> "" Then
RecNum = MyNo
End If
Get #FileNum, RecNum, Dictionary
Text1.Text = Trim(Dictionary.Arabic)
Text2.Text = Trim(Dictionary.English)
Text3.Text = RecNum
CurrentRecord = MyNo
Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)
End Sub

Private Sub cmdSave_Click()
'Fill Dictionary with the currently Displayed Data
Dim confirm
confirm = MsgBox("Are You Sure?", vbOKCancel, "Confirm")
If confirm = vbOK Then
If Left(Text1.Text, 1) = "-" Then
Dictionary.Arabic = Text1.Text
Else
Dictionary.Arabic = "- " & Text1.Text
End If
Dictionary.English = Text2.Text

'Save Dictionary to the current record
Put #FileNum, CurrentRecord, Dictionary
'Close #FileNum
Else
Exit Sub
End If
End Sub
Public Sub SaveCurrentRecord()
'Fill Dictionary with the currently Displayed Data
Dictionary.Arabic = Text1.Text
Dictionary.English = Text2.Text
'Save Dictionary to the current record
Put #FileNum, CurrentRecord, Dictionary
End Sub

Private Sub cmdTranslate_Click()
Dim NameToSearch As String
Dim Found As Integer

NameToSearch = "- " & Text1.Text
If NameToSearch = "" Then
Text1.SetFocus
Exit Sub
End If

Y = 1
Found = False

For RecNum = Y To LastRecord Step 1
Get #FileNum, RecNum, Dictionary

'If Trim(Dictionary.Arabic) Like ("*" & NameToSearch & "*") Then
Dim a
For a = 1 To Len(Dictionary.Arabic) - Len(NameToSearch) - 1
If UCase(Mid(Dictionary.Arabic, a, Len(NameToSearch))) = UCase(NameToSearch) Then
Found = True
GoTo 10:
End If
Next a


20:
Next RecNum

10:
If Found = True Then

    CurrentRecord = RecNum
    Get #FileNum, CurrentRecord, Dictionary

    Text1.Text = Text1.Text & vbNewLine & CurrentRecord & "-" & Trim(Dictionary.Arabic)
    Text2.Text = Text2.Text & vbNewLine & CurrentRecord & "-" & Trim(Dictionary.English)
    Text3.Text = Text3.Text & vbNewLine & CStr(CurrentRecord) & "-"

End If

If RecNum < LastRecord Then
Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)
Y = Y + 1
GoTo 20
Else
MsgBox "Reach Last Record", vbOKOnly, "End of File"
End If
Text1.SetFocus


End Sub

Private Sub cmdTranslate2_Click()
Dim NameToSearch As String
Dim Found As Integer



NameToSearch = Text2.Text
If NameToSearch = "" Then
Text2.SetFocus
Exit Sub
End If
'NameToSearch = UCase(NameToSearch)

Y = 1

Found = False

For RecNum = Y To LastRecord Step 1
Get #FileNum, RecNum, Dictionary
If Trim(Dictionary.English) Like ("*" & NameToSearch & "*") Then
Found = True
'Exit For
GoTo 10:
End If
20:
Next
10:
If Found = True Then
'SaveCurrentRecord
CurrentRecord = RecNum

'ShowCurrentRecord

Get #FileNum, CurrentRecord, Dictionary
'Display Dictionary
  
    If Text2.Text <> "" Then
    Text1.Text = Text1.Text & vbNewLine & Trim(Dictionary.Arabic)
    Text2.Text = Text2.Text & vbNewLine & Trim(Dictionary.English)
    If CurrentRecord < LastRecord Then
    Text3.Text = Text3.Text & vbNewLine & CStr(CurrentRecord) & "-"
    End If
    Else
    Text1.Text = Trim(Dictionary.Arabic)
    Text2.Text = Trim(Dictionary.English)
    Text3.Text = CStr(CurrentRecord) & "-"
    End If
    
'Display the current Record Number in the caption of the form
'Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)
End If
If RecNum < LastRecord Then
Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)
Y = Y + 1
GoTo 20
Else
MsgBox "Reach Last Record", vbOKOnly, "End of File"
'MsgBox "Name " & NameToSearch & " Not Found"
End If
Text2.SetFocus

End Sub

Private Sub Form_Load()
Form1.Icon = LoadPicture(App.Path & "\Misc43.ico")
Y = 1
X = 1
RecordLen = Len(Dictionary)
'Get the Next Available File Number
FileNum = FreeFile(1 - 10)
Dim MyPath
MyPath = App.Path
Dim MyFile
MyFile = MyPath & "\" & "Trans.txt"
'Open(or creat) a file for random access
Open MyFile For Random As FileNum Len = RecordLen
'Update current record
CurrentRecord = 1
'Find What is the last record number of the file
LastRecord = FileLen(MyFile) / RecordLen
'If the file was just created (i.e.LastRecord=0)
'Update LastRecord to 1
If LastRecord = 0 Then
LastRecord = 1
End If
'Display the current record
'ShowCurrentRecord       'this is a name of procedure
Text1.Text = ""
Text2.Text = ""
Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)
End Sub
Public Sub ShowCurrentRecord()
'Fill Dictionary with the data of the current record
Get #FileNum, CurrentRecord, Dictionary
'Display Dictionary
Text1.Text = Trim(Dictionary.Arabic)
Text2.Text = Trim(Dictionary.English)

'Display the current Record Number in the caption of the form
Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)

End Sub


Private Sub cmdNew_Click()
'SaveCurrentRecord    'Procedure name
''add a new blank record
LastRecord = LastRecord + 1
Dictionary.Arabic = ""
Dictionary.English = ""

Put #FileNum, LastRecord, Dictionary
'Update CurrentRecord
CurrentRecord = LastRecord
'Display the record just created
ShowCurrentRecord
Form1.Caption = "Record " & Str(CurrentRecord) & " / " & Str(LastRecord)
Text3.Text = LastRecord
Text1.SetFocus

End Sub

Private Sub Form_Unload(Cancel As Integer)
Close #FileNum
End Sub

Private Sub mnuExit_Click()
End
End Sub



Private Sub mnuHelp_Click()
Dim MyFile2, MyPath
MyPath = App.Path
MyFile2 = MyPath & "\Trans.hlp"
Const HelpFinder = &HB&
CommonDialog1.Action = 6 '6 means run winhlp32.exe
CommonDialog1.HelpFile = MyFile2

CommonDialog1.HelpCommand = HelpFinder
CommonDialog1.ShowHelp

End Sub

Private Sub Timer1_Timer()
Banner.Caption = Mid("The Dictinary that you can add words to it", 1, X)
X = X + 1
If X > Len("The Dictinary that you can add words to it") Then
X = 1
End If
End Sub
