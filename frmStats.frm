VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmStats 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visual Basic 6.0 Code Statistics"
   ClientHeight    =   3735
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   Icon            =   "frmStats.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3735
   ScaleWidth      =   5700
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdScan 
      Caption         =   "&Scan"
      Height          =   375
      Left            =   4800
      TabIndex        =   20
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame framProj 
      Caption         =   "Statistics"
      Height          =   3255
      Left            =   0
      TabIndex        =   3
      Top             =   480
      Width           =   4695
      Begin VB.Label lblApi 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3720
         TabIndex        =   34
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDApi 
         BackStyle       =   0  'Transparent
         Caption         =   "API Declarations :"
         Height          =   255
         Left            =   2400
         TabIndex        =   33
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblControl 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   32
         Top             =   2160
         Width           =   855
      End
      Begin VB.Label lblDControl 
         BackStyle       =   0  'Transparent
         Caption         =   "Controls :"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label lblProp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3240
         TabIndex        =   30
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblDProp 
         BackStyle       =   0  'Transparent
         Caption         =   "Properties :"
         Height          =   255
         Left            =   2400
         TabIndex        =   29
         Top             =   720
         Width           =   855
      End
      Begin VB.Label lblClass 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1320
         TabIndex        =   28
         Top             =   1800
         Width           =   615
      End
      Begin VB.Label lblDClass 
         BackStyle       =   0  'Transparent
         Caption         =   "Class Modules :"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   1800
         Width           =   1215
      End
      Begin VB.Label lblVar 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3840
         TabIndex        =   26
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblDVariables 
         BackStyle       =   0  'Transparent
         Caption         =   "Variables Declared :"
         Height          =   255
         Left            =   2400
         TabIndex        =   25
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label lblCode 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   24
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblDCode 
         BackStyle       =   0  'Transparent
         Caption         =   "Code Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   23
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   22
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblDTotal 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   21
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label lblMod 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   840
         TabIndex        =   19
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label lblDMod 
         BackStyle       =   0  'Transparent
         Caption         =   "Modules :"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label lblForm 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   720
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblDForm 
         BackStyle       =   0  'Transparent
         Caption         =   "Forms :"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label lblFunc 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   2880
         Width           =   735
      End
      Begin VB.Label lblDFunc 
         BackStyle       =   0  'Transparent
         Caption         =   "Functions :"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   2880
         Width           =   855
      End
      Begin VB.Label lblProc 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   1080
         TabIndex        =   13
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label lblDProc 
         BackStyle       =   0  'Transparent
         Caption         =   "Procedures :"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label lblComm 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3600
         TabIndex        =   11
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label lblDComm 
         BackStyle       =   0  'Transparent
         Caption         =   "Comment Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   10
         Top             =   2160
         Width           =   1215
      End
      Begin VB.Label lblBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   255
         Left            =   3360
         TabIndex        =   9
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblDBlank 
         BackStyle       =   0  'Transparent
         Caption         =   "Blank Lines :"
         Height          =   255
         Left            =   2400
         TabIndex        =   8
         Top             =   1800
         Width           =   975
      End
      Begin VB.Label lblVersion 
         BackStyle       =   0  'Transparent
         Caption         =   "v1.0.0"
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDVer 
         BackStyle       =   0  'Transparent
         Caption         =   "Version :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lblName 
         BackStyle       =   0  'Transparent
         Caption         =   "Project1"
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label lblDName 
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name :"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Text            =   "C:\"
      Top             =   120
      Width           =   3855
   End
   Begin MSComDlg.CommonDialog cdgFiles 
      Left            =   4800
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label lblFile 
      BackStyle       =   0  'Transparent
      Caption         =   "Filename"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "frmStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const FormStartCode = "Attribute VB_Exposed "
Const ModStartCode = "Attribute VB_Name "
Const ClsStartCode = "Attribute VB_Exposed"
Const CtlStartCode = "Attribute VB_Exposed"
Const VbpTitle = "Title"
Const VbpMajor = "MajorVer"
Const VbpMinor = "MinorVer"
Const VbpRevision = "RevisionVer"
Const VbpForm = "Form"
Const VbpMod = "Module"
Const VbpClass = "Class" 'This is actually made up of "Class="<object name>"; "<class filename>"
Const VbpControl = "UserControl"
Const BrowseFilter = "VB Project *.Vbp|*.Vbp|VB Modules *.Bas|*.Bas|VB Forms *.Frm|*.Frm|VB Class Modules (*.Cls)|*.Cls|VB User Controls (*.Ctl)|*.Ctl|All Files *.*|*.*"

'code and project counters
Dim NumBlank As Long
Dim NumProc As Long
Dim NumFunc As Long
Dim NumComments As Long
Dim NumForms As Long
Dim NumModules As Long
Dim NumClasses As Long
Dim NumControls As Long
Dim NumProperties As Long
Dim NumCode As Long
Dim NumVariables As Long
Dim NumVarLines As Long
Dim NumAPI As Long
Dim Version As String
Dim ProjectName As String

Public Sub ResetValues()
'reset values and variables

lblName.Caption = ""
Version = "v"
NumBlank = 0
NumProc = 0
NumFunc = 0
NumComments = 0
NumForms = 0
NumModules = 0
NumCode = 0
NumVariables = 0
NumVarLines = 0
NumClasses = 0
NumProperties = 0
NumAPI = 0
NumControls = 0
End Sub

Public Sub DisplayValues()
'This will enter all the appropiate details into the lables and
'total the number of lines found

'half the number of properties countes because there are two property
'statements per property, Let and Get.
'NumProperties = NumProperties / 2

'display results
If Trim(lblName.Caption) = "" Then
    'if the project name is blank then use the default name
    lblName.Caption = "Project1"
End If
If LCase(Version) = "v" Then
    'if version if blank, then set it to default
    Version = "v1.0.0"
End If
lblVersion.Caption = Version
lblBlank.Caption = Format(NumBlank, "0")
lblComm.Caption = Format(NumComments, "0")
lblForm.Caption = Format(NumForms, "0")
lblMod.Caption = Format(NumModules, "0")
lblClass.Caption = Format(NumClasses, "0")
lblProc.Caption = Format(NumProc, "0")
lblFunc.Caption = Format(NumFunc, "0")
lblProp.Caption = Format(NumProperties / 2, "0")
lblCode.Caption = Format(NumCode, "0")
lblVar.Caption = Format(NumVariables, "0")
lblControl.Caption = Format(NumControls, "0")
lblApi.Caption = Format(NumAPI, "0")

'total results accounting for headers/footers of procedures/functions
lblTotal.Caption = Format((NumBlank + NumComments + ((NumProc + NumFunc + NumProperties) * 2) + NumAPI + NumVarLines + NumCode), "0")
End Sub

Public Sub ReadProject(Path As String)
'This will read an entire project and set the values for statistics

Dim FileNum As Integer 'used for the .vbp file
Dim Line As String
Dim FileName As String
Dim StartScan As Boolean

'if path is invalid, then quit
If Dir(Path) = "" Then
    Exit Sub
End If

Call ResetValues
StartScan = False

'open project
FileNum = FreeFile
Open Path For Input As #FileNum
    While Not EOF(FileNum)
        Line Input #FileNum, Line
        
        Select Case GetBefore(Line)
        Case VbpTitle
            lblName.Caption = GetAfter(Line)
        
        Case VbpMajor, VbpMinor
            Version = Version & GetAfter(Line) & "."
        
        Case VbpRevision
            Version = Version & GetAfter(Line)
        
        Case VbpForm
            'scan form
            NumForms = NumForms + 1
            Call ScanFile(AddFile(GetPath(Path), GetAfter(Line)), FormStartCode)
            
        Case VbpMod
            'scan module
            NumModules = NumModules + 1
            Call ScanFile(AddFile(GetPath(Path), GetMod(Line)), ModStartCode)
        
        Case VbpClass
            'scan class module
            NumClasses = NumClasses + 1
            Call ScanFile(AddFile(GetPath(Path), GetClass(Line)), ClsStartCode)
            
        Case VbpControl
            'scan user control
            NumControls = NumControls + 1
            Call ScanFile(AddFile(GetPath(Path), GetAfter(Line)), CtlStartCode)
        
        End Select
        
    Wend
Close #FileNum

Call DisplayValues
End Sub

Public Sub IncrementVal(Line As String)
'This will increment the appropiate values based on the text

Const EndProc = "End Sub"
Const EndFunc = "End Function"
Const EndProp = "End Property"
Const DecApi = "Declare "
Const LibApi = " Lib "
Const VarA = "Public"
Const VarB = "Private"
Const VarC = "Global"
Const VarD = "Dim"
Const VarE = "Static"
Const VarAs = " As "
Const Comment = "'"
Const Blank = ""

'Comments
If Left(Line, 1) = Comment Then
    NumComments = NumComments + 1
    Exit Sub
End If

'Blanks
If Line = Blank Then
    NumBlank = NumBlank + 1
    Exit Sub
End If

'the footers of the functions and procedures. I'm counting the footers
'because they are always the same no matter what keywords the
'title has.
If Left(Line, Len(EndProc)) = EndProc Then
    NumProc = NumProc + 1
    
    'code num as already counted the header, so subtract this.
    NumCode = NumCode - 1
    Exit Sub
End If
If Left(Line, Len(EndFunc)) = EndFunc Then
    NumFunc = NumFunc + 1
    
    'code num as already counted the header, so subtract this.
    NumCode = NumCode - 1
    Exit Sub
End If
If Left(Line, Len(EndFunc)) = EndProp Then
    NumProperties = NumProperties + 1
    
    'code num as already counted the header, so subtract this.
    NumCode = NumCode - 1
    Exit Sub
End If

'check for api declarations
If (InStr(1, Line, DecApi) <> 0) And (InStr(1, Line, LibApi) <> 0) Then
    NumAPI = NumAPI + 1
    Exit Sub
End If

'variable declarations
'if the left part of the string contains one of the variable decalration
'keywords and also contains the keyword " As " and does not contain
'the api declaration keyword "Declare", then the string is a variable
'declaration.
'NOTE: These variables do NOT count procedure/function parameters.
'Also, the number of variables is not the same as the number of
'lines used to declare them eg,
'Dim MyVar1, MyVar2, MyVar3 As Integer
If ((Left(Line, Len(VarA)) = VarA) Or (Left(Line, Len(VarB)) = VarB) Or (Left(Line, Len(VarC)) = VarC) Or (Left(Line, Len(VarD)) = VarD) Or (Left(Line, Len(VarE)) = VarE)) And (InStr(1, Line, VarAs) <> 0) And (InStr(1, Line, DecApi) = 0) Then
    NumVariables = NumVariables + 1 + CommaCount(Line)
    NumVarLines = NumVarLines + 1
    Exit Sub
End If

'else the line is code
NumCode = NumCode + 1
End Sub

Public Function GetPath(Address As String) As String
'This function returns the path from a string containing the full
'path and filename of a file.

Dim Counter As Integer
Dim LastPos As Integer

'find the position of the last "\" mark in the string
LastPos = 1
For Counter = 1 To Len(Address)
    If Mid(Address, Counter, 1) = "\" Then
        LastPos = Counter
    End If
Next Counter

'return everything before the last "\" mark
GetPath = Left(Address, (LastPos - 1))
End Function

Public Function GetBefore(Sentence As String) As String
'This procedure returns all the character of a
'string before the "=" sign.

Const Sign = "="

Dim Counter As Integer
Dim Before As String

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If (Counter <> Len(Sentence)) And (Counter <> 0) Then
    Before = Left(Sentence, (Counter - 1))
Else
    Before = ""
End If

GetBefore = Before
End Function

Public Function GetAfter(Sentence As String, Optional ByVal Character As String = "=") As String
'This procedure returns all the character of a
'string after the "=" sign.

'Const Sign = "="

Dim Counter As Integer
Dim Rest As String
Dim Sign As String

Sign = Character

'find the position of the equals sign
Counter = InStr(1, Sentence, Sign)

If Counter <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - Counter))
Else
    Rest = ""
End If

GetAfter = Rest
End Function

Public Function GetMod(Sentence As String) As String
'This procedure returns all the character of a
'string after the ";" sign.

Const ModName = ";"

Dim Rest As String
Dim ModPos As Integer

'find the position of the ; sign
ModPos = InStr(1, Sentence, ModName) + 1

If ModPos <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - ModPos))
Else
    Rest = ""
End If

GetMod = Rest
End Function

Public Function GetClass(Sentence As String) As String
'This procedure returns all the character of a
'string after the "; " sign.

Const ClassName = "; "

Dim Rest As String
Dim ClassPos As Integer

'find the position of the ; sign
ClassPos = InStr(1, Sentence, ClassName) + 1

If ClassPos <> Len(Sentence) Then
    Rest = Right(Sentence, (Len(Sentence) - ClassPos))
Else
    Rest = ""
End If

GetClass = Rest
End Function

Private Sub cmdBrowse_Click()
cdgFiles.Filter = BrowseFilter
cdgFiles.InitDir = GetPath(txtPath.Text)
cdgFiles.ShowOpen
txtPath.Text = cdgFiles.FileName
End Sub

Private Sub cmdScan_Click()
'Try to scan the file specified in the text box

Const ProjExt = "vbp"
Const FormExt = "frm"
Const ModuleExt = "bas"
Const ClassExt = "cls"
Const ControlExt = "ctl"

Dim Extention As String
Dim FilePath As String

FilePath = txtPath.Text
Extention = GetAfter(FilePath, ".")

'don't try to scan file if it doesn't exist
If (Dir(FilePath) = "") Or (FilePath = "") Then
    Exit Sub
End If

'scan each file type differently
Select Case LCase(Extention)
Case LCase(ProjExt)
    'scan an entire project
    Call ReadProject(FilePath)

Case LCase(FormExt)
    'scan one form
    Call ResetValues
    NumForms = NumForms + 1
    Call ScanFile(FilePath, FormStartCode)
    Call DisplayValues

Case LCase(ModuleExt)
    'scan one module
    Call ResetValues
    NumModules = NumModules + 1
    Call ScanFile(FilePath, ModStartCode)
    Call DisplayValues

Case LCase(ClassExt)
    'scan one class
    Call ResetValues
    NumClasses = NumClasses + 1
    Call ScanFile(FilePath, ClsStartCode)
    Call DisplayValues
    
Case LCase(ControlExt)
    'scan one control
    Call ResetValues
    NumControls = NumControls + 1
    Call ScanFile(FilePath, CtlStartCode)
    Call DisplayValues
    
End Select
End Sub

Private Sub ScanFile(Path As String, Start As String)
'This procedure will scan a file starting at the first point with the
'specified starting string.

Dim FileNum As Integer
Dim Line As String
Dim StartScan As Boolean

FileNum = FreeFile

If Dir(Path) = "" Then
    'invalid path
    Exit Sub
End If

Open Path For Input As #FileNum
    'scan file
    While Not EOF(FileNum)
        Line Input #FileNum, Line
        If StartScan Then
            Call IncrementVal(LTrim(Line))
        End If
        
        If Left(Line, Len(Start)) = Start Then
            'scan code
            StartScan = True
        End If
    Wend
Close #FileNum
End Sub

Private Sub Form_Load()
txtPath.Text = App.Path
txtPath.SelLength = Len(txtPath.Text)
End Sub

Public Function CommaCount(ByVal Line As String) As Integer
'This will return the number of commas foun in the string. Mainly
'use to find the number of variables declared on the same line

Dim Counter As Integer
Dim LastPos As Integer
Dim CommaNum As Integer

LastPos = 0

Do
    Counter = InStr(LastPos + 1, Line, ",")
    
    If Counter <> 0 Then
        CommaNum = CommaNum + 1
    End If
    LastPos = Counter
Loop Until LastPos = 0

'return result
CommaCount = CommaNum
End Function

Public Function AddFile(ByVal Directory As String, ByVal FileName As String) As String
'This will add a file name to a directory path to create a full filepath.

If Right(Directory, 1) <> "\" Then
    'insert a backslash
    Directory = Directory & "\"
End If

'append the file name to the directory path now
AddFile = Directory & FileName
End Function

