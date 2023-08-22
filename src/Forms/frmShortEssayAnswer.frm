VERSION 5.00
Begin VB.Form frmShortEssayAnswer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Short Answers to Moodle Quiz XML"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6960
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTemplate 
      Caption         =   "Template"
      Height          =   675
      Left            =   90
      TabIndex        =   10
      Top             =   5910
      Width           =   1965
   End
   Begin VB.CommandButton cmdZoom 
      Caption         =   "Z"
      Height          =   525
      Left            =   6300
      TabIndex        =   9
      ToolTipText     =   "Zoom (Full Screen Editor)"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton cmdOutputFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6300
      TabIndex        =   2
      ToolTipText     =   "Output file path"
      Top             =   2820
      Width           =   495
   End
   Begin VB.CommandButton cmdInputFile 
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6300
      TabIndex        =   1
      Top             =   390
      Width           =   495
   End
   Begin VB.TextBox txtOutputFile 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2820
      Width           =   6165
   End
   Begin VB.TextBox txtInputFile 
      Height          =   2055
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   390
      Width           =   6165
   End
   Begin VB.TextBox txtTags 
      Height          =   1905
      Left            =   105
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmShortEssayAnswer.frx":0000
      Top             =   3900
      Width           =   6765
   End
   Begin VB.CommandButton ConvertToMoodleXML 
      Caption         =   "Convert"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   4920
      TabIndex        =   4
      Top             =   5910
      Width           =   1965
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Output file in .xml format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   8
      Top             =   2490
      Width           =   2955
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paste directly or load Input File in .txt format"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   195
      TabIndex        =   0
      Top             =   60
      Width           =   5340
   End
   Begin VB.Label Label 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Questions' Tags"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   105
      TabIndex        =   5
      Top             =   3600
      Width           =   1950
   End
End
Attribute VB_Name = "frmShortEssayAnswer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInputFile_Click()
    ' Get input file path using FileDialog
    inputFile = OpenDialog(Me.hWnd, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*", "Select Input Text File", App.Path)
    If inputFile = "" Then Exit Sub
    
    ' Read the input text file
    Open inputFile For Input As #1
    txtInputFile = Trim$(Input$(LOF(1), 1))
    Close #1
End Sub
Private Sub cmdOutputFile_Click()
    txtOutputFile = SaveDialog(Me.hWnd, "XML Files (*.xml)|*.xml|All Files (*.*)|*.*", "Save Output XML File...", App.Path)
End Sub

Private Sub cmdZoom_Click()
    With frmZoom
        .txtInput = txtInputFile
        .Show 1
        If .Ok Then
            txtInputFile = .txtInput
        End If
    End With
    Unload frmZoom
End Sub

Private Sub cmdTemplate_Click()
    With frmZoom
        .txtInput = "" & _
            "Question: <Question Text Here in single line>" & vbCrLf & _
            "Answer: <Answer text here in single line without line breaks (new lines)." & vbCrLf & _
            "Question: <Question Text Here in single line>" & vbCrLf & _
            "Answer: <Answer text here in single line without line breaks (new lines)."
        .Show 1
    End With
    Unload frmZoom
End Sub

Private Sub ConvertToMoodleXML_Click()
    Dim inputFile As String
    Dim outputFile As String
    Dim inputText As String
    Dim outputXML As String
    
    
    inputFile = txtInputFile.Text
    outputFile = txtOutputFile.Text
    
    
    ' Read the input text file
    inputText = Trim$(txtInputFile)
    
    ' Get tags from txtTags
    Dim tags() As String
    tags = Split(Trim(txtTags.Text), vbCrLf)
    
    ' Split the input text into question and answer pairs
    Dim questionLines() As String
    questionLines = Split(inputText, vbCrLf)
    
    ' Generate Moodle XML output
    outputXML = ""
    For i = 0 To UBound(questionLines) Step 2
        Dim questionText As String
        Dim shortAnswer As String
        questionText = questionLines(i)
        If Trim(questionText) <> "" Then
            shortAnswer = questionLines(i + 1)
            outputXML = outputXML & GenerateMoodleXML(questionText, shortAnswer, tags)
        End If
    Next i
    
    ' Wrap the generated XML in the necessary Moodle XML structure
    outputXML = "<quiz>" & vbCrLf & outputXML & vbCrLf & "</quiz>"
    
    ' Write the Moodle XML to the output file
    Open outputFile For Output As #2
    Print #2, outputXML
    Close #2
    
    MsgBox "Conversion complete. Moodle XML file generated."
End Sub

Private Function GenerateMoodleXML(questionText As String, shortAnswer As String, tags() As String) As String
    Dim moodleXML As String
    
    ' Escape special characters in question and short answer
    questionText = Replace(questionText, "&", "&amp;")
    questionText = Replace(questionText, "<", "&lt;")
    questionText = Replace(questionText, ">", "&gt;")
    
    shortAnswer = Replace(shortAnswer, "&", "&amp;")
    shortAnswer = Replace(shortAnswer, "<", "&lt;")
    shortAnswer = Replace(shortAnswer, ">", "&gt;")
    
    ' Construct Moodle XML for the question
    moodleXML = vbCrLf & "<question type=""essay"">"
    moodleXML = moodleXML & vbCrLf & "    <name>"
    moodleXML = moodleXML & vbCrLf & "        <text>" & Replace(questionText, "Question: ", "") & "</text>"
    moodleXML = moodleXML & vbCrLf & "    </name>"
    moodleXML = moodleXML & vbCrLf & "    <questiontext format=""html"">"
    moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[" & Replace(questionText, "Question: ", "") & "]]></text>"
    moodleXML = moodleXML & vbCrLf & "    </questiontext>"
    moodleXML = moodleXML & vbCrLf & "    <generalfeedback format=""html"">"
    moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[<p dir=""ltr"" style=""text-align: left;"">General Feedback<br></p>]]></text>"
    moodleXML = moodleXML & vbCrLf & "    </generalfeedback>"
    moodleXML = moodleXML & vbCrLf & "    <defaultgrade>1.0000000</defaultgrade>"
    moodleXML = moodleXML & vbCrLf & "    <penalty>0.0000000</penalty>"
    moodleXML = moodleXML & vbCrLf & "    <hidden>0</hidden>"
    moodleXML = moodleXML & vbCrLf & "    <idnumber></idnumber>"
    moodleXML = moodleXML & vbCrLf & "    <responseformat>plain</responseformat>"
    moodleXML = moodleXML & vbCrLf & "    <responserequired>1</responserequired>"
    moodleXML = moodleXML & vbCrLf & "    <responsefieldlines>10</responsefieldlines>"
    moodleXML = moodleXML & vbCrLf & "    <minwordlimit></minwordlimit>"
    moodleXML = moodleXML & vbCrLf & "    <maxwordlimit></maxwordlimit>"
    moodleXML = moodleXML & vbCrLf & "    <attachments>0</attachments>"
    moodleXML = moodleXML & vbCrLf & "    <attachmentsrequired>0</attachmentsrequired>"
    moodleXML = moodleXML & vbCrLf & "    <maxbytes>0</maxbytes>"
    moodleXML = moodleXML & vbCrLf & "    <filetypeslist></filetypeslist>"
    moodleXML = moodleXML & vbCrLf & "    <graderinfo format=""html"">"
    moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[<p dir=""ltr"" style=""text-align: left;"">" & Replace(shortAnswer, "Answer: ", "") & "<br></p>]]></text>"
    moodleXML = moodleXML & vbCrLf & "    </graderinfo>"
    moodleXML = moodleXML & vbCrLf & "    <responsetemplate format=""html"">"
    moodleXML = moodleXML & vbCrLf & "        <text></text>"
    moodleXML = moodleXML & vbCrLf & "    </responsetemplate>"
    
    ' Add tags
    Dim Tag
    If UBound(tags) >= 0 Then
        moodleXML = moodleXML & vbCrLf & "    <tags>"
        For Each Tag In tags
            If Trim(Tag) <> "" Then moodleXML = moodleXML & vbCrLf & "        <tag><text>" & Tag & "</text></tag>"
        Next Tag
        moodleXML = moodleXML & vbCrLf & "    </tags>"
    End If
    
    moodleXML = moodleXML & vbCrLf & "</question>"
    
    GenerateMoodleXML = moodleXML
End Function

