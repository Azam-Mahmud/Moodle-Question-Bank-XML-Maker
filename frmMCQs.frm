VERSION 5.00
Begin VB.Form frmMCQs 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MCQs to Moodle Quiz XML"
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
      Left            =   105
      TabIndex        =   10
      Top             =   5910
      Width           =   1965
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Z"
      Height          =   525
      Left            =   6300
      TabIndex        =   9
      ToolTipText     =   "Zoom"
      Top             =   1920
      Width           =   495
   End
   Begin VB.CommandButton Command1 
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
      ToolTipText     =   "Load File..."
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
      Text            =   "frmMCQs.frx":0000
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
Attribute VB_Name = "frmMCQs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInputFile_Click()
    ' Get input file path using FileDialog
    inputFile = OpenDialog(Me.hWnd, "Text Files (*.txt)|*.txt|All Files (*.*)|*.*", "Select Input Text File", App.Path)
    
    ' Read the input text file
    Open inputFile For Input As #1
    txtInputFile = Trim$(Input$(LOF(1), 1))
    Close #1
End Sub
Private Sub cmdTemplate_Click()
    With frmZoom
        .txtInput = "" & _
            "Question: This is a question text" & vbCrLf & _
            "a.choice1" & vbCrLf & _
            "b.choice2" & vbCrLf & _
            "c.choice3" & vbCrLf & _
            "d.choice4" & vbCrLf & _
            "Answer: d.choice 4" & vbCrLf & _
            "" & vbCrLf & _
            "Question: this is a 2nd question text." & vbCrLf & _
            "a.choice1" & vbCrLf & _
            "b.choice2" & vbCrLf & _
            "Answer: a.choice1"
        .Show 1
    End With
    Unload frmZoom
End Sub
Private Sub Command1_Click()
    txtOutputFile = SaveDialog(Me.hWnd, "XML Files (*.xml)|*.xml|All Files (*.*)|*.*", "Save Output XML File...", App.Path)
End Sub

Private Sub Command2_Click()
    With frmZoom
        .txtInput = txtInputFile
        .Show 1
        If .Ok Then
            txtInputFile = .txtInput
        End If
    End With
    Unload frmZoom
End Sub

'Private Sub ConvertToMoodleXML_Click()
'    Dim inputFile As String
'    Dim outputFile As String
'    Dim inputText As String
'    Dim outputXML As String
'
'
'    inputFile = txtInputFile.Text
'    outputFile = txtOutputFile.Text
'
'
'    ' Read the input text file
'    inputText = Trim$(txtInputFile)
'
'    ' Get tags from txtTags
'    Dim tags() As String
'    tags = Split(Trim(txtTags.Text), vbCrLf)
'
'    ' Split the input text into question and answer pairs
'    Dim questionLines() As String
'    questionLines = Split(inputText, vbCrLf)
'
'    ' Generate Moodle XML output
'    outputXML = ""
'    For i = 0 To UBound(questionLines) Step 2
'        Dim questionText As String
'        Dim shortAnswer As String
'        questionText = questionLines(i)
'        If Trim(questionText) <> "" Then
'            shortAnswer = questionLines(i + 1)
'            outputXML = outputXML & GenerateMoodleXML(questionText, shortAnswer, tags)
'        End If
'    Next i
'
'    ' Wrap the generated XML in the necessary Moodle XML structure
'    outputXML = "<quiz>" & vbCrLf & outputXML & vbCrLf & "</quiz>"
'
'    ' Write the Moodle XML to the output file
'    Open outputFile For Output As #2
'    Print #2, outputXML
'    Close #2
'
'    MsgBox "Conversion complete. Moodle XML file generated."
'End Sub
'
'Private Function GenerateMoodleXML(questionText As String, shortAnswer As String, tags() As String) As String
'    Dim moodleXML As String
'
'    ' Escape special characters in question and short answer
'    questionText = Replace(questionText, "&", "&amp;")
'    questionText = Replace(questionText, "<", "&lt;")
'    questionText = Replace(questionText, ">", "&gt;")
'
'    shortAnswer = Replace(shortAnswer, "&", "&amp;")
'    shortAnswer = Replace(shortAnswer, "<", "&lt;")
'    shortAnswer = Replace(shortAnswer, ">", "&gt;")
'
'    ' Construct Moodle XML for the question
'    moodleXML = vbCrLf & "<question type=""essay"">"
'    moodleXML = moodleXML & vbCrLf & "    <name>"
'    moodleXML = moodleXML & vbCrLf & "        <text>" & Replace(questionText, "Question: ", "") & "</text>"
'    moodleXML = moodleXML & vbCrLf & "    </name>"
'    moodleXML = moodleXML & vbCrLf & "    <questiontext format=""html"">"
'    moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[" & Replace(questionText, "Question: ", "") & "]]></text>"
'    moodleXML = moodleXML & vbCrLf & "    </questiontext>"
'    moodleXML = moodleXML & vbCrLf & "    <generalfeedback format=""html"">"
'    moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[<p dir=""ltr"" style=""text-align: left;"">General Feedback<br></p>]]></text>"
'    moodleXML = moodleXML & vbCrLf & "    </generalfeedback>"
'    moodleXML = moodleXML & vbCrLf & "    <defaultgrade>1.0000000</defaultgrade>"
'    moodleXML = moodleXML & vbCrLf & "    <penalty>0.0000000</penalty>"
'    moodleXML = moodleXML & vbCrLf & "    <hidden>0</hidden>"
'    moodleXML = moodleXML & vbCrLf & "    <idnumber></idnumber>"
'    moodleXML = moodleXML & vbCrLf & "    <responseformat>plain</responseformat>"
'    moodleXML = moodleXML & vbCrLf & "    <responserequired>1</responserequired>"
'    moodleXML = moodleXML & vbCrLf & "    <responsefieldlines>10</responsefieldlines>"
'    moodleXML = moodleXML & vbCrLf & "    <minwordlimit></minwordlimit>"
'    moodleXML = moodleXML & vbCrLf & "    <maxwordlimit></maxwordlimit>"
'    moodleXML = moodleXML & vbCrLf & "    <attachments>0</attachments>"
'    moodleXML = moodleXML & vbCrLf & "    <attachmentsrequired>0</attachmentsrequired>"
'    moodleXML = moodleXML & vbCrLf & "    <maxbytes>0</maxbytes>"
'    moodleXML = moodleXML & vbCrLf & "    <filetypeslist></filetypeslist>"
'    moodleXML = moodleXML & vbCrLf & "    <graderinfo format=""html"">"
'    moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[<p dir=""ltr"" style=""text-align: left;"">" & Replace(shortAnswer, "Answer: ", "") & "<br></p>]]></text>"
'    moodleXML = moodleXML & vbCrLf & "    </graderinfo>"
'    moodleXML = moodleXML & vbCrLf & "    <responsetemplate format=""html"">"
'    moodleXML = moodleXML & vbCrLf & "        <text></text>"
'    moodleXML = moodleXML & vbCrLf & "    </responsetemplate>"
'
'    ' Add tags
'    Dim Tag
'    If UBound(tags) >= 0 Then
'        moodleXML = moodleXML & vbCrLf & "    <tags>"
'        For Each Tag In tags
'            If Trim(Tag) <> "" Then moodleXML = moodleXML & vbCrLf & "        <tag><text>" & Tag & "</text></tag>"
'        Next Tag
'        moodleXML = moodleXML & vbCrLf & "    </tags>"
'    End If
'
'    moodleXML = moodleXML & vbCrLf & "</question>"
'
'    GenerateMoodleXML = moodleXML
'End Function


Private Sub ConvertToMoodleXML_Click()
    Dim inputFile As String
    Dim outputFile As String
    Dim inputText As String
    Dim outputXML As String
    
    inputFile = txtInputFile.Text
    outputFile = txtOutputFile.Text
    
    ' Read the input
    inputText = txtInputFile
    
    ' Split the input text into question and answer blocks
    Dim questionBlocks() As String
    questionBlocks = Split(inputText, vbCrLf & vbCrLf)
    
    ' Generate Moodle XML output
    outputXML = ""
    For i = 0 To UBound(questionBlocks)
        Dim questionBlock As String
        questionBlock = Trim$(questionBlocks(i))
        If Len(questionBlock) > 0 Then
            outputXML = outputXML & GenerateMoodleXML(questionBlock)
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

Private Function GenerateMoodleXML(questionBlock As String) As String
    Dim moodleXML As String
    
    ' Split the question block into lines
    Dim lines() As String
    
    lines = Split(Trim$(questionBlock), vbCrLf)
    
    ' Process question and choices
    Dim questionText As String
    Dim choices() As String
    Dim correctChoiceIndex As Integer
    Dim currentChoice As Integer
    
    Dim answerLine() As String
    
    questionText = Replace(lines(0), "Question: ", "")
    
    For i = 1 To UBound(lines)
        t = lines(i)
        If InStr(1, t, "Answer: ") Then
            answerLine() = Split(t, "Answer: ")
            answerNumber = Left(answerLine(1), 2)
        End If
    Next i
    
    If answerNumber = "" Then Exit Function
    
    
    
    
    ' Construct Moodle XML for the question
    
    moodleXML = moodleXML & vbCrLf & "<question type=""multichoice"">"
    moodleXML = moodleXML & vbCrLf & "    <name>"
    moodleXML = moodleXML & vbCrLf & "      <text><![CDATA[" & questionText & "]]></text>"
    moodleXML = moodleXML & vbCrLf & "    </name>"
    moodleXML = moodleXML & vbCrLf & "    <questiontext format=""html"">"
    moodleXML = moodleXML & vbCrLf & "      <text><![CDATA[" & questionText & "]]></text>"
    moodleXML = moodleXML & vbCrLf & "    </questiontext>"
    moodleXML = moodleXML & vbCrLf & "    <generalfeedback format=""html"">"
    moodleXML = moodleXML & vbCrLf & "      <text></text>"
    moodleXML = moodleXML & vbCrLf & "    </generalfeedback>"
    moodleXML = moodleXML & vbCrLf & "    <defaultgrade>1.0000000</defaultgrade>"
    moodleXML = moodleXML & vbCrLf & "    <penalty>0.3333333</penalty>"
    moodleXML = moodleXML & vbCrLf & "    <hidden>0</hidden>"
    moodleXML = moodleXML & vbCrLf & "    <idnumber></idnumber>"
    moodleXML = moodleXML & vbCrLf & "    <single>true</single>"
    moodleXML = moodleXML & vbCrLf & "    <shuffleanswers>true</shuffleanswers>"
    moodleXML = moodleXML & vbCrLf & "    <answernumbering>abc</answernumbering>"
    moodleXML = moodleXML & vbCrLf & "    <showstandardinstruction>1</showstandardinstruction>"
    moodleXML = moodleXML & vbCrLf & "    <correctfeedback format=""html"">"
    moodleXML = moodleXML & vbCrLf & "      <text>Well Done.</text>"
    moodleXML = moodleXML & vbCrLf & "    </correctfeedback>"
    moodleXML = moodleXML & vbCrLf & "    <partiallycorrectfeedback format=""html"">"
    moodleXML = moodleXML & vbCrLf & "      <text>Good but you can do better.</text>"
    moodleXML = moodleXML & vbCrLf & "    </partiallycorrectfeedback>"
    moodleXML = moodleXML & vbCrLf & "    <incorrectfeedback format=""html"">"
    moodleXML = moodleXML & vbCrLf & "      <text>Better luck next time</text>"
    moodleXML = moodleXML & vbCrLf & "    </incorrectfeedback>"
    
    For i = 1 To UBound(lines) - 1
'        MsgBox lines(i)
        If InStr(1, lines(i), "Answer:") Then
        
        Else
            choiceText = Right(lines(i), Len(lines(i)) - 2)
            
            If answerNumber = Left(lines(i), 2) Then
                moodleXML = moodleXML & vbCrLf & "    <answer fraction=""100"" format=""html"">"
                feedBack = "Well Done"
            Else
                moodleXML = moodleXML & vbCrLf & "    <answer fraction=""0"" format=""html"">"
                feedBack = ""
            End If
            moodleXML = moodleXML & vbCrLf & "        <text><![CDATA[" & choiceText & "]]></text>"
            moodleXML = moodleXML & vbCrLf & "        <feedback format=""html"">"
            moodleXML = moodleXML & vbCrLf & "            <text>" & feedBack & "</text>"
            moodleXML = moodleXML & vbCrLf & "        </feedback>"
            moodleXML = moodleXML & vbCrLf & "    </answer>"
        End If
    Next i
    
     
    moodleXML = moodleXML & vbCrLf & "</question>"
    
    GenerateMoodleXML = moodleXML
End Function

