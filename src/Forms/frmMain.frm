VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Convert to Moodle XML"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   13.5
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   4155
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "MCQs"
      Height          =   1005
      Left            =   210
      TabIndex        =   2
      Top             =   150
      Width           =   3765
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fill in the Blanks"
      Height          =   1005
      Left            =   210
      TabIndex        =   1
      Top             =   1350
      Width           =   3765
   End
   Begin VB.CommandButton cmdEssayShortAnswer 
      Caption         =   "Short Answer (Essays)"
      Height          =   1005
      Left            =   210
      TabIndex        =   0
      Top             =   2520
      Width           =   3765
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdEssayShortAnswer_Click()
    frmShortEssayAnswer.Show 1
End Sub

Private Sub Command1_Click()
    frmFillBlanks.Show 1
End Sub

Private Sub Command2_Click()
    frmMCQs.Show 1
End Sub
