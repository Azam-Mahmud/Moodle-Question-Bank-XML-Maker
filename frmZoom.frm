VERSION 5.00
Begin VB.Form frmZoom 
   Caption         =   "Input Questions"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13725
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   13725
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   645
      Left            =   11400
      TabIndex        =   2
      Top             =   5820
      Width           =   2115
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Ok"
      Height          =   645
      Left            =   360
      TabIndex        =   1
      Top             =   5820
      Width           =   2115
   End
   Begin VB.TextBox txtInput 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5655
      Left            =   30
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   30
      Width           =   13665
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Ok As Boolean

Private Sub cmdCancel_Click()
    Ok = False
    Me.Hide
End Sub

Private Sub cmdOk_Click()
    Ok = True
    Me.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Select Case UnloadMode
        Case vbFormControlMenu
            Cancel = True
            Ok = False
            Me.Hide
'        Case vbFormCode
'            MsgBox Caption & ": Unload statement from code."
'        Case vbAppWindows
'            MsgBox Caption & ": Windows session ending."
'        Case vbAppTaskManager
'            MsgBox Caption & ": Task Manager close."
'        Case vbFormMDIForm
'            MsgBox Caption & ": MDI parent is closing."
'        Case vbFormOwner
'            MsgBox Caption & ": Owner is closing."
    End Select
End Sub
Private Sub Form_Resize()
    txtInput.Move 100, 100, Me.ScaleWidth - 200, Me.ScaleHeight - 400 - cmdOk.Height
    cmdOk.Move 400, txtInput.Height + 250
    cmdCancel.Move Me.ScaleWidth - cmdCancel.Width - 400, cmdOk.Top
End Sub
