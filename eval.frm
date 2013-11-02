VERSION 5.00
Begin VB.Form EvalForm 
   Caption         =   "Form1"
   ClientHeight    =   6360
   ClientLeft      =   2880
   ClientTop       =   2505
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8220
   Begin VB.Frame Frame1 
      Caption         =   "INSERT EVALUATOR DETAILS"
      Height          =   5055
      Left            =   1080
      TabIndex        =   0
      Top             =   600
      Width           =   6255
      Begin VB.CommandButton cmdInsertEval 
         Caption         =   "&Insert"
         Default         =   -1  'True
         Height          =   495
         Left            =   2040
         TabIndex        =   7
         Top             =   4200
         Width           =   2175
      End
      Begin VB.Frame Frame4 
         Caption         =   "Evaluator's Phone Number"
         Height          =   975
         Left            =   720
         TabIndex        =   3
         Top             =   2880
         Width           =   4815
         Begin VB.TextBox Text3 
            Height          =   375
            Left            =   1320
            MaxLength       =   12
            TabIndex        =   6
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Evaluator Name"
         Height          =   975
         Left            =   720
         TabIndex        =   2
         Top             =   1800
         Width           =   4815
         Begin VB.TextBox Text2 
            Height          =   375
            Left            =   1320
            TabIndex        =   5
            Top             =   360
            Width           =   2775
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Evaluator Code"
         Height          =   975
         Left            =   720
         TabIndex        =   1
         Top             =   720
         Width           =   4815
         Begin VB.TextBox Text1 
            Height          =   375
            Left            =   1320
            TabIndex        =   4
            Top             =   360
            Width           =   2775
         End
      End
   End
End
Attribute VB_Name = "EvalForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdInsertEval_Click()
On Error GoTo Handler
ProjEnv.Insert_Eval Text1.Text, Text2.Text, Text3.Text
ClearControls
Exit Sub
Handler:
 SubjectForm.ErrorMsg "An error occured."
End Sub
Private Sub ClearControls()
 Text1.Text = ""
 Text2.Text = ""
 Text3.Text = ""
End Sub




Private Sub Text1_Change()
 SubjectForm.ToUpper Text1
End Sub

Private Sub Text2_Change()
SubjectForm.ToUpper Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
CheckAlpha KeyAscii
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
  SubjectForm.CheckNum KeyAscii
End Sub
Public Function CheckAlpha(ByRef KeyAscii As Integer)
Select Case KeyAscii
    Case 65 To 90, 97 To 122, 46, 8, 127
      'do nothing
    Case Else
      KeyAscii = 0

  End Select
End Function
