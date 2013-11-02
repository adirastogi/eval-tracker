VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   ScaleHeight     =   4275
   ScaleWidth      =   6840
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame EvalCopyDetails 
      Caption         =   "Frame1"
      Height          =   2895
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   5535
      Begin VB.CommandButton Command1 
         Caption         =   "Print Report"
         Default         =   -1  'True
         Height          =   495
         Left            =   1560
         TabIndex        =   3
         Top             =   1680
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   2400
         TabIndex        =   1
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Enter Evaluator Code :"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   720
         Width           =   2055
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  If Text1.Text = "" Then
  ErrorMsg "Please enter an evaaluator code"
  Else
   ProjEnv.Eval_Copydetails Text1.Text
   EvalDetailsReport
   Text1.Text = ""
   ProjEnv.rs
  End If
  
End Sub

Private Sub Text1_Change()
 SubjectForm.ToUpper Text1
End Sub

