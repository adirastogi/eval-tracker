VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form IssueForm 
   Caption         =   "Form2"
   ClientHeight    =   7665
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9075
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form2"
   ScaleHeight     =   7665
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "ISSUE BUNDLE FOR CHECKING"
      Height          =   6615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      Begin VB.CommandButton cmdIssue 
         Caption         =   "&Issue Bundle"
         Height          =   615
         Left            =   3120
         TabIndex        =   22
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Evaluator Details"
         Height          =   3255
         Left            =   4320
         TabIndex        =   7
         Top             =   2280
         Width           =   3615
         Begin VB.TextBox Text7 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   2040
            Width           =   2895
         End
         Begin VB.TextBox Text6 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   960
            Width           =   2895
         End
         Begin VB.Label Label6 
            Caption         =   "Phone Number"
            Height          =   375
            Left            =   360
            TabIndex        =   21
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Evaluator Name"
            Height          =   375
            Left            =   360
            TabIndex        =   20
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bundle Details"
         Height          =   3255
         Left            =   240
         TabIndex        =   2
         Top             =   2280
         Width           =   3615
         Begin VB.TextBox Text5 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            TabIndex        =   12
            Top             =   2640
            Width           =   1455
         End
         Begin VB.TextBox Text4 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   1560
            TabIndex        =   11
            Top             =   2040
            Width           =   1455
         End
         Begin VB.TextBox Text3 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   10
            Top             =   1560
            Width           =   1455
         End
         Begin VB.TextBox Text2 
            DataField       =   "Start_Serial"
            DataMember      =   "Bundle_Details"
            DataSource      =   "ProjEnv"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   960
            Width           =   1455
         End
         Begin VB.TextBox Text1 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1560
            TabIndex        =   8
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "Number of Copies Checked"
            Height          =   375
            Index           =   3
            Left            =   240
            TabIndex        =   17
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Total Number of Copies"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   16
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "End Serial Number"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   15
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Start Serial Number"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   14
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Subject Code"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Bundle"
         Height          =   1215
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   7695
         Begin MSDataListLib.DataCombo dcEvalCode 
            Bindings        =   "issue.frx":0000
            Height          =   315
            Left            =   2520
            TabIndex        =   4
            Top             =   720
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Eval_Code"
            Text            =   ""
            Object.DataMember      =   "Eval_List"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSDataListLib.DataCombo dcBundleNo 
            Bindings        =   "issue.frx":0016
            Height          =   315
            Left            =   2520
            TabIndex        =   3
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Style           =   2
            ListField       =   "Bundle_No"
            Text            =   ""
            Object.DataMember      =   "Bundle_Combo"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label2 
            Caption         =   "Select Evaluator Code"
            Height          =   255
            Left            =   360
            TabIndex        =   6
            Top             =   720
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Select Bundle Number"
            Height          =   255
            Left            =   360
            TabIndex        =   5
            Top             =   240
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "IssueForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RefreshBundles()
  
  dcBundleNo.Text = ""
  Set dcBundleNo.RowSource = Nothing
  dcBundleNo.Refresh
  ProjEnv.rsBundle_Combo.Requery
  Set dcBundleNo.RowSource = ProjEnv.rsBundle_Combo
  dcBundleNo.ListField = "Bundle_No"
  dcBundleNo.RowMember = Bundle_Combo
  dcBundleNo.ReFill
  dcBundleNo.Refresh
  CheckIfBundleEmpty
  
  
End Sub
Private Sub RefreshEval()
  dcEvalCode.Text = ""
  Set dcEvalCode.RowSource = Nothing
  dcEvalCode.Refresh
  ProjEnv.rsEval_List.Requery
  Set dcEvalCode.RowSource = ProjEnv.rsEval_List
  dcEvalCode.ListField = "Eval_Code"
  dcEvalCode.RowMember = Eval_List
  dcEvalCode.ReFill
  dcEvalCode.Refresh
  CheckIfEvalEmpty
  
End Sub
Private Sub cmdIssue_Click()
  Dim Unchecked As Integer
  Unchecked = Val(Text4.Text) - Val(Text5.Text)
  Dim day, temp1
  Dim evalchecked As Integer, temp As String
  day = Format(Now, "Short Date")
  ProjEnv.Eval_Daycopies day, dcEvalCode.Text
  temp1 = ProjEnv.rsEval_Daycopies("Total")
  If IsNull(temp1) = False Then
  evalchecked = Val(temp1)
  Else
  evalchecked = 0
  End If
  ProjEnv.rsEval_Daycopies.Close
  On Error GoTo Handler:
  If (evalchecked + Unchecked > 60) Then
  SubjectForm.ErrorMsg "This evaluator cannot check more than 60 copies in a day.Please select a different evaluator or bundle to assign"
  RefreshEval
  Else
  
    ProjEnv.Insert_Checking dcEvalCode.Text, dcBundleNo.Text, Now
    ProjEnv.Update_Bundle dcEvalCode.Text, dcBundleNo.Text
    ProjEnv.Update_Eval_Issue dcEvalCode.Text
    RefreshBundles
    RefreshEval
  
  End If
  Exit Sub
Handler:
  SubjectForm.ErrorMsg "An error occurred"

End Sub

Private Sub dcBundleNo_Click(Area As Integer)
Dim evalcode

If Area = 2 Then
CheckIfBundleEmpty
ProjEnv.Bundle_details dcBundleNo.Text
Text1.Text = ProjEnv.rsBundle_Details("Subject_Code")
Text2.Text = ProjEnv.rsBundle_Details("Start_Serial")
Text3.Text = ProjEnv.rsBundle_Details("End_Serial")
Text4.Text = ProjEnv.rsBundle_Details("Total")
Text5.Text = ProjEnv.rsBundle_Details("Copies_Checked")
evalcode = ProjEnv.rsBundle_Details("Eval_Code")
ProjEnv.rsBundle_Details.Close
If IsNull(evalcode) = False Then
dcEvalCode.Text = evalcode
dcEvalCode.Enabled = False
ProjEnv.Eval_Details dcEvalCode.Text
Text6.Text = ProjEnv.rsEval_Details("Eval_Name")
Text7.Text = ProjEnv.rsEval_Details("Eval_Phone")
ProjEnv.rsEval_Details.Close
Else
dcEvalCode.Enabled = True
End If
End If
End Sub
Private Sub CheckIfEvalEmpty()
  If ProjEnv.rsEval_List.RecordCount = 0 Then
  SubjectForm.ErrorMsg "There are no evaluators free right now."
  dcEvalCode.Enabled = False
  cmdIssue.Enabled = False
  dcBundleNo.Enabled = False
  Else
  dcEvalCode.Enabled = True
  cmdIssue.Enabled = True
  dcBundleNo.Enabled = True
  End If
End Sub
Private Sub CheckIfBundleEmpty()
  If ProjEnv.rsBundle_Combo.RecordCount = 0 Then
  SubjectForm.ErrorMsg "There are no more bundles to be issued"
  dcBundleNo.Enabled = False
  cmdIssue.Enabled = False
  dcEvalCode.Enabled = False
  Else
  dcBundleNo.Enabled = True
  cmdIssue.Enabled = True
  dcEvalCode.Enabled = True
  End If
End Sub

Private Sub dcEvalCode_Click(Area As Integer)
If Area = 2 Then
CheckIfEvalEmpty
ProjEnv.Eval_Details dcEvalCode.Text
Text6.Text = ProjEnv.rsEval_Details("Eval_Name")
Text7.Text = ProjEnv.rsEval_Details("Eval_Phone")
ProjEnv.rsEval_Details.Close
End If
End Sub


Private Sub Form_Load()
  
  CheckIfBundleEmpty
  CheckIfEvalEmpty
End Sub
