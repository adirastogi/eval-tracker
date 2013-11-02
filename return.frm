VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form ReturnForm 
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
      Caption         =   "RETURN AN ISSUED BUNDLE"
      Height          =   6615
      Left            =   720
      TabIndex        =   0
      Top             =   360
      Width           =   8175
      Begin VB.CommandButton cmdReturn 
         Caption         =   "&Return Bundle"
         Height          =   615
         Left            =   3120
         TabIndex        =   20
         Top             =   5760
         Width           =   1935
      End
      Begin VB.Frame Frame4 
         Caption         =   "Evaluator Details"
         Height          =   3255
         Left            =   4320
         TabIndex        =   5
         Top             =   2280
         Width           =   3615
         Begin VB.TextBox Text8 
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
            TabIndex        =   23
            Top             =   720
            Width           =   2895
         End
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
            TabIndex        =   17
            Top             =   2640
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
            TabIndex        =   16
            Top             =   1680
            Width           =   2895
         End
         Begin VB.Label Label7 
            Caption         =   "Evaluator Code"
            Height          =   375
            Left            =   360
            TabIndex        =   24
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label6 
            Caption         =   "Phone Number"
            Height          =   375
            Left            =   360
            TabIndex        =   19
            Top             =   2400
            Width           =   1575
         End
         Begin VB.Label Label5 
            Caption         =   "Evaluator Name"
            Height          =   375
            Left            =   360
            TabIndex        =   18
            Top             =   1440
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
            TabIndex        =   10
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
            TabIndex        =   9
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
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   360
            Width           =   1455
         End
         Begin VB.Label Label4 
            Caption         =   "No of Copies Already Checked "
            Height          =   495
            Index           =   3
            Left            =   240
            TabIndex        =   15
            Top             =   2640
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Total Number of Copies"
            Height          =   375
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "End Serial Number"
            Height          =   375
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   1560
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Start Serial Number"
            Height          =   375
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Subject Code"
            Height          =   375
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Select Bundle"
         Height          =   1455
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   7695
         Begin VB.OptionButton Option1 
            Caption         =   "Enter Number"
            Height          =   255
            Left            =   2640
            TabIndex        =   26
            Top             =   960
            Value           =   -1  'True
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Complete Bundle"
            Height          =   255
            Left            =   6000
            TabIndex        =   25
            Top             =   960
            Width           =   1695
         End
         Begin VB.TextBox txtCopiesChecked 
            Height          =   375
            Left            =   4080
            TabIndex        =   22
            Top             =   840
            Width           =   1575
         End
         Begin MSDataListLib.DataCombo dcBundleNo 
            Bindings        =   "return.frx":0000
            Height          =   315
            Left            =   2640
            TabIndex        =   21
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            ListField       =   "Bundle_No"
            Text            =   ""
            Object.DataMember      =   "Bundle_Retlist"
         End
         Begin VB.Label Label2 
            Caption         =   "Enter No. of Copies Checked"
            Height          =   255
            Left            =   360
            TabIndex        =   4
            Top             =   960
            Width           =   2535
         End
         Begin VB.Label Label1 
            Caption         =   "Select Bundle Number"
            Height          =   255
            Left            =   360
            TabIndex        =   3
            Top             =   360
            Width           =   1815
         End
      End
   End
End
Attribute VB_Name = "ReturnForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub RefreshBundles()
  
  dcBundleNo.Text = ""
  Set dcBundleNo.RowSource = Nothing
  dcBundleNo.Refresh
  ProjEnv.rsBundle_Retlist.Requery
  Set dcBundleNo.RowSource = ProjEnv.rsBundle_Retlist
  dcBundleNo.ListField = "Bundle_No"
  dcBundleNo.RowMember = Bundle_Combo
  dcBundleNo.ReFill
  dcBundleNo.Refresh
  CheckIfBundleEmpty
  
  
End Sub
Private Function GetNoCopies() As String
   If Option1.Value = True Then
    GetNoCopies = txtCopiesChecked.Text
   Else
    GetNoCopies = Text4.Text
   End If
   
End Function

  
Private Sub cmdReturn_Click()
 Dim Unchecked As Integer
 Unchecked = Val(Text4.Text) - Val(Text5.Text)
 
   If GetNoCopies = "" Then
   SubjectForm.ErrorMsg "Please enter the no of copies checked"
   Else
     If Val(GetNoCopies) > Unchecked Then
        SubjectForm.ErrorMsg "The number of copies checked cannot be greater than the number of unchecked copies "
        txtCopiesChecked.Text = ""
     Else
       ' On Error GoTo Handler
                ProjEnv.Bundle_Retupdate GetNoCopies, dcBundleNo.Text
                ProjEnv.Log_Retupdate GetNoCopies, Now, dcBundleNo.Text, Text8.Text
                ProjEnv.Eval_Retupdate Text8.Text
                txtCopiesChecked.Text = ""
                RefreshBundles
       ' Exit Sub

     End If
   
   
   End If
'Handler:
   'SubjectForm.ErrorMsg "An error occured."
End Sub

  'Exit Sub
'Handler:
  'SubjectForm.ErrorMsg "An error occurred"

'End Sub

Private Sub dcBundleNo_Click(Area As Integer)


If Area = 2 Then
CheckIfBundleEmpty
ProjEnv.Bundle_details dcBundleNo.Text
Text1.Text = ProjEnv.rsBundle_Details("Subject_Code")
Text2.Text = ProjEnv.rsBundle_Details("Start_Serial")
Text3.Text = ProjEnv.rsBundle_Details("End_Serial")
Text4.Text = ProjEnv.rsBundle_Details("Total")
Text5.Text = ProjEnv.rsBundle_Details("Copies_Checked")
Text8.Text = ProjEnv.rsBundle_Details("Eval_Code")
ProjEnv.Eval_Details Text8.Text
Text6.Text = ProjEnv.rsEval_Details("Eval_Name")
Text7.Text = ProjEnv.rsEval_Details("Eval_Phone")
ProjEnv.rsEval_Details.Close
ProjEnv.rsBundle_Details.Close


End If
End Sub

Private Sub CheckIfBundleEmpty()
  If ProjEnv.rsBundle_Retlist.RecordCount = 0 Then
  SubjectForm.ErrorMsg "There are no bundles to be returned"
  dcBundleNo.Enabled = False
  txtCopiesChecked.Enabled = False
  txtCopiesChecked.Enabled = False
  cmdReturn.Enabled = False
  Else
  dcBundleNo.Enabled = True
  txtCopiesChecked.Enabled = True
  cmdReturn.Enabled = True
  End If
End Sub

Private Sub Form_Load()
  CheckIfBundleEmpty
End Sub

Private Sub Option1_Click()
 txtCopiesChecked.BackColor = &HFFFFFF
 txtCopiesChecked.Enabled = True
End Sub

Private Sub Option2_Click()
 txtCopiesChecked.BackColor = &H8000000F
 txtCopiesChecked.Enabled = False
 txtCopiesChecked.Text = ""
End Sub

Private Sub txtCopiesChecked_KeyPress(KeyAscii As Integer)
 SubjectForm.CheckNum KeyAscii
End Sub
