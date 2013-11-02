VERSION 5.00
Begin VB.Form SubjectForm 
   Caption         =   "Form1"
   ClientHeight    =   8700
   ClientLeft      =   2025
   ClientTop       =   1035
   ClientWidth     =   10440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8700
   ScaleWidth      =   10440
   Begin VB.Frame Frame1 
      Caption         =   "ADD SUBJECT DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7935
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   6495
      Begin VB.CommandButton cmdReset 
         Caption         =   "Reset"
         Height          =   495
         Left            =   3600
         TabIndex        =   17
         Top             =   6840
         Width           =   1575
      End
      Begin VB.CommandButton cmdSubmit 
         Caption         =   "&Submit"
         Default         =   -1  'True
         Height          =   495
         Left            =   1560
         TabIndex        =   16
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Frame Frame4 
         Caption         =   "Copy Series Number"
         Height          =   1095
         Left            =   360
         TabIndex        =   3
         Top             =   5040
         Width           =   5775
         Begin VB.TextBox txtManCopy 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   285
            Left            =   3720
            TabIndex        =   12
            Top             =   600
            Width           =   1455
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Enter Manually"
            Height          =   255
            Left            =   840
            TabIndex        =   11
            Top             =   720
            Width           =   1695
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Automatically Generated"
            Height          =   255
            Left            =   840
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   2055
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Bundle Number"
         Height          =   1215
         Left            =   360
         TabIndex        =   2
         Top             =   3360
         Width           =   5775
         Begin VB.TextBox txtManBundle 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   375
            Left            =   3240
            TabIndex        =   9
            Top             =   720
            Width           =   1215
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Enter Manually"
            Height          =   375
            Left            =   960
            TabIndex        =   8
            Top             =   720
            Width           =   2175
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Automatically Generated"
            Height          =   255
            Left            =   960
            TabIndex        =   7
            Top             =   240
            Value           =   -1  'True
            Width           =   2295
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Subject Details"
         Height          =   2415
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   5775
         Begin VB.TextBox txtNoCopies 
            Height          =   375
            Left            =   3000
            TabIndex        =   6
            Top             =   1560
            Width           =   2295
         End
         Begin VB.TextBox txtSubName 
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtSubCode 
            Height          =   375
            Left            =   3000
            TabIndex        =   4
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label3 
            Caption         =   "Number of Copies"
            Height          =   255
            Left            =   600
            TabIndex        =   15
            Top             =   1680
            Width           =   1575
         End
         Begin VB.Label Label2 
            Caption         =   "Subject Name"
            Height          =   375
            Left            =   600
            TabIndex        =   14
            Top             =   960
            Width           =   2055
         End
         Begin VB.Label Label1 
            Caption         =   "Subject Code"
            Height          =   375
            Left            =   600
            TabIndex        =   13
            Top             =   360
            Width           =   1935
         End
      End
   End
End
Attribute VB_Name = "SubjectForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdReset_Click()
ClearControls
End Sub
Private Function ClearControls()
Dim ctrl As Control
For Each ctrl In SubjectForm.Controls
If TypeOf ctrl Is TextBox Then
ctrl.Text = ""
End If
Next
txtSubCode.SetFocus

End Function


Private Sub cmdSubmit_Click()
 Dim x, n, i, incr, NextBun, NextCopy As Long
 n = Val(txtNoCopies.Text)
   If n Mod 30 = 0 Then
      x = n \ 30
   Else
      x = (n \ 30) + 1
   End If
   NextBun = GetNextBundle
   NextCopy = GetNextCopy
  On Error GoTo Handler
   If NextBun > 0 And NextCopy > 0 Then
   
   

   
   ProjEnv.Insert_Subject txtSubCode.Text, txtSubName.Text, txtNoCopies.Text, Str(NextBun), Str(NextBun + x - 1)
    For i = 0 To x - 1
   
            If (i = x - 1) Then
              incr = n Mod 30
            Else
              incr = 30
            End If
            
            ProjEnv.Insert_Bundle Str(NextBun), txtSubCode.Text, Str(NextCopy), Str(NextCopy + incr - 1)
               
            NextCopy = NextCopy + incr
            NextBun = NextBun + 1
            
   Next
   
   NextCopy = NextCopy - 1
   NextBun = NextBun - 1
   WriteToConfig NextBun, NextCopy
   ClearControls
   
  End If
  Exit Sub
Handler:

  ErrorMsg "An error occured "
   
End Sub
Public Function WriteToConfig(ByVal bundle As Long, ByVal copy As Long)
Open App.Path & "\config.txt" For Output As #2
Write #2, bundle
Write #2, copy
Close #2
End Function
Public Function GetNextCopy() As Long
Dim PrevBundle, PrevCopy, temp As Long

Open App.Path & "\config.txt" For Input As #1 'Open file for input.
  Input #1, PrevBundle 'Read Previous Bundle
  Input #1, PrevCopy 'Read Prev Copy
 Close #1
 
 If Option3.Value = False Then
  temp = Val(txtManCopy.Text)
  If (temp <= PrevCopy) Then
    ErrorMsg "Invalid. Value must be greater than " & Str(PrevCopy)
    txtManCopy.Text = ""
    GetNextCopy = -1
  Else
    GetNextCopy = temp
  End If
Else
GetNextCopy = PrevCopy + 1
End If
 
End Function
Public Function GetNextBundle() As Long
Dim PrevBundle, temp As Long

Open App.Path & "\config.txt" For Input As #1 'Open file for input.
  Input #1, PrevBundle 'Read Previous Bundle
Close #1 'Close file.

If Option1.Value = False Then
  temp = Val(txtManBundle.Text)
  If (temp <= PrevBundle) Then
    ErrorMsg "Invalid. Value must be greater than " & Str(PrevBundle)
    txtManBundle.Text = ""
    GetNextBundle = -1
  Else
    GetNextBundle = temp
  End If
Else
GetNextBundle = PrevBundle + 1
End If

End Function
Public Function ErrorMsg(Msg As String)
 MsgBox Msg & Err.Description
 
End Function






Private Sub Option1_Click()
  txtManBundle.BackColor = &H8000000F
  txtManBundle.Enabled = False
  txtManBundle.Text = ""
End Sub

Private Sub Option2_Click()
  txtManBundle.BackColor = &HFFFFFF
  txtManBundle.Enabled = True
End Sub

Private Sub Option3_Click()
 txtManCopy.BackColor = &H8000000F
 txtManCopy.Enabled = False
 txtManCopy.Text = ""
End Sub

Private Sub Option4_Click()
 txtManCopy.BackColor = &HFFFFFF
 txtManCopy.Enabled = True
End Sub





Private Sub txtManBundle_KeyPress(KeyAscii As Integer)
CheckNum KeyAscii
End Sub

Private Sub txtManCopy_KeyPress(KeyAscii As Integer)
CheckNum KeyAscii
End Sub

Private Sub txtNoCopies_KeyPress(KeyAscii As Integer)
   CheckNum KeyAscii
   
End Sub

Public Function CheckNum(ByRef KeyAscii As Integer)
 Select Case KeyAscii
    Case 8, 48 To 57  'TAB, -, 0-9
      'don't do anything, these are all okay

    Case Else
      KeyAscii = 0   'this kills any other keystrokes

  End Select
End Function
Public Function ToUpper(ByRef ctrl As TextBox)
ctrl.Text = UCase(ctrl.Text)
ctrl.SelStart = Len(ctrl.Text) + 1
End Function

Private Sub txtSubCode_Change()
ToUpper txtSubCode
End Sub

Private Sub txtSubName_Change()
ToUpper txtSubName
End Sub
