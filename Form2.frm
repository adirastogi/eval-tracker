VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7080
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   7080
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Parameters"
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   9135
      Begin MSDataListLib.DataCombo DataCombo1 
         Bindings        =   "Form2.frx":0000
         DataField       =   "Eval_Code"
         DataMember      =   "Eval_Codes"
         DataSource      =   "ProjEnv"
         Height          =   315
         Left            =   1920
         TabIndex        =   10
         Top             =   840
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   7440
         TabIndex        =   8
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   4680
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate Report"
         Height          =   375
         Left            =   3960
         TabIndex        =   6
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0.00;(0.00)"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   1
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label4 
         Caption         =   "Evaluator Code"
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Min Copies"
         Height          =   375
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Conveyance"
         Height          =   375
         Left            =   3600
         TabIndex        =   4
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Remuneration Rate"
         Height          =   375
         Left            =   360
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
   End
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer91 
      Height          =   7485
      Left            =   120
      TabIndex        =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   9165
      lastProp        =   500
      _cx             =   16166
      _cy             =   13203
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   0   'False
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   0   'False
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   0   'False
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   0   'False
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Report As New CrystalReport1

Private Sub Command1_Click()

Dim Rate, Conv, MinCopies As Double
Rate = Val(Text1.Text)
Conv = Val(Text2.Text)
MinCopies = Val(Text3.Text)
If DataCombo1.Text = "" Then
  SubjectForm.ErrorMsg "Please Enter an Evaluator Code."
Else
ProjEnv.Get_Eval_Bill Text1.Text, DataCombo1.Text
Report.Database.Tables.Item(1).SetDataSource ProjEnv.rsGet_Eval_Bill
CRViewer91.ReportSource = Report
CRViewer91.ViewReport


End If

End Sub

Private Sub Form_Load()
Screen.MousePointer = vbHourglass

CRViewer91.ReportSource = Report

Screen.MousePointer = vbDefault

End Sub

Private Sub Form_Resize()
CRViewer91.Top = 0
CRViewer91.Left = 0
CRViewer91.Height = ScaleHeight
CRViewer91.Width = ScaleWidth

End Sub





