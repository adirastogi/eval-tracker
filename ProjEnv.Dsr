VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} ProjEnv 
   ClientHeight    =   9780
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10290
   _ExtentX        =   18150
   _ExtentY        =   17251
   FolderFlags     =   5
   TypeLibGuid     =   "{CBA27E1F-5EBB-43C1-BC41-A1525657A23E}"
   TypeInfoGuid    =   "{CEDD1820-18B9-4DE4-B2FA-8EA71BAA82A7}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "SubjectInsert"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   $"ProjEnv.dsx":0000
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   19
   BeginProperty Recordset1 
      CommandName     =   "Insert_Subject"
      CommDispId      =   1002
      RsDispId        =   -1
      CommandText     =   $"ProjEnv.dsx":0091
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   5
      BeginProperty P1 
         RealName        =   "Subcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Subname"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "NoCopies"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Startbun"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P5 
         RealName        =   "Endbun"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "Insert_Bundle"
      CommDispId      =   1009
      RsDispId        =   -1
      CommandText     =   "INSERT INTO Bundle(Bundle_No,Subject_Code,Start_Serial,End_Serial) VALUES(Bunno,Subcode,Startserial,Endserial)"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "Bunno"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "Subcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "Startserial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "Endserial"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "Bundle_Combo"
      CommDispId      =   1011
      RsDispId        =   1094
      CommandText     =   "SELECT Bundle_No FROM Bundle WHERE Issued = 0 AND Copies_Checked < (End_Serial-Start_Serial+1)"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Bundle_No"
         Caption         =   "Bundle_No"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset4 
      CommandName     =   "Bundle_Details"
      CommDispId      =   1015
      RsDispId        =   1033
      CommandText     =   $"ProjEnv.dsx":011B
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   6
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Subject_Code"
         Caption         =   "Subject_Code"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Start_Serial"
         Caption         =   "Start_Serial"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "End_Serial"
         Caption         =   "End_Serial"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Copies_Checked"
         Caption         =   "Copies_Checked"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Eval_Code"
         Caption         =   "Eval_Code"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Total"
         Caption         =   "Total"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "bundleno"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset5 
      CommandName     =   "Eval_List"
      CommDispId      =   1034
      RsDispId        =   1040
      CommandText     =   "SELECT Eval_Code FROM Evaluator WHERE Occupied = 0"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Eval_Code"
         Caption         =   "Eval_Code"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset6 
      CommandName     =   "Eval_Details"
      CommDispId      =   1041
      RsDispId        =   1045
      CommandText     =   $"ProjEnv.dsx":01AF
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Eval_Name"
         Caption         =   "Eval_Name"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Eval_Phone"
         Caption         =   "Eval_Phone"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset7 
      CommandName     =   "Insert_Checking"
      CommDispId      =   1046
      RsDispId        =   -1
      CommandText     =   "INSERT INTO Checking_Log(Eval_Code,Bundle_No,Start_Time) VALUES(evalcode,bundleno,starttime)"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "bundleno"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "starttime"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset8 
      CommandName     =   "Update_Bundle"
      CommDispId      =   1057
      RsDispId        =   -1
      CommandText     =   $"ProjEnv.dsx":01FA
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "bundleno"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset9 
      CommandName     =   "Update_Eval_Issue"
      CommDispId      =   1059
      RsDispId        =   -1
      CommandText     =   "UPDATE Evaluator SET Occupied = 1 WHERE Eval_Code = evalcode"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset10 
      CommandName     =   "Bundle_Retlist"
      CommDispId      =   1061
      RsDispId        =   1065
      CommandText     =   "SELECT Bundle_No From Bundle WHERE Issued = 1"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Bundle_No"
         Caption         =   "Bundle_No"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset11 
      CommandName     =   "Bundle_Retupdate"
      CommDispId      =   1066
      RsDispId        =   -1
      CommandText     =   "UPDATE Bundle SET Copies_Checked = Copies_Checked + checked, Issued = 0 WHERE Bundle_No = bundleno"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "checked"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "bundleno"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset12 
      CommandName     =   "Log_Retupdate"
      CommDispId      =   1073
      RsDispId        =   -1
      CommandText     =   $"ProjEnv.dsx":024D
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   4
      BeginProperty P1 
         RealName        =   "checked"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "time"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "bundleno"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P4 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset13 
      CommandName     =   "Eval_Retupdate"
      CommDispId      =   1079
      RsDispId        =   -1
      CommandText     =   "UPDATE Evaluator SET Occupied = 0 WHERE Eval_Code = evalcode "
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset14 
      CommandName     =   "Eval_Daycopies"
      CommDispId      =   1089
      RsDispId        =   1093
      CommandText     =   $"ProjEnv.dsx":02DE
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Total"
         Caption         =   "Total"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "day"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset15 
      CommandName     =   "Insert_Eval"
      CommDispId      =   1101
      RsDispId        =   -1
      CommandText     =   $"ProjEnv.dsx":035D
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "code"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "name"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "phone"
         Direction       =   1
         Precision       =   255
         Scale           =   255
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset16 
      CommandName     =   "Bundle_Slips"
      CommDispId      =   1103
      RsDispId        =   1107
      CommandText     =   $"ProjEnv.dsx":03B0
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Bundle_No"
         Caption         =   "Bundle_No"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Subject_Code"
         Caption         =   "Subject_Code"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Start_Serial"
         Caption         =   "Start_Serial"
      EndProperty
      BeginProperty Field4 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "End_Serial"
         Caption         =   "End_Serial"
      EndProperty
      BeginProperty Field5 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Total"
         Caption         =   "Total"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset17 
      CommandName     =   "Get_Eval_Bill"
      CommDispId      =   1108
      RsDispId        =   1112
      CommandText     =   $"ProjEnv.dsx":041F
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      IsRSReturning   =   -1  'True
      NumFields       =   7
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Bundle_No"
         Caption         =   "Bundle_No"
      EndProperty
      BeginProperty Field2 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "No_of_Copies"
         Caption         =   "No_of_Copies"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Subject_Code"
         Caption         =   "Subject_Code"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Subject_Name"
         Caption         =   "Subject_Name"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Day_of_Checking"
         Caption         =   "Day_of_Checking"
      EndProperty
      BeginProperty Field6 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "Checked_By_Eval"
         Caption         =   "Checked_By_Eval"
      EndProperty
      BeginProperty Field7 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Renumeration"
         Caption         =   "Renumeration"
      EndProperty
      NumGroups       =   0
      ParamCount      =   2
      BeginProperty P1 
         RealName        =   "rate"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "evalcode"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset18 
      CommandName     =   "Get_Conv_Detail"
      CommDispId      =   1113
      RsDispId        =   1117
      CommandText     =   $"ProjEnv.dsx":05C5
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "DAY"
         Caption         =   "DAY"
      EndProperty
      BeginProperty Field2 
         Precision       =   15
         Size            =   8
         Scale           =   0
         Type            =   5
         Name            =   "Checked"
         Caption         =   "Checked"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Conveyance"
         Caption         =   "Conveyance"
      EndProperty
      NumGroups       =   0
      ParamCount      =   3
      BeginProperty P1 
         RealName        =   "conv"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P2 
         RealName        =   "code"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      BeginProperty P3 
         RealName        =   "arg"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   510
         DataType        =   202
         HostType        =   8
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset19 
      CommandName     =   "Eval_Codes"
      CommDispId      =   1118
      RsDispId        =   1122
      CommandText     =   "SELECT Eval_Code FROM Evaluator"
      ActiveConnectionName=   "SubjectInsert"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   1
      BeginProperty Field1 
         Precision       =   0
         Size            =   255
         Scale           =   0
         Type            =   202
         Name            =   "Eval_Code"
         Caption         =   "Eval_Code"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "ProjEnv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
