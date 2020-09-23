VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} DBenviroment 
   ClientHeight    =   5880
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7770
   _ExtentX        =   13705
   _ExtentY        =   10372
   FolderFlags     =   1
   TypeLibGuid     =   "{FFF0996F-1B9E-11D6-94CA-00E0291DAC4E}"
   TypeInfoGuid    =   "{FFF09970-1B9E-11D6-94CA-00E0291DAC4E}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "DBConnection"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\WINNT\Profiles\Gordon\Desktop\vbcode\Users.mdb;Persist Security Info=False"
      Expanded        =   -1  'True
      QuoteChar       =   96
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   2
   BeginProperty Recordset1 
      CommandName     =   "tblUsers"
      CommDispId      =   1002
      RsDispId        =   1006
      CommandText     =   "Users"
      ActiveConnectionName=   "DBConnection"
      CommandType     =   2
      dbObjectType    =   1
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "UserCode"
         Caption         =   "UserCode"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "UserName"
         Caption         =   "UserName"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Profile"
         Caption         =   "Profile"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "tblProfile"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   "Profile"
      ActiveConnectionName=   "DBConnection"
      CommandType     =   2
      dbObjectType    =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "tblUsers"
      IsRSReturning   =   -1  'True
      NumFields       =   2
      BeginProperty Field1 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "Profile"
         Caption         =   "Profile"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   202
         Name            =   "ProfileName"
         Caption         =   "ProfileName"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "Profile"
         ChildField      =   "Profile"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "DBenviroment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()
    'To set the database location to the same folder as the application
    DBConnection.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Users.mdb;Persist Security Info=False"
End Sub
