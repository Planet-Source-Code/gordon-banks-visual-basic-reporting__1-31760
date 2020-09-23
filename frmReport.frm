VERSION 5.00
Begin VB.Form frmReport 
   Caption         =   "Reports"
   ClientHeight    =   2250
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4170
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4170
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   2220
      TabIndex        =   3
      Top             =   1470
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run"
      Height          =   495
      Left            =   870
      TabIndex        =   2
      Top             =   1485
      Width           =   1215
   End
   Begin VB.OptionButton Option2 
      Caption         =   "Unbound Report"
      Height          =   495
      Left            =   300
      TabIndex        =   1
      Top             =   870
      Width           =   2190
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Database Bound Report"
      Height          =   495
      Left            =   330
      TabIndex        =   0
      Top             =   360
      Value           =   -1  'True
      Width           =   2115
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*  This simple project shows how to incorporate reports into a     *
'*  project showing both Bound and Unbound Reports.                 *
'*  The Unbound Report data could be retrieved in one SQL statement *
'*  But I wanted to show how to run thru 2 recordsets and bring the *
'*  information into one recordset.                                 *
'*                                                                  *
'*  The Program uses an ADODB connection to the Database            *
'********************************************************************

Option Explicit

Dim rptBound     As New BoundReport   'Declare Report Object on Report Created
Dim rptUnbound   As New UnboundReport 'Declare Report Object on Report Created
Dim SelRpt       As String            'Variable to control Report Selection
Dim rptdata      As ADODB.Recordset   'Create Recordset for passing to the report
Dim rptDBUsers   As ADODB.Recordset   'Create Recordset to fetch data from Database
Dim rptDBProfile As ADODB.Recordset   'Create Recordset to fetch data from Database
Dim dbConnect    As ADODB.Connection  'Create Connection to Database

Private Sub Command1_Click()
    Select Case SelRpt
        Case "Bound"
            'Show Report
            rptBound.Show
        Case "Unbound"
            'Create Fields in Recordset to Match Report Fields
            rptdata.Fields.Append "User", adChar, 20, adFldFixed
            rptdata.Fields.Append "Profile", adChar, 30, adFldFixed
            rptdata.Fields.Append "ProfileName", adChar, 40, adFldFixed
            
            'check to See in Recordset is already open
            If rptDBUsers.State = adStateOpen Then rptDBUsers.Close
            'Fetch Data
            rptDBUsers.Open "Select * from Users", dbConnect, adOpenDynamic
            'Move to First Record
            If Not rptDBUsers.EOF Then
                rptDBUsers.MoveFirst
            End If
            'Open Recordset for Data
            rptdata.Open
            'Loop thru Database information
            While Not rptDBUsers.EOF
                'check to See in Recordset is already open
                If rptDBProfile.State = adStateOpen Then rptDBProfile.Close
                'Fetch Data
                rptDBProfile.Open "Select ProfileName from Profile where Profile = '" & rptDBUsers.Fields("Profile") & "'", dbConnect, adOpenDynamic
                'Move to First Record
                If Not rptDBProfile.EOF Then
                    rptDBProfile.MoveFirst
                End If
                'Create New Record
                rptdata.AddNew
                'Populate Recorset for Report
                rptdata.Fields("User") = rptDBUsers.Fields("UserName")
                rptdata.Fields("Profile") = rptDBUsers.Fields("Profile")
                rptdata.Fields("ProfileName") = rptDBProfile.Fields("ProfileName")
                'Confirm Update
                rptdata.Update
                'Move to Next User
                rptDBUsers.MoveNext
            Wend
            'Pass Recordset to report
            Set rptUnbound.DataSource = rptdata
            'Show report
            rptUnbound.Show
    End Select
End Sub

Private Sub Command2_Click()
    End
End Sub

Private Sub Form_Load()

    Set rptdata = New ADODB.Recordset       'Initialise Recordset
    Set rptDBUsers = New ADODB.Recordset     'Initialise Recordset
    Set rptDBProfile = New ADODB.Recordset     'Initialise Recordset
    Set dbConnect = New ADODB.Connection    'Initialise DB Connection
    
    'Set connection string to database
    dbConnect.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Users.mdb;Persist Security Info=False"
    'Open Connection
    dbConnect.Open
    
    SelRpt = "Bound"
End Sub

Private Sub Option1_Click()
    SelRpt = "Bound"    'Set Report Variable to Bound Report
End Sub

Private Sub Option2_Click()
    SelRpt = "Unbound"  'Set Report Variable to Unbound Report
End Sub
