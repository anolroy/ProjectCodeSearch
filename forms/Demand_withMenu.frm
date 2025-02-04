VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmDemands_withMenu 
   BackColor       =   &H00FFDFC0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8640
   ClientLeft      =   540
   ClientTop       =   990
   ClientWidth     =   10185
   Icon            =   "Demand_withMenu.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   8640
   ScaleLeft       =   250
   ScaleMode       =   0  'User
   ScaleTop        =   250
   ScaleWidth      =   10185
   Visible         =   0   'False
   Begin Crystal.CrystalReport CR1 
      Left            =   6960
      Top             =   3000
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Demand ID:"
      Height          =   855
      Left            =   3000
      TabIndex        =   62
      Top             =   5640
      Width           =   3015
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   120
         TabIndex        =   64
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton cmdFind 
         Caption         =   "&Find Demand"
         Height          =   375
         Left            =   1560
         TabIndex        =   63
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Scroll Records"
      Height          =   855
      Left            =   120
      TabIndex        =   57
      Top             =   5640
      Width           =   2775
      Begin VB.CommandButton cmdMoveFirst 
         Caption         =   "|<"
         Height          =   375
         Left            =   120
         TabIndex        =   61
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMovePrevious 
         Caption         =   "<"
         Height          =   375
         Left            =   720
         TabIndex        =   60
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMoveNext 
         Caption         =   ">"
         Height          =   375
         Left            =   1440
         TabIndex        =   59
         Top             =   240
         Width           =   615
      End
      Begin VB.CommandButton cmdMoveLast 
         Caption         =   ">|"
         Height          =   375
         Left            =   2040
         TabIndex        =   58
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClearDemands 
      Caption         =   "Clear Demands"
      Height          =   495
      Left            =   7200
      TabIndex        =   56
      Top             =   4200
      Width           =   1215
   End
   Begin VB.CommandButton cmdDeleteOld 
      Caption         =   "Delete &Old Demands"
      Height          =   495
      Left            =   7200
      TabIndex        =   55
      Top             =   4800
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrintThis 
      Caption         =   "Print T&his Demand"
      Height          =   495
      Left            =   7200
      TabIndex        =   54
      Top             =   5400
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelEdit 
      Caption         =   "&Cancel Edit"
      Height          =   495
      Left            =   4200
      TabIndex        =   53
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReprintAll 
      Caption         =   "Reprint &All Demands"
      Height          =   495
      Left            =   2280
      TabIndex        =   52
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReprintSome 
      Caption         =   "Reprint &Selected Demands"
      Height          =   495
      Left            =   4200
      TabIndex        =   51
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintBatch 
      Caption         =   "Print a &Batch of Demands"
      Height          =   495
      Left            =   7200
      TabIndex        =   50
      Top             =   6000
      Width           =   1215
   End
   Begin VB.CheckBox chkPrint 
      BackColor       =   &H00FFDFC0&
      Height          =   255
      Left            =   120
      TabIndex        =   49
      Top             =   5280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.CommandButton cmdCancelReprint 
      Caption         =   "&Cancel Reprint Demands"
      Height          =   495
      Left            =   6120
      TabIndex        =   45
      Top             =   7560
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelPrint 
      Caption         =   "&Cancel Print Demands"
      Height          =   495
      Left            =   6120
      TabIndex        =   44
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintSome 
      Caption         =   "Print &Selected Demands"
      Height          =   495
      Left            =   4200
      TabIndex        =   43
      Top             =   6600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdPrintAll 
      Caption         =   "Print &All Unprinted Demands"
      Height          =   495
      Left            =   2280
      TabIndex        =   42
      Top             =   8040
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdReprint 
      Caption         =   "&Reprint Demands"
      Height          =   495
      Left            =   7200
      TabIndex        =   41
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&Print Demands"
      Height          =   495
      Left            =   5880
      TabIndex        =   40
      Top             =   6600
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenAll 
      Caption         =   "Generate &Automatic Demands"
      Height          =   495
      Left            =   0
      TabIndex        =   39
      Top             =   7080
      Width           =   1575
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   37
      Top             =   1440
      Width           =   1935
   End
   Begin VB.CommandButton cmdCancelNew 
      Caption         =   "&Cancel Manual Demand"
      Height          =   495
      Left            =   5160
      TabIndex        =   9
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdSaveNew 
      Caption         =   "&Save Manual Demand"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   7080
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit Demands"
      Height          =   495
      Left            =   4560
      TabIndex        =   35
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete Demand"
      Height          =   495
      Left            =   3240
      TabIndex        =   34
      Top             =   7080
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerateManual 
      Caption         =   "Generate a &Manual Demand"
      Height          =   495
      Left            =   1680
      TabIndex        =   33
      Top             =   7080
      Width           =   1455
   End
   Begin VB.CheckBox chk3 
      BackColor       =   &H00FFDFC0&
      Enabled         =   0   'False
      Height          =   255
      Left            =   4560
      TabIndex        =   32
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox chk2 
      BackColor       =   &H00FFDFC0&
      Enabled         =   0   'False
      Height          =   255
      Left            =   3240
      TabIndex        =   31
      Top             =   5280
      Width           =   255
   End
   Begin VB.CheckBox chk1 
      BackColor       =   &H00FFDFC0&
      Enabled         =   0   'False
      Height          =   255
      Left            =   1680
      TabIndex        =   30
      Top             =   5280
      Width           =   255
   End
   Begin VB.TextBox txt11 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   30
      MultiLine       =   -1  'True
      TabIndex        =   5
      Top             =   4800
      Width           =   3710
   End
   Begin VB.TextBox txt6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   3
      Top             =   3840
      Width           =   1935
   End
   Begin VB.TextBox txt4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   0
      Top             =   2400
      Width           =   1935
   End
   Begin VB.TextBox txt1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   25
      Top             =   600
      Width           =   3710
   End
   Begin VB.ComboBox cboTenant 
      Height          =   315
      Left            =   1800
      TabIndex        =   24
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
   End
   Begin VB.ComboBox cboDemand 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   22
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox txt3 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   21
      Top             =   960
      Width           =   1335
   End
   Begin VB.TextBox txt2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      TabIndex        =   20
      Top             =   960
      Width           =   1935
   End
   Begin VB.ComboBox cboType 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   2
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txt8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   18
      Top             =   3360
      Width           =   1335
   End
   Begin VB.TextBox txt9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.TextBox txt7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   7
      Top             =   2880
      Width           =   1335
   End
   Begin VB.TextBox txt10 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   100
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   4320
      Width           =   3710
   End
   Begin VB.TextBox txt5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1800
      MaxLength       =   10
      TabIndex        =   1
      Top             =   2880
      Width           =   1935
   End
   Begin VB.Label lbl4 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Excel"
      Height          =   195
      Left            =   4920
      TabIndex        =   48
      Top             =   5280
      Width           =   390
   End
   Begin VB.Label lbl3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Sage"
      Height          =   195
      Left            =   3600
      TabIndex        =   47
      Top             =   5280
      Width           =   375
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Printed"
      Height          =   195
      Left            =   2040
      TabIndex        =   46
      Top             =   5280
      Width           =   495
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Demand ID:"
      Height          =   195
      Left            =   840
      TabIndex        =   38
      Top             =   1440
      Width           =   855
   End
   Begin VB.Label lbl1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Send to Print"
      Height          =   195
      Left            =   480
      TabIndex        =   36
      Top             =   5280
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Sage Text:"
      Height          =   195
      Left            =   915
      TabIndex        =   29
      Top             =   4800
      Width           =   780
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Reference:"
      Height          =   195
      Left            =   900
      TabIndex        =   28
      Top             =   3840
      Width           =   795
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Issue Date:"
      Height          =   195
      Left            =   885
      TabIndex        =   27
      Top             =   2400
      Width           =   810
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Company Name:"
      Height          =   195
      Left            =   525
      TabIndex        =   26
      Top             =   600
      Width           =   1170
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Demand:"
      Height          =   195
      Left            =   1050
      TabIndex        =   23
      Top             =   1920
      Width           =   645
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Type of Demand:"
      Height          =   195
      Left            =   465
      TabIndex        =   19
      Top             =   3360
      Width           =   1230
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Total:"
      Height          =   195
      Left            =   4410
      TabIndex        =   17
      Top             =   3840
      Width           =   405
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "VAT Amount:"
      Height          =   195
      Left            =   3870
      TabIndex        =   16
      Top             =   3360
      Width           =   945
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Amount:"
      Height          =   195
      Left            =   4230
      TabIndex        =   15
      Top             =   2880
      Width           =   585
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Sage A/C Number:"
      Height          =   195
      Left            =   345
      TabIndex        =   14
      Top             =   960
      Width           =   1350
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Unit Number:"
      Height          =   195
      Left            =   3885
      TabIndex        =   13
      Top             =   960
      Width           =   930
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Description:"
      Height          =   195
      Left            =   855
      TabIndex        =   12
      Top             =   4320
      Width           =   840
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Date Due:"
      Height          =   195
      Left            =   960
      TabIndex        =   11
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00FFDFC0&
      Caption         =   "Tenant:"
      Height          =   195
      Left            =   1140
      TabIndex        =   10
      Top             =   120
      Visible         =   0   'False
      Width           =   555
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMain 
         Caption         =   "Main"
      End
      Begin VB.Menu mnuTenants 
         Caption         =   "Tenants"
      End
      Begin VB.Menu mnuUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuLease 
         Caption         =   "Lease"
      End
      Begin VB.Menu mnuGlobal 
         Caption         =   "Global"
      End
      Begin VB.Menu mnuShopCentre 
         Caption         =   "Shopping Centre"
      End
   End
   Begin VB.Menu mnuDemands 
      Caption         =   "Demands"
      Begin VB.Menu mnuGenManual 
         Caption         =   "Generate a Manual Demand"
      End
      Begin VB.Menu mnuDeleteDemand 
         Caption         =   "Delete Demand"
      End
      Begin VB.Menu mnuEditDemands 
         Caption         =   "Edit Demands"
      End
      Begin VB.Menu mnuPrint 
         Caption         =   "Print Demands"
      End
      Begin VB.Menu mnuReprint 
         Caption         =   "Reprint Demands"
      End
      Begin VB.Menu mnuPrintBatch 
         Caption         =   "Print a batch of Demands"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGenAll 
         Caption         =   "Generate all Demands"
      End
   End
End
Attribute VB_Name = "frmDemands_withMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public TenantCode As String
Public TenantName As String
Public Unit As String
Public typeofdemand As Integer
Dim Conn As New RDO.rdoConnection
Dim Env As rdoEnvironment
Dim Envs As rdoEnvironments
Dim Rst1 As rdoResultset
Dim Rst2 As rdoResultset
Dim SQLStr1 As String
Dim SQLStr2 As String
Dim cnnDB As ADODB.Connection
Dim cnnDB2 As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rs1 As ADODB.Recordset
Dim rs2 As ADODB.Recordset
Dim strDBpath
Dim strSQLTitles
Public NomCodeAmount As String
Public NomNameAmount As String
Public NomCodeVAT As String
Public NomNameVAT As String
Public NomCodeTotal As String
Public NomNameTotal As String
Public Prefix As String
Dim DemandId As Long
Public text As String
Public description As String
Public EditMode As Boolean
Public PrintMode As Boolean
Public ReprintMode As Boolean
Public CutOffDate As String
Public temp1 As String
Public temp2 As String
Public temp3 As String
Public temp4 As String
Public temp5 As String
Public DueDate As String
Dim a As String

Public Amount As Double
Public batch As Integer
Public first As Integer
Public last As Integer


Private Sub cboTenant_Click()
'the user has selected a tenant to generate a manual demand for

Dim i, j, match As Integer
Dim temp1 As String
Dim temp2 As String
Dim temp3 As String
match = 0

If cboTenant.text = "" Then
    MsgBox "You must Select a Tenant to generate a demand for.", vbOKOnly + vbCritical, "No Tenant Selected"
    Exit Sub
End If
j = cboTenant.ListCount - 1
For i = 0 To j
    If cboTenant.List(i) = cboTenant.text Then
        match = 1
        Exit For
    End If
Next i
If match = 0 Then
    MsgBox "Then tenant selected is invalid.", vbOKOnly + vbCritical, "Invalid Tenant"
    cboTenant.text = ""
    Exit Sub
End If
'put the tenant company name and sage account numberin the boxes on screen.
For i = 2 To 10
    If Mid(cboTenant.text, i, 3) = " / " Then
        TenantCode = Left(cboTenant.text, i - 1)
        TenantName = Right(cboTenant.text, Len(cboTenant.text) - i - 2)
    End If
Next i

txt1.text = TenantName
txt2.text = TenantCode

'connect to database
Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

'get the unit for the tenant selected
Set Rst1 = Conn.OpenResultset("SELECT CurrentRental FROM Tenants WHERE SageAccountNumber = '" & TenantCode & "'", rdOpenStatic, rdConcurReadOnly)

Unit = Rst1!CurrentRental

Rst1.Close
Conn.Close
'work out current date in dd/mm/yy format and put in the Issue Date and Due Date boxes
temp1 = CStr(Day(Date))
temp2 = CStr(Month(Date))
temp3 = CStr(Year(Date))

If Len(temp1) = 1 Then temp1 = "0" & temp1
If Len(temp2) = 1 Then temp2 = "0" & temp2
If Len(temp3) > 2 Then temp3 = Right(temp3, 2)
txt4.text = temp1 & "/" & temp2 & "/" & temp3
txt5.text = temp1 & "/" & temp2 & "/" & temp3
txt3.text = Unit
txt11.text = "S/L " & TenantCode

cboDemand.text = "Manual"
'cboTenant.Enabled = False

Call FillCbos
Call EnableBoxes
cmdSaveNew.Visible = True
cmdCancelNew.Visible = True
cboType.SetFocus

End Sub

Private Sub cboType_LostFocus()
'the user has just selected the type of demand

Dim i, j, match As Integer
Dim DemandType As String
match = 0

If cboType.text = "" Then
    MsgBox "You must Select a Type of demand", vbOKOnly + vbCritical, "Type of Demand"
    Exit Sub
End If
j = cboType.ListCount - 1
For i = 0 To j
    If cboType.List(i) = cboType.text Then
        match = 1
        Exit For
    End If
Next i
If match = 0 Then
    MsgBox "Type of demand selected is invalid.", vbOKOnly + vbCritical, "Invalid Type of demand"
    cboType.text = ""
    Exit Sub
End If
For i = 2 To 4
    If Mid(cboType.text, i, 3) = " / " Then
        DemandId = CLng(Left(cboType.text, i - 1))
        DemandType = Right(cboType.text, Len(cboType.text) - i - 2)
        Exit For
    End If
Next i
'connect to database
Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

'Get the details for the demand type selected
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '" & DemandId & "'"
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

'if some details have not been set up for demand type selected do not allow the user to continue with demand
If IsNull(Rst1!NominalCodeforAmount) Then
    MsgBox "You have not selected a Nominal Code for This Type of Demand." & Chr(13) & "Please select Nominal Codes for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Nominal Codes"
    Rst1.Close
    Conn.Close
    Exit Sub
End If
If IsNull(Rst1!NominalCodeForVAT) Then
    MsgBox "You have not selected a Nominal Code for this Type of Demand." & Chr(13) & "Please select Nominal Codes for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Nominal Codes"
    Rst1.Close
    Conn.Close
    Exit Sub
End If
If IsNull(Rst1!NominalCodeForTotal) Then
    MsgBox "You have not selected a Nominal Code for this Type of Demand." & Chr(13) & "Please select Nominal Codes for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Nominal Codes"
    Rst1.Close
    Conn.Close
    Exit Sub
End If
If IsNull(Rst1!Prefix) Then
    MsgBox "You have not entered a Sage Prefix for this Type of Demand." & Chr(13) & "Please enter a Sage Prefix for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Sage Prefix"
    Rst1.Close
    Conn.Close
    Exit Sub
End If
'set the details to variables
NomCodeAmount = Rst1!NominalCodeforAmount
NomNameAmount = Rst1!NominalNameforAmount
If DemandId <> 4 Then NomCodeVAT = Rst1!NominalCodeForVAT
If DemandId <> 4 Then NomNameVAT = Rst1!NominalNameforVAT
NomCodeTotal = Rst1!NominalCodeForTotal
NomNameTotal = Rst1!NominalNameforTotal
Prefix = Rst1!Prefix
a = Rst1!InvCrd
Rst1.Close
Conn.Close

Call GetReference

End Sub


Private Sub cmdCancelEdit_Click()
'User selected to cancel edit mode
Call SaveChanges
EditMode = False

Call DisableBoxes
txt8.Enabled = False

cmdCancelEdit.Visible = False
cmdGenerateManual.Visible = True
cmdGenAll.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdEdit.Visible = True
cmdGenAll.Visible = True
cmdPrint.Visible = True
cmdReprint.Visible = True
cmdPrintThis.Visible = True
cmdPrintBatch.Visible = True
Call EnableMenu

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset
'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get details for current demand on screen
strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber =" & CInt(Text1.text)
rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

Call EmptyBoxes
Call GetRecord

End Sub

Private Sub cmdCancelNew_Click()
'the user has selected to cancel add a manual demand
Call EmptyBoxes
Call DisableBoxes

cboTenant.Visible = False
Label1.Visible = False

cmdGenerateManual.Visible = True
cmdGenAll.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdEdit.Visible = True
cmdGenAll.Visible = True
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdPrint.Visible = True
cmdReprint.Visible = True
cmdPrintThis.Visible = True
cmdPrintBatch.Visible = True
cmdMoveFirst.Visible = True
cmdMovePrevious.Visible = True
cmdMoveNext.Visible = True
cmdMoveLast.Visible = True
cmdFind.Visible = True
Text2.Visible = True
lbl4.Visible = True
lbl2.Visible = True
lbl3.Visible = True
chk1.Visible = True
chk2.Visible = True
chk3.Visible = True
Call EnableMenu

Call GetFirstDemand

MsgBox "The manual demand has not been saved.", vbOKOnly + vbInformation, "Cancelled"

End Sub

Private Sub cmdCancelPrint_Click()
'The user has selected to exit Print Mode
Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get sendtoprint field for all demands that were set to be sent to print
strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic
'set the sendtoprint field to not be sent to print - blank
If rs.EOF = False Then
    While rs.EOF = False
        rs!SendToPrint = ""
        rs.Update
        rs.MoveNext
    Wend
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

cmdCancelPrint.Visible = False
cmdPrintAll.Visible = False
cmdPrintBatch.Visible = True
cmdPrintSome.Visible = False
cmdGenAll.Visible = True
cmdGenerateManual.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdPrint.Visible = True
cmdPrintThis.Visible = True
cmdReprint.Visible = True
chkPrint.Visible = False
lbl1.Visible = False
Call EnableMenu

PrintMode = False

Call EmptyBoxes
Call GetFirstDemand

End Sub

Private Sub cmdCancelReprint_Click()
'user selected to exit Reprint mode
Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get the sendtoprint field from demands that were set to be sent to print
strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

'set the sendtoprint field to not be sent to print - blank
If rs.EOF = False Then
    While rs.EOF = False
        rs!SendToPrint = ""
        rs.Update
        rs.MoveNext
    Wend
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

ReprintMode = False

cmdCancelReprint.Visible = False
cmdReprintSome.Visible = False
cmdReprintAll.Visible = False
cmdPrintBatch.Visible = False
cmdPrintAll.Visible = False
cmdPrintBatch.Visible = True
cmdPrintSome.Visible = False
cmdGenAll.Visible = True
cmdGenerateManual.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdPrint.Visible = True
cmdPrintThis.Visible = True
cmdReprint.Visible = True
chkPrint.Visible = False
lbl1.Visible = False
Call EnableMenu

Call EmptyBoxes
Call GetFirstDemand

End Sub

Private Sub cmdDelete_Click()

Call Delete

End Sub

Private Sub cmdDeleteOld_Click()
'user wants to delete all old demands - ones that have been printed, exported to sage and exported to excel
If MsgBox("Do you really want to delete old demands?", vbYesNo + vbQuestion, "Delete Old Demands") = vbNo Then Exit Sub
MousePointer = vbHourglass

Call DeleteDemands
'MsgBox "Old demands deleted successfully", vbOKOnly + vbInformation, "Deleted"
Call EmptyBoxes
Call GetFirstDemand
MsgBox "Old demands deleted successfully", vbOKOnly + vbInformation, "Deleted"
MousePointer = vbDefault

End Sub

Private Sub cmdEdit_Click()

Call Edit

End Sub

Private Sub cmdFind_Click()
'user wants to find a demand
Dim i As Integer
Dim char As String
Dim check As Boolean

If Text2.text = "" Then
    MsgBox "You must select the Demand ID of the demand you want to find!", vbOKOnly + vbCritical, "Find Demand"
    Exit Sub
End If
'check that user entered a number in the text box
check = True
For i = 1 To Len(Text2.text)
    char = Mid(Text2.text, i, 1)
    If Asc(char) < 48 Or Asc(char) > 57 Then check = False
Next i

If check = False Then
    MsgBox "Invalid Demand Reference Number!", vbOKOnly + vbCritical, "Invalid Number"
    Exit Sub
End If

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset
'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get the details of the demand with uniquerefnumber the user entered
strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text2.text)
'for the various modes only select it if the printed and exported to sage status is ok.
If EditMode = True Then strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text2.text) & " AND ExportedToSage = 'N'"
If PrintMode = True Then strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text2.text) & " AND IsPrinted = 'N'"
If ReprintMode = True Then strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text2.text) & " AND IsPrinted = 'Y'"

rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly
'if record exists display on screen if not tell user invalid Demand Id
If rs.EOF = False Then
    Call EmptyBoxes
    Call GetRecord
Else
    MsgBox "You have entered an invalid Demand ID", vbOKOnly + vbCritical, "Invalid Demand Id"
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

Text2.text = ""

End Sub

Private Sub cmdGenAll_Click()

Call GenerateAll

End Sub

Private Sub cmdGenerateManual_Click()

Call AddManualDemand

End Sub

Private Sub cmdMoveFirst_Click()
'move to the first demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made
If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
If EditMode = True Then Call SaveChanges
Call EmptyBoxes
Call GetFirstDemand

End Sub

Private Sub cmdMoveLast_Click()
'move to last demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made
If Text1.text = "" Then
    MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
    Exit Sub
End If
Dim b As Integer
b = 1

If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
If EditMode = True Then Call SaveChanges

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset
'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get all uniquerefnumbers of demands in program
strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords"
'if in edit, print or reprint mode get uniquerefnumbers of demands that have correct printed and exported to sage status
If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE ExportedToSage = 'N'"
If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N'"
If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y'"

rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly
'find the largest uniquerefnumber and set this to b
If rs.EOF = False Then
    While rs.EOF = False
        If b < rs!UniqueRefNumber Then b = rs!UniqueRefNumber
        rs.MoveNext
    Wend
End If

rs.Close

Set rs = Nothing
Set rs = New ADODB.Recordset
'get the demand details for demand with uniquerefnumber b
strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & b
rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

Call EmptyBoxes
Call GetRecord

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

Private Sub cmdMoveNext_Click()
'move to next demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made
If Text1.text = "" Then
    MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
    Exit Sub
End If

Dim b As Integer
Dim last As Boolean
last = False

If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
If EditMode = True Then Call SaveChanges

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset
'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get uniquerefnumbers from demands where uniqueref number is greater than current demand on screen
strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.text)
'if in edit, print or reprint mode get uniquerefnumbers where printed and exported to sage are correct status
If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.text) & " AND ExportedToSage = 'N'"
If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N' AND UniqueRefNumber > " & CInt(Text1.text)
If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y' AND UniqueRefNumber > " & CInt(Text1.text)

rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockPessimistic
'work out which is the smallest uniquerefnumber and set to b
If rs.EOF = False Then
    rs.MoveLast
    rs.MoveFirst
    If rs.RecordCount = 1 Then last = True
    b = rs!UniqueRefNumber
    While rs.EOF = False
        If b > rs!UniqueRefNumber Then b = rs!UniqueRefNumber
        rs.MoveNext
    Wend
Else ' if no records so current demand is last demand
    MsgBox "This is the last demand.", vbOKOnly + vbInformation, "Last Demand"
    rs.Close
    Set rs = Nothing
    cnnDB.Close
    Set cnnDB = Nothing
    Exit Sub
End If
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset
'get details for demand with Uniquerefnumber b
strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & b
rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

If Not rs.EOF = True And rs.BOF = True Then
    MsgBox "This is the last demand", vbOKOnly + vbInformation, "Last Demand"
    rs.Close
    Set rs = Nothing
    cnnDB.Close
    Set cnnDB = Nothing
    Exit Sub
End If

Call EmptyBoxes
Call GetRecord

If last = True Then MsgBox "This is the last demand", vbOKOnly + vbInformation, "Last Demand"

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

Private Sub cmdMovePrevious_Click()

'move to previous demand in database
'if in print or reprint mode need to update the status of
'sendtoprint for demand currently on screen.
'if in edit mode need to prompt to save changes that have been made

If Text1.text = "" Then
    MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
    Exit Sub
End If

Dim b As Integer
Dim a As Integer
a = CInt(Text1.text)
b = 1
If PrintMode = True Or ReprintMode = True Then Call UpdatePrint
If EditMode = True Then Call SaveChanges

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset
'connect to database
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
'get all uniquerefnumbers that are less than that of current demand on screen
strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber < " & CInt(Text1.text)
'if in edit, print or reprint mode get uniquerefnumbers where printed and exported to sage are correct status
'If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.Text) & " AND ExportedToSage = 'N'"
If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber < " & CInt(Text1.text) & " AND ExportedToSage = 'N' "
If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N' AND UniqueRefNumber < " & CInt(Text1.text)
If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y' AND UniqueRefNumber < " & CInt(Text1.text)

rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

'work out which is the biggest of uniquerefnumbers in record set and set to b
If rs.EOF = False Then
    While rs.EOF = False
        If b < rs!UniqueRefNumber Then b = rs!UniqueRefNumber
        rs.MoveNext
    Wend
Else ' no records so set b current demand ref number
    b = a
End If

rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset
'get details of demand with ref number b
strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & b
rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

Call EmptyBoxes
Call GetRecord

If b = a Then MsgBox "This is the first demand", vbOKOnly + vbInformation, "First Demand"

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

Private Sub cmdPrint_Click()

Call PrintDemands

End Sub

Private Sub cmdPrintAll_Click()
'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE IsPrinted = 'N'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        'rs.EditMode
        rs!IsPrinted = "C"
        rs!SendToPrint = ""
        rs.Update
        rs.MoveNext
    Wend
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

CR1.ReportFileName = App.Path & "\Demand" & SCID & ".rpt"
CR1.printReport

Call SetPrintedtoYes
cmdCancelPrint.Visible = False
cmdPrintAll.Visible = False
cmdPrintBatch.Visible = True
cmdPrintSome.Visible = False
cmdGenAll.Visible = True
cmdGenerateManual.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdPrint.Visible = True
cmdReprint.Visible = True
cmdPrintThis.Visible = True
chkPrint.Visible = False
lbl1.Visible = False
Call EnableMenu

PrintMode = False

Call EmptyBoxes
Call GetFirstDemand

End Sub

Private Sub cmdPrintBatch_Click()

Call PrintBatchSelected

End Sub

Private Sub cmdPrintSome_Click()

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

MousePointer = vbHourglass

Call UpdatePrint

Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If Rst1.EOF = False Then
    While Rst1.EOF = False
        Rst1.Edit
        Rst1!IsPrinted = "C"
        Rst1!SendToPrint = ""
        Rst1.Update
        Rst1.MoveNext
    Wend
    Rst1.Close
    Conn.Close
Else
    MsgBox "There are no demands selected to print!", vbOKOnly + vbInformation, "Print Demands"
    Rst1.Close
    Conn.Close
    Exit Sub
End If

'print the selected demands so those that have IsPrinted = 'C'
CR1.ReportFileName = App.Path & "\Demand" & SCID & ".rpt"
CR1.printReport

Call SetPrintedtoYes
cmdCancelPrint.Visible = False
cmdPrintAll.Visible = False
cmdPrintBatch.Visible = True
cmdPrintSome.Visible = False
cmdGenAll.Visible = True
cmdGenerateManual.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdPrint.Visible = True
cmdPrintThis.Visible = True
cmdReprint.Visible = True
chkPrint.Visible = False
lbl1.Visible = False
Call EnableMenu

PrintMode = False

Call EmptyBoxes
Call GetFirstDemand

MousePointer = vbDefault

End Sub

Private Sub cmdPrintThis_Click()

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

MousePointer = vbHourglass

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT IsPrinted FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text1.text)
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    rs!IsPrinted = "C"
    rs.Update
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

CR1.ReportFileName = App.Path & "\Demand" & SCID & ".rpt"
CR1.printReport

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT IsPrinted FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text1.text)
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    rs!IsPrinted = "Y"
    rs.Update
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

MousePointer = vbDefault

End Sub

Private Sub cmdReprint_Click()

Call ReprintDemands

End Sub

Private Sub cmdReprintAll_Click()

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE IsPrinted = 'Y'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        'rs.EditMode
        rs!IsPrinted = "C"
        rs!SendToPrint = ""
        rs.Update
        rs.MoveNext
    Wend
End If
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

CR1.ReportFileName = App.Path & "\Demand" & SCID & ".rpt"
'CR1.Connect = "DSN=WDLimited01;UID=;PWD=;DBQ=<CRWDC>Database=WDLimited01"
CR1.printReport

Call SetPrintedtoYes
ReprintMode = False

cmdCancelReprint.Visible = False
cmdReprintSome.Visible = False
cmdReprintAll.Visible = False
cmdPrintBatch.Visible = False
cmdPrintThis.Visible = True
cmdPrintAll.Visible = False
cmdPrintBatch.Visible = True
cmdPrintSome.Visible = False
cmdGenAll.Visible = True
cmdGenerateManual.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdPrint.Visible = True
cmdReprint.Visible = True
chkPrint.Visible = False
lbl1.Visible = False
Call EnableMenu

Call EmptyBoxes
Call GetFirstDemand

End Sub

Private Sub cmdReprintSome_Click()

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

Call UpdatePrint

Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT IsPrinted, SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If Rst1.EOF = False Then
    While Rst1.EOF = False
        Rst1.Edit
        Rst1!IsPrinted = "C"
        Rst1!SendToPrint = ""
        Rst1.Update
        Rst1.MoveNext
    Wend
    Rst1.Close
    Conn.Close
Else
    MsgBox "There are no demands selected to print!", vbOKOnly + vbInformation, "Print Demands"
    Rst1.Close
    Conn.Close
    Exit Sub
End If

CR1.ReportFileName = App.Path & "\Demand" & SCID & ".rpt"
CR1.printReport

Call SetPrintedtoYes
ReprintMode = False

cmdCancelReprint.Visible = False
cmdReprintSome.Visible = False
cmdReprintAll.Visible = False
cmdPrintBatch.Visible = False
cmdPrintAll.Visible = False
cmdPrintBatch.Visible = True
cmdPrintSome.Visible = False
cmdGenAll.Visible = True
cmdGenerateManual.Visible = True
cmdEdit.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdPrint.Visible = True
cmdReprint.Visible = True
cmdPrintThis.Visible = True
chkPrint.Visible = False
lbl1.Visible = False
Call EnableMenu

Call EmptyBoxes
Call GetFirstDemand

End Sub

Private Sub cmdSaveNew_Click()
Dim emz As String

If txt5.text = "" Then
    MsgBox "You must enter a Due Date!", vbOKOnly + vbCritical, "Due Date Required"
    Exit Sub
End If
If txt6.text = "" Then
    MsgBox "You must enter a Reference!", vbOKOnly + vbCritical, "Reference Required"
    Exit Sub
End If
If txt7.text = "" Then
    MsgBox "You must enter an Amount!", vbOKOnly + vbCritical, "Amount Required"
    Exit Sub
End If
If a = "C" Then
    If Asc(Left(txt7.text, 1)) <> 45 Then
        emz = txt7.text
        txt7.text = "-" & emz
    End If
End If
If a = "I" Then
    If Asc(Left(txt7.text, 1)) = 45 Then
        emz = txt7.text
        txt7.text = Right(emz, Len(emz) - 1)
    End If
End If
If txt10.text = "" Then
    MsgBox "You must enter a Description!", vbOKOnly + vbCritical, "Description Required"
    Exit Sub
End If
If txt11.text = "" Then
    MsgBox "You must enter a Sage Text!", vbOKOnly + vbCritical, "Sage Text Required"
    Exit Sub
End If

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="

'strSQLTitles = "SELECT InvCrd FROM DemandTypes WHERE ID = '" & DemandId & "'"
'rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

'a = rs!InvCrd
'rs.Close
Set rs = Nothing

Set rs = New ADODB.Recordset

strSQLTitles = "SELECT * FROM DemandRecords"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

rs.AddNew
rs!AutomaticManual = "M"
rs!SageAccountNumber = txt2.text
rs!TenantCompanyName = txt1.text
rs!UnitNumber = txt3.text
rs!NominalCodeforAmount = NomCodeAmount
rs!NominalNameforAmount = NomNameAmount
If cboType.ListIndex <> 3 Then rs!NominalCodeForVAT = NomCodeVAT
If cboType.ListIndex <> 3 Then rs!NominalNameforVAT = NomNameVAT
rs!NominalCodeForTotal = NomCodeTotal
rs!NominalNameforTotal = NomNameTotal
rs!Source = 1
rs!TotalAmount = txt9.text
rs!VATAmount = txt8.text
rs!Amount = txt7.text
If txt4.text <> "" Then rs!IssueDate = Left(txt4.text, 6) + Right(txt4.text, 2)
rs!DueDate = Left(txt5.text, 6) + Right(txt5.text, 2)
rs!VATMonth = Month(txt4.text)
rs!typeofdemand = CInt(Left(cboType.text, 1)) - 1
'MsgBox rs!typeofdemand
If a = "C" Then rs!TransactionType = 5
If a = "I" Then rs!TransactionType = 4
rs!Reference = txt6.text
rs!text = txt11.text
rs!description = txt10.text
rs!IsPrinted = "N"
rs!ExportedToSage = "N"
rs!ExportedToExcel = "N"
rs.Update
rs.Close
strSQLTitles = "SELECT * FROM DemandRecords"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic
rs.MoveLast
Text1.text = rs!UniqueRefNumber
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

MsgBox "Your Manual Demand has been Saved.", vbOKCancel + vbInformation, "Saved"
chk1.Value = 0
chk2.Value = 0
chk3.Value = 0

cboTenant.Visible = False
Label1.Visible = False

cmdGenerateManual.Visible = True
cmdGenAll.Visible = True
cmdDelete.Visible = True
cmdDeleteOld.Visible = True
cmdEdit.Visible = True
cmdPrint.Visible = True
cmdReprint.Visible = True
cmdPrintThis.Visible = True
cmdPrintBatch.Visible = True
Call EnableMenu
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdMoveFirst.Visible = True
cmdMovePrevious.Visible = True
cmdMoveNext.Visible = True
cmdMoveLast.Visible = True
cboTenant.Visible = False
'lbl1.Visible = True
lbl4.Visible = True
lbl2.Visible = True
lbl3.Visible = True
chk1.Visible = True
chk2.Visible = True
chk3.Visible = True
'chkPrint.Visible = True
cmdFind.Visible = True
Text2.Visible = True

Call DisableBoxes

End Sub

Private Sub Form_Load()
    Me.Top = 50
    Me.Left = 50

    Dim a As Integer
    
    'Me.Move (Screen.Width - Width) / 2, 0
    Me.Caption = "Demands - " & gCurrentShopCentreName
     
    EditMode = False
    PrintMode = False
    ReprintMode = False

    Call DisableBoxes
    
'   FillCbos() collects all type of Demond and push into cboType Combo Box.
    Call FillCbos
    
    Call GetFirstDemand
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set cnnDB = Nothing

    frmMMain.fraCmdButton.Enabled = True
End Sub

Private Sub mnuDeleteDemand_Click()

Call Delete

End Sub

Private Sub mnuEditDemands_Click()

Call Edit

End Sub

Private Sub mnuExit_Click()

    Call ExitProgram

End Sub

Private Sub mnuGenAll_Click()

    Call GenerateAll

End Sub

Private Sub mnuGenManual_Click()

    Call AddManualDemand

End Sub

Public Sub AddManualDemand()

    Call FillcboTenant
    Label1.Visible = True
    cboTenant.Visible = True
    Call EmptyBoxes
    chk1.Visible = False
    chk2.Visible = False
    chk3.Visible = False
    
    cmdEdit.Visible = False
    cmdGenerateManual.Visible = False
    cmdDelete.Visible = False
    cmdDeleteOld.Visible = False
    cmdCancelEdit.Visible = False
    cmdSaveNew.Visible = False
    cmdCancelNew.Visible = False
    cmdGenAll.Visible = False
    cmdPrint.Visible = False
    cmdPrintThis.Visible = False
    cmdPrintBatch.Visible = False
    cmdReprint.Visible = False
    mnuEditDemands.Enabled = False
    mnuGenManual.Enabled = False
    mnuGenAll.Enabled = False
    mnuPrint.Enabled = False
    mnuReprint.Enabled = False
    cmdMoveFirst.Visible = False
    cmdMoveLast.Visible = False
    cmdMovePrevious.Visible = False
    cmdMoveNext.Visible = False
    lbl1.Visible = False
    lbl2.Visible = False
    lbl3.Visible = False
    lbl4.Visible = False
    cmdFind.Visible = False
    Text2.Visible = False

End Sub

Public Sub EnableBoxes()

    txt4.Enabled = True
    txt5.Enabled = True
    txt6.Enabled = True
    txt7.Enabled = True
    txt10.Enabled = True
    txt11.Enabled = True
    cboType.Enabled = True

End Sub

Public Sub DisableBoxes()
    
    txt4.Enabled = False
    txt5.Enabled = False
    txt6.Enabled = False
    txt7.Enabled = False
    txt10.Enabled = False
    txt11.Enabled = False
    cboType.Enabled = False

End Sub

Public Sub EmptyBoxes()

Text1.text = ""
txt1.text = ""
txt2.text = ""
txt3.text = ""
txt4.text = ""
txt5.text = ""
txt6.text = ""
txt7.text = ""
txt8.text = ""
txt9.text = ""
txt10.text = ""
txt11.text = ""
cboDemand.text = ""
cboType.text = ""

End Sub

Public Sub FillcboTenant()

cboTenant.Clear

Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt
                                            
SQLStr1 = "SELECT CompanyName, SageAccountNumber, Currentrental FROM Tenants ORDER BY CompanyName"
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

If Rst1.EOF = False Then
    While Rst1.EOF = False
        If Rst1!CurrentRental <> "" Then cboTenant.AddItem Rst1!SageAccountNumber & " / " & Rst1!CompanyName
        Rst1.MoveNext
    Wend
End If

Rst1.Close
Conn.Close

End Sub

'*********************
'   FillCbos() collects all type of Demond and push into cboType Combo Box.
'*********************
Public Sub FillCbos()

    cboDemand.Enabled = True
    cboDemand.AddItem "Manual", 0
    cboDemand.AddItem "Automatic", 1
    cboDemand.Enabled = False
    
    Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
    Conn.CursorDriver = rdUseIfNeeded
    Conn.EstablishConnection rdDriverNoPrompt
    
    SQLStr1 = "SELECT ID, Type FROM DemandTypes"
    Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)
    
    If Rst1.EOF = False Then
        While Rst1.EOF = False
            cboType.AddItem Rst1!ID & " / " & Rst1!Type
            Rst1.MoveNext
        Wend
    End If
    
    Rst1.Close
    Conn.Close

End Sub

Private Sub mnuGlobal_Click()

'Load frmGlobal
'frmGlobal.Show
'Unload Me

End Sub

Private Sub mnuLease_Click()
'
'Load frmLease
'frmLease.Show
'Unload Me

End Sub

Private Sub mnuMain_Click()

    'Load frmMain
    'frmMain.Show
    frmMMain.fraCmdButton.Enabled = True
    Unload Me

End Sub

Private Sub mnuPrint_Click()

Call PrintDemands

End Sub

Private Sub mnuPrintBatch_Click()

Call PrintBatchSelected

End Sub

Private Sub mnuShopCentre_Click()

'Load frmShoppingCentre
'frmShoppingCentre.Show
'Unload Me

End Sub

Private Sub mnuTenants_Click()

'    Load frmTenant
'    frmTenant.Show
'    Unload Me

End Sub

Private Sub mnuUnits_Click()

'    Load frmUnit
'    frmUnit.Show
'    Unload Me

End Sub

Private Sub txt4_LostFocus()

    If CheckDate(txt4.text) = False Then txt4.text = ""
    Call GetReference

End Sub

Private Sub txt5_LostFocus()

    If CheckDate(txt5.text) = False Then txt5.text = ""

End Sub

Private Sub txt7_Change()
    If txt7.text <> "" And txt7.text <> "-" Then
        If NumberCheck2(txt7.text) = False Then
            txt7.text = ""
        Else
            If DemandId = 4 Then
                txt9.text = txt7.text
                txt8.text = "0"
            Else
                txt8.text = Round(CDbl(txt7.text) * VatRate / 100, 2)
                txt9.text = CDbl(txt7.text) + CDbl(txt8.text)
            End If
        End If
    End If
End Sub

Public Sub GetRecord()

txt1.text = rs!TenantCompanyName
txt2.text = rs!SageAccountNumber
txt3.text = rs!UnitNumber
Text1.text = rs!UniqueRefNumber
If rs!AutomaticManual = "A" Then cboDemand.text = "Automatic"
If rs!AutomaticManual = "M" Then cboDemand.text = "Manual"
txt4.text = rs!IssueDate
txt5.text = rs!DueDate
If EditMode = True Then typeofdemand = rs!typeofdemand
cboType.text = cboType.List(rs!typeofdemand)

'that is a very messy way of doing this...code added by KC 11:01 9th October
Call cboType_LostFocus

'If txt4.Text <> "" Then
'    Prefix = rs!Prefix
'    If cboType.Text <> "" Then
'        txt6.Text = Prefix & Left(txt4.Text, 2) & Mid(txt4.Text, 4, 2) & Right(txt4.Text, 2)
'    End If
'End If
'txt6.Text = rs!Reference '...this is a bit precarious

txt7.text = rs!Amount
txt8.text = rs!VATAmount
txt9.text = rs!TotalAmount
txt10.text = rs!description
txt11.text = rs!text
If rs!IsPrinted = "Y" Then chk1.Value = 1 And chkPrint.Value = 0
If rs!IsPrinted = "N" Then chk1.Value = 0 And chkPrint.Value = 0
If rs!SendToPrint = "Y" Then chkPrint.Value = 1 Else chkPrint.Value = 0
If rs!ExportedToSage = "Y" Then chk2.Value = 1
If rs!ExportedToSage = "N" Then chk2.Value = 0
If rs!ExportedToExcel = "Y" Then chk3.Value = 1
If rs!ExportedToExcel = "N" Then chk3.Value = 0

End Sub

Public Sub GetReference()

If txt4.text <> "" Then
    If cboType.text <> "" Then
        txt6.text = Prefix & Left(txt4.text, 2) & Mid(txt4.text, 4, 2) & Right(txt4.text, 2)
    End If
End If

End Sub

Public Sub Edit()

EditMode = True
Call EmptyBoxes
Call GetFirstDemand

cmdEdit.Visible = False
cmdGenerateManual.Visible = False
cmdGenAll.Visible = False
cmdDelete.Visible = False
cmdDeleteOld.Visible = False
cmdSaveNew.Visible = False
cmdCancelNew.Visible = False
cmdPrint.Visible = False
cmdPrintThis.Visible = False
cmdPrintBatch.Visible = False
cmdReprint.Visible = False
cmdCancelEdit.Visible = True
mnuEditDemands.Enabled = False
mnuGenManual.Enabled = False
mnuGenAll.Enabled = False
mnuPrint.Enabled = False
mnuReprint.Enabled = False
mnuPrintBatch.Enabled = False

Call EnableBoxes
txt8.Enabled = True

End Sub

Public Sub Delete()
Dim last As Boolean
Dim b As Integer
Dim c As Integer

If MsgBox("Are you sure you want to delete demand " & Text1.text & " ?", vbYesNo + vbQuestion, "Delete Demand") = vbNo Then Exit Sub

If MsgBox("Are you sure you want to delete demand " & Text1.text, vbYesNo + vbQuestion, "Delete Demand") = vbNo Then Exit Sub

MousePointer = vbHourglass

Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text1.text)
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

Rst1.Delete
'Rst1.Update
Rst1.Close
Conn.Close

MsgBox "Demand " & Text1.text & " deleted.", vbInformation + vbOKOnly, "Deleted"

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber > " & CInt(Text1.text)
rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

If rs.EOF = False Then
    rs.MoveLast
    rs.MoveFirst
    If rs.RecordCount = 1 Then last = True
    b = rs!UniqueRefNumber
    While rs.EOF = False
        If b > rs!UniqueRefNumber Then b = rs!UniqueRefNumber
        rs.MoveNext
    Wend
Else
    rs.Close
    strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE UniqueRefNumber < " & CInt(Text1.text)
    rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly
    If rs.EOF = False Then
        rs.MoveLast
        rs.MoveFirst
        If rs.RecordCount = 1 Then last = True
        b = rs!UniqueRefNumber
        While rs.EOF = False
            If b < rs!UniqueRefNumber Then b = rs!UniqueRefNumber
            rs.MoveNext
        Wend
    Else
        'no other demands
        rs.Close
        Call EmptyBoxes
        MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
        MousePointer = vbDefault
        Exit Sub
    End If
End If
rs.Close
strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & b
rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly

Call EmptyBoxes
If rs.EOF = False Then Call GetRecord

rs.Close
cnnDB.Close

Set rs = Nothing
Set cnnDB = Nothing

MousePointer = vbDefault

End Sub

Public Sub GenerateAll()

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

'Call DeleteDemands

Dim BRcount As Integer
Dim SCcount As Integer
Dim IPcount As Integer
Dim NextUniqueRefNo As Long
Dim a As Integer
Dim enddate As String

batch = 0
NextUniqueRefNo = 0
first = 0
If MsgBox("Generate Automatic Demands due in " & DaysB4Due & " days or less", vbYesNo + vbQuestion, _
    "Generate Automatic Demands") = vbNo Then Exit Sub

On Error GoTo ErrH

'so Msgbox = vbYe so generate demands
MousePointer = vbHourglass
'first work out the cut off date for demands to be generated for.
CutOffDate = DateAdd("d", DaysB4Due, Date)

'open connection
Set cnnDB = New ADODB.Connection
cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="

'Check all nominal codes entered for demand types
'Interest Payments
Set rs = New ADODB.Recordset
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '4'"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly

If IsNull(rs!Prefix) Then
    MsgBox "You have not entered a Sage Prefix for Interest Payments!" & Chr(13) & "Please enter a Sage Prefix for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Sage Prefix"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeforAmount) Then
    MsgBox "You have not selected Nominal Accounts for Interest Payments!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforAmount) Then
    MsgBox "You have not selected Nominal Accounts for Interest Payments!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeForTotal) Then
    MsgBox "You have not selected Nominal Accounts for Interest Payments!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforTotal) Then
    MsgBox "You have not selected Nominal Accounts for Interest Payments!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
rs.Close

'Base Rent
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '2'"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly
If IsNull(rs!Prefix) Then
    MsgBox "You have not entered a Sage Prefix for Base Rent Charges!" & Chr(13) & "Please enter a Sage Prefix for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Sage Prefix"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeforAmount) Then
    MsgBox "You have not selected Nominal Accounts for Base Rent Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforAmount) Then
    MsgBox "You have not selected Nominal Accounts for Base Rent Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeForVAT) Then
    MsgBox "You have not selected Nominal Accounts for Base Rent Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforVAT) Then
    MsgBox "You have not selected Nominal Accounts for Base Rent Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeForTotal) Then
    MsgBox "You have not selected Nominal Accounts for Base Rent Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforTotal) Then
    MsgBox "You have not selected Nominal Accounts for Base Rent Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
rs.Close

'Service Charge
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '1'"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly
If IsNull(rs!Prefix) Then
    MsgBox "You have not entered a Sage Prefix for Service Charges!" & Chr(13) & "Please enter a Sage Prefix for all Demand Types before generating demands.", vbOKOnly + vbCritical, "Missing Sage Prefix"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeforAmount) Then
    MsgBox "You have not selected Nominal Accounts for Service Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforAmount) Then
    MsgBox "You have not selected Nominal Accounts for Service Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeForVAT) Then
    MsgBox "You have not selected Nominal Accounts for Service Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforVAT) Then
    MsgBox "You have not selected Nominal Accounts for Service Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalCodeForTotal) Then
    MsgBox "You have not selected Nominal Accounts for Service Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
If IsNull(rs!NominalNameforTotal) Then
    MsgBox "You have not selected Nominal Accounts for Service Charges!" & Chr(13) & "Please select Nominal Accounts for all Demand Types before generating Demands.", vbOKOnly + vbCritical, "Missing Nominal Accounts"
    rs.Close
    cnnDB.Close
    MousePointer = vbDefault
    Exit Sub
End If
rs.Close

'Get the next batch number
SQLStr1 = "SELECT * FROM Batches"
Set rs = New ADODB.Recordset

rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly

If rs.EOF = False Then
    While rs.EOF = False
        If rs!batch > batch Then
            batch = rs!batch
            first = rs!last
        End If
        rs.MoveNext
    Wend
End If
batch = batch + 1
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

SQLStr1 = "SELECT UniqueRefNumber FROM DemandRecords"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly

If rs.EOF = False Then
    While rs.EOF = False
        If rs!UniqueRefNumber > first Then first = rs!UniqueRefNumber
        rs.MoveNext
    Wend
Else
    first = 1
End If

rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

'Connect to Demands table to add new demands.
Set rs2 = New ADODB.Recordset

SQLStr2 = "SELECT * FROM DemandRecords"
rs2.Open SQLStr2, cnnDB, adOpenDynamic, adLockPessimistic

'Interest Payment demands
'get nominal codes and prefix for base rent from demand types.
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '4'"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly

Prefix = rs!Prefix
NomCodeAmount = rs!NominalCodeforAmount
NomNameAmount = rs!NominalNameforAmount
NomCodeTotal = rs!NominalCodeForTotal
NomNameTotal = rs!NominalNameforTotal
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

'Get last unique ref number for batch
SQLStr1 = "SELECT UniqueRefNumber FROM DemandRecords"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly
If rs.EOF = False Then
    While rs.EOF = False
        If NextUniqueRefNo < rs!UniqueRefNumber Then NextUniqueRefNo = rs!UniqueRefNumber
        rs.MoveNext
    Wend
End If
NextUniqueRefNo = NextUniqueRefNo + 1
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

'Select those tenants who have an interest charge
SQLStr1 = "SELECT * FROM LeaseDetails WHERE InterestChargeable = 'Y'"
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = True And rs.BOF = True Then
    rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    GoTo BRDemands
End If
While rs.EOF = False
    description = "Interest Charges For " & rs!DaysAfterInterestPayable & " Days on " & rs!InterestChargedOn
    text = "S/L " & rs!SageAccountNumber
    'add new demand to demand table.
    rs2.AddNew
    rs2!batch = batch
    rs2!AutomaticManual = "A"
    rs2!SageAccountNumber = rs!SageAccountNumber
    rs2!TenantCompanyName = rs!CompanyName
    rs2!UnitNumber = rs!UnitNumber
    rs2!NominalCodeforAmount = NomCodeAmount
    rs2!NominalNameforAmount = NomNameAmount
    'rs2!NominalCodeForVAT = NomCodeVAT
    'rs2!NominalNameforVAT = NomNameVAT
    rs2!NominalCodeForTotal = NomCodeTotal
    rs2!NominalNameforTotal = NomNameTotal
    rs2!Source = 1
    rs2!TransactionType = 4
    rs2!DueDate = rs!BRNextDueDate
    rs2!typeofdemand = 3
    rs2!Amount = rs!InterestAmount
    rs2!VATAmount = 0
    rs2!VATMonth = Month(rs!IssueDate)
    rs2!Reference = Prefix & Left(rs!BRNextDueDate, 2) & Mid(rs!BRNextDueDate, 4, 2) & Right(rs!BRNextDueDate, 2)
    rs2!TotalAmount = rs2!Amount + rs2!VATAmount
    rs2!IssueDate = Date
    rs2!text = text
    rs2!description = description
    rs2!IsPrinted = "N"
    rs2!SendToPrint = ""
    rs2!ExportedToSage = "N"
    rs2!ExportedToExcel = "N"
    rs2.Update
    'rs.Edit
    rs!InterestChargeable = "N"
    rs!DaysAfterInterestPayable = Null
    rs!InterestChargedOn = Null
    rs!InterestAmount = Null
    rs.Update
    IPcount = IPcount + 1
    rs.MoveNext
Wend
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

'Base Rent Demands
'get nominal codes and prefix for base rent from demand types.
'MsgBox "Interest Charges Done"
BRDemands:
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '2'"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly

Prefix = rs!Prefix
NomCodeAmount = rs!NominalCodeforAmount
NomNameAmount = rs!NominalNameforAmount
NomCodeVAT = rs!NominalCodeForVAT
NomNameVAT = rs!NominalNameforVAT
NomCodeTotal = rs!NominalCodeForTotal
NomNameTotal = rs!NominalNameforTotal
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

SQLStr1 = "SELECT * FROM LeaseDetails WHERE BRPayable = 'Y'"
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = True And rs.BOF = True Then 'no records
    MsgBox "There are no Automatic Demands to Generate.  No Lease Details have been entered.", vbOKOnly + vbInformation, "Generate Automatic Demands"
    rs.Close
    Set rs = Nothing
    cnnDB.Close
    Set cnnDB = Nothing
    Call EmptyBoxes
    Call GetFirstDemand
    Exit Sub
End If
'Base Rent Demands
'Temp1 = beginning of payment period
'Temp2 = shortened version of temp1
'Temp3 = end of payment period
'Temp4 = shortened version of temp3
'Temp5 = next due date
'****same for Service charge demands*****
While rs.EOF = False
    If DateDiff("d", rs!BRNextDueDate, CutOffDate) < 0 Then
            rs.MoveNext
    Else
        enddate = Left(rs!enddate, 6) & "20" & Right(rs!enddate, 2)
        Select Case rs!BRfrequency
            Case 1: 'weekly in advance
                temp1 = rs!BRNextDueDate
                temp3 = DateAdd("d", 6, rs!BRNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("d", 7, rs!BRNextDueDate)
            Case 2: 'weekly in arrears
                temp1 = DateAdd("d", -7, rs!BRNextDueDate)
                temp3 = DateAdd("d", -1, rs!BRNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period" & temp1 & " to " & temp3
                temp5 = DateAdd("d", 7, rs!BRNextDueDate)
            Case 3: 'fortnightly in advance
                temp1 = rs!BRNextDueDate
                temp3 = DateAdd("d", 13, rs!BRNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("d", 14, rs!BRNextDueDate)
            Case 4: 'fortnightly in arrears
                temp1 = DateAdd("d", -14, rs!BRNextDueDate)
                temp3 = DateAdd("d", -1, rs!BRNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("d", 14, rs!BRNextDueDate)
            Case 5: 'monthly in advance
                temp1 = rs!BRNextDueDate
                temp3 = DateAdd("m", 1, rs!BRNextDueDate)
                temp3 = DateAdd("d", -1, temp3)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = enddate
                    temp3 = enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("m", 1, rs!BRNextDueDate)
            Case 6: 'monthly in arrears
                temp1 = DateAdd("m", -1, rs!BRNextDueDate)
                temp3 = DateAdd("d", -1, rs!BRNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                temp5 = DateAdd("m", 1, rs!BRNextDueDate)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
            Case 7: 'quarterly in advance
                temp1 = rs!BRNextDueDate
                Select Case rs!BRNextDueDate
                    Case quarterly1:
                        temp3 = DateAdd("d", -1, quarterly2)
                        temp5 = quarterly2 'nextduedate
                    Case quarterly2:
                        temp3 = DateAdd("d", -1, quarterly3)
                        temp5 = quarterly3 'nextduedate
                    Case quarterly3:
                        temp3 = DateAdd("d", -1, quarterly4)
                        temp5 = quarterly4 'nextduedate
                    Case quarterly4:
                        temp3 = DateAdd("yyyy", 1, quarterly1)
                        temp3 = DateAdd("d", -1, temp3)
                        temp5 = DateAdd("yyyy", 1, quarterly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
            Case 8: 'quarterly in arrears
                temp3 = DateAdd("d", -1, rs!BRNextDueDate)
                Select Case rs!BRNextDueDate
                    Case quarterly1:
                        temp1 = DateAdd("yyyy", -1, quarterly4)
                        temp5 = quarterly2 'nextduedate
                    Case quarterly2:
                        temp1 = quarterly1
                        temp5 = quarterly3 'nextduedate
                    Case quarterly3:
                        temp1 = quarterly2
                        temp5 = quarterly4 'nextduedate
                    Case quarterly4:
                        temp1 = quarterly3
                        temp5 = DateAdd("yyyy", 1, quarterly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
            Case 9: 'half yearly in advance
                temp1 = rs!BRNextDueDate
                Select Case rs!BRNextDueDate
                    Case halfyearly1:
                        temp3 = DateAdd("d", -1, halfyearly2)
                        temp5 = halfyearly2 'nextduedate
                    Case halfyearly2:
                        temp3 = DateAdd("d", -1, halfyearly1)
                        temp3 = DateAdd("yyyy", 1, temp3)
                        temp5 = DateAdd("yyyy", 1, halfyearly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
               ' Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
            Case 10: 'half yearly in arrears
                temp3 = DateAdd("d", -1, rs!BRNextDueDate)
                Select Case rs!BRNextDueDate
                    Case halfyearly1:
                        temp3 = DateAdd("yyyy", -1, halfyearly2)
                        temp5 = halfyearly2 'nextduedate
                    Case halfyearly2:
                        temp3 = halfyearly1
                        temp5 = DateAdd("yyyy", 1, halfyearly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
            Case 11: 'yearly in advance
                temp1 = rs!BRNextDueDate
                temp3 = DateAdd("yyyy", 1, rs!BRNextDueDate)
                temp3 = DateAdd("d", -1, temp3)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("yyyy", 1, rs!BRNextDueDate) 'nextduedate
            Case 12: 'yearly in arrears
                temp1 = rs!BRNextDueDate
                temp3 = DateAdd("yyyy", -1, temp1)
                temp3 = DateAdd("d", -1, temp3)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                temp5 = DateAdd("yyyy", 1, rs!BRNextDueDate)
               ' Text = "Base Rent " & temp2 & " to " & temp4
                description = "Base Rent For Period " & temp1 & " to " & temp3
        End Select
        'MsgBox rs!SageAccountNumber
        rs2.AddNew
        rs2!batch = batch
        rs2!AutomaticManual = "A"
        rs2!SageAccountNumber = rs!SageAccountNumber
        rs2!TenantCompanyName = rs!CompanyName
        rs2!UnitNumber = rs!UnitNumber
        rs2!NominalCodeforAmount = NomCodeAmount
        rs2!NominalNameforAmount = NomNameAmount
        rs2!NominalCodeForVAT = NomCodeVAT
        rs2!NominalNameforVAT = NomNameVAT
        rs2!NominalCodeForTotal = NomCodeTotal
        rs2!NominalNameforTotal = NomNameTotal
        rs2!Source = 1
        rs2!TransactionType = 4
        rs2!DueDate = rs!BRNextDueDate
        rs2!typeofdemand = 1
        rs2!Amount = rs!BRAmount
        rs2!VATAmount = Round(rs!BRAmount * VatRate / 100, 2)
        rs2!VATMonth = Month(rs!BRNextDueDate)
        rs2!Reference = Prefix & Left(rs!BRNextDueDate, 2) & Mid(rs!BRNextDueDate, 4, 2) & Right(rs!BRNextDueDate, 2)
        rs2!TotalAmount = rs2!Amount + rs2!VATAmount
        rs2!IssueDate = Date
        rs2!text = "S/L " & rs!SageAccountNumber
        rs2!description = description
        rs2!IsPrinted = "N"
        rs2!SendToPrint = ""
        rs2!ExportedToSage = "N"
        rs2!ExportedToExcel = "N"
        rs2.Update
        temp1 = ""
        temp2 = ""
        temp3 = ""
        temp4 = ""
        rs!BRNextDueDate = temp5
        rs.Update
        temp5 = ""
        BRcount = BRcount + 1
        rs.MoveNext
    End If
Wend
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

'MsgBox "Base Rent done"

'Service Charge demands
'get nominal codes and prefix for base rent from demand types.
SQLStr1 = "SELECT * FROM DemandTypes WHERE ID = '1'"
rs.Open SQLStr1, cnnDB, adOpenStatic, adLockReadOnly

Prefix = rs!Prefix
NomCodeAmount = rs!NominalCodeforAmount
NomNameAmount = rs!NominalNameforAmount
NomCodeVAT = rs!NominalCodeForVAT
NomNameVAT = rs!NominalNameforVAT
NomCodeTotal = rs!NominalCodeForTotal
NomNameTotal = rs!NominalNameforTotal
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

SQLStr1 = "SELECT * FROM LeaseDetails WHERE SCPayable = 'Y'"
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = True And rs.BOF = True Then 'no records
    MsgBox "There are no Automatic Demands to Generate.  No Lease Details have been entered.", vbOKOnly + vbInformation, "Generate Automatic Demands"
    rs.Close
    Set rs = Nothing
    cnnDB.Close
    Set cnnDB = Nothing
    Call EmptyBoxes
    Call GetFirstDemand
    Exit Sub
End If

While rs.EOF = False
    If DateDiff("d", rs!SCNextDueDate, CutOffDate) < 0 Then
        rs.MoveNext
    Else
        enddate = Left(rs!enddate, 6) & "20" & Right(rs!enddate, 2)
        Select Case rs!SCfrequency
            Case 1: 'weekly in advance
                temp1 = rs!SCNextDueDate
                temp3 = DateAdd("d", 6, rs!SCNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("d", 7, rs!SCNextDueDate)
            Case 2: 'weekly in arrears
                temp1 = DateAdd("d", -7, rs!SCNextDueDate)
                temp3 = DateAdd("d", -1, rs!SCNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period" & temp1 & " to " & temp3
                temp5 = DateAdd("d", 7, rs!SCNextDueDate)
            Case 3: 'fortnightly in advance
                temp1 = rs!SCNextDueDate
                temp3 = DateAdd("d", 13, rs!SCNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("d", 14, rs!SCNextDueDate)
            Case 4: 'fortnightly in arrears
                temp1 = DateAdd("d", -14, rs!SCNextDueDate)
                temp3 = DateAdd("d", -1, rs!SCNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("d", 14, rs!SCNextDueDate)
            Case 5: 'monthly in advance
                temp1 = rs!SCNextDueDate
                temp3 = DateAdd("m", 1, rs!SCNextDueDate)
                temp3 = DateAdd("d", -1, temp3)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = enddate
                    temp3 = enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
                temp5 = DateAdd("m", 1, rs!SCNextDueDate)
            Case 6: 'monthly in arrears
                temp1 = DateAdd("m", -1, rs!SCNextDueDate)
                temp3 = DateAdd("d", -1, rs!SCNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                temp5 = DateAdd("m", 1, rs!SCNextDueDate)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
            Case 7: 'quarterly in advance
                temp1 = rs!SCNextDueDate
                Select Case rs!SCNextDueDate
                    Case quarterly1:
                        temp3 = DateAdd("d", -1, quarterly2)
                        temp5 = quarterly2 'nextduedate
                    Case quarterly2:
                        temp3 = DateAdd("d", -1, quarterly3)
                        temp5 = quarterly3 'nextduedate
                    Case quarterly3:
                        temp3 = DateAdd("d", -1, quarterly4)
                        temp5 = quarterly4 'nextduedate
                    Case quarterly4:
                        temp3 = DateAdd("yyyy", 1, quarterly1)
                        temp3 = DateAdd("d", -1, temp3)
                        temp5 = DateAdd("yyyy", 1, quarterly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
            Case 8: 'quarterly in arrears
                temp3 = DateAdd("d", -1, rs!SCNextDueDate)
                Select Case rs!SCNextDueDate
                    Case quarterly1:
                        temp1 = DateAdd("yyyy", -1, quarterly4)
                        temp5 = quarterly2 'nextduedate
                    Case quarterly2:
                        temp1 = quarterly1
                        temp5 = quarterly3 'nextduedate
                    Case quarterly3:
                        temp1 = quarterly2
                        temp5 = quarterly4 'nextduedate
                    Case quarterly4:
                        temp1 = quarterly3
                        temp5 = DateAdd("yyyy", 1, quarterly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp3 = rs!enddate
                    temp1 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
            Case 9: 'half yearly in advance
                temp1 = rs!SCNextDueDate
                Select Case rs!SCNextDueDate
                    Case halfyearly1:
                        temp3 = DateAdd("d", -1, halfyearly2)
                        temp5 = halfyearly2 'nextduedate
                    Case halfyearly2:
                        temp3 = DateAdd("d", -1, halfyearly1)
                        temp3 = DateAdd("yyyy", 1, temp3)
                        temp5 = DateAdd("yyyy", 1, halfyearly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp2) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
            Case 10: 'half yearly in arrears
                temp3 = DateAdd("d", -1, rs!SCNextDueDate)
                Select Case rs!SCNextDueDate
                    Case halfyearly1:
                        temp3 = DateAdd("yyyy", -1, halfyearly2)
                        temp5 = halfyearly2 'nextduedate
                    Case halfyearly2:
                        temp3 = halfyearly1
                        temp5 = DateAdd("yyyy", 1, halfyearly1) 'nextduedate
                End Select
                If DateDiff("d", enddate, temp1) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
            Case 11: 'yearly in advance
                temp1 = rs!SCNextDueDate
                temp3 = DateAdd("yyyy", 1, temp1)
                temp3 = DateAdd("d", -1, temp3)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                temp5 = DateAdd("yyyy", 1, rs!SCNextDueDate) 'nextduedate
                'Text = "Ser Charge " & Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2) & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
            Case 12: 'yearly in arrears
                temp1 = DateAdd("yyyy", -1, rs!SCNextDueDate)
                temp3 = DateAdd("d", -1, rs!SCNextDueDate)
                If DateDiff("d", enddate, temp3) > 0 Then
                    temp1 = rs!enddate
                    temp3 = rs!enddate
                End If
                temp2 = Left(temp1, 2) & Mid(temp1, 4, 2) & Right(temp1, 2)
                temp4 = Left(temp3, 2) & Mid(temp3, 4, 2) & Right(temp3, 2)
                temp5 = DateAdd("yyyy", 1, rs!SCNextDueDate)
                'Text = "Ser Charge " & temp2 & " to " & temp4
                description = "Service Charge For Period " & temp1 & " to " & temp3
        End Select
        rs2.AddNew
        rs2!batch = batch
        rs2!AutomaticManual = "A"
        rs2!SageAccountNumber = rs!SageAccountNumber
        rs2!TenantCompanyName = rs!CompanyName
        rs2!UnitNumber = rs!UnitNumber
        rs2!NominalCodeforAmount = NomCodeAmount
        rs2!NominalNameforAmount = NomNameAmount
        rs2!NominalCodeForVAT = NomCodeVAT
        rs2!NominalNameforVAT = NomNameVAT
        rs2!NominalCodeForTotal = NomCodeTotal
        rs2!NominalNameforTotal = NomNameTotal
        rs2!Source = 1
        rs2!TransactionType = 4
        rs2!DueDate = rs!SCNextDueDate
        rs2!typeofdemand = 0
        rs2!Amount = rs!SCAmount
        rs2!VATAmount = Round(rs!SCAmount * VatRate / 100, 2)
        rs2!VATMonth = Month(rs!SCNextDueDate)
        rs2!Reference = Prefix & Left(rs!SCNextDueDate, 2) & Mid(rs!SCNextDueDate, 4, 2) & Right(rs!SCNextDueDate, 2)
        rs2!TotalAmount = rs2!Amount + rs2!VATAmount
        rs2!IssueDate = Date
        rs2!text = "S/L " & rs!SageAccountNumber
        rs2!description = description
        rs2!IsPrinted = "N"
        rs2!SendToPrint = ""
        rs2!ExportedToSage = "N"
        rs2!ExportedToExcel = "N"
        rs2.Update
        temp1 = ""
        temp2 = ""
        temp3 = ""
        temp4 = ""
        rs!SCNextDueDate = temp5
        rs.Update
        temp5 = ""
        SCcount = SCcount + 1
        rs.MoveNext
    End If
Wend
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

'MsgBox "Service Charge done"

rs2.Close
Set rs2 = Nothing
Set rs = New ADODB.Recordset

strSQLTitles = "SELECT * FROM DemandRecords"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        If a < rs!UniqueRefNumber Then a = rs!UniqueRefNumber
        rs.MoveNext
    Wend
    rs.MoveFirst
    Call EmptyBoxes
    Call GetRecord
End If
rs.Close
Set rs = Nothing
Set rs = New ADODB.Recordset

last = a

SQLStr1 = "SELECT * FROM Batches"
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

rs.AddNew
rs!batch = batch
rs!first = first
rs!last = last
rs.Update
rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

MousePointer = vbDefault

cmdMoveFirst.Visible = True
cmdMovePrevious.Visible = True
cmdMoveLast.Visible = True
cmdMoveNext.Visible = True

Dim Msg As String

Msg = "Batch " & batch & " has been generated." & Chr(13)
Msg = Msg & BRcount & " Demands for Base Rent were generated." & Chr(13)
Msg = Msg & SCcount & " Demands for Service Charge were generated." & Chr(13)
Msg = Msg & IPcount & " Demands for Interest Payment were generated." & Chr(13)
Msg = Msg & "A total of " & BRcount + SCcount + IPcount & " demands were generated."

MsgBox Msg, vbOKOnly + vbInformation, "Demands Generated"

If MsgBox("Do you want to run Demand Status report now? To run now click OK.  To run later click Cancel.", vbOKCancel + vbQuestion, "Run Report") = vbOK Then
    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset

    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
    SQLStr1 = "SELECT ExportedToSage FROM DemandRecords WHERE IsPrinted = 'Y' AND ExportedToSage = 'N'OR ExportedToSage = 'P'"
    rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic
    
    If rs.EOF = False Then
        While rs.EOF = False
            rs!ExportedToSage = "P"
            rs.Update
            rs.MoveNext
        Wend
    End If
    rs.Close
    Set rs = Nothing
    Set rs = New ADODB.Recordset
    CR1.ReportTitle = "Demand Status"
    CR1.ReportFileName = App.Path & "\PreUpdate" & SCID & ".rpt"
    CR1.printReport

    SQLStr1 = "SELECT ExportedToSage FROM DemandRecords WHERE ExportedToSage = 'P'"
    rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic
    
    If rs.EOF = False Then
        While rs.EOF = False
            rs!ExportedToSage = "N"
            rs.Update
            rs.MoveNext
        Wend
    End If
    rs.Close
    cnnDB.Close
    Set rs = Nothing
    Set cnnDB = Nothing
End If

Exit Sub

ErrH:
    'This can only pick up error 13 (type mis-match) and it is at the users discretion to not enter a date.
    'MsgBox Err.Number & " - " & Err.Description, vbOKOnly, "Error"
    Resume Next
End Sub

Public Sub PrintDemands()

PrintMode = True
Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE IsPrinted = 'N'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        rs!SendToPrint = "Y"
        rs.Update
        rs.MoveNext
    Wend
Else
    MsgBox "There are no demands to print.", vbOKOnly + vbInformation, "Print Demands"
    rs.Close
    cnnDB.Close
    Set rs = Nothing
    Set cnnDB = Nothing
    PrintMode = False
    Exit Sub
End If

rs.Close
cnnDB.Close
Set cnnDB = Nothing
Set rs = Nothing

Call EmptyBoxes
Call GetFirstDemand

'user selected Print Demands
lbl1.Visible = True
chkPrint.Visible = True
chkPrint.Value = 1
cmdGenAll.Visible = False
cmdGenerateManual.Visible = False
cmdEdit.Visible = False
cmdDelete.Visible = False
cmdDeleteOld.Visible = False
cmdReprint.Visible = False
cmdPrint.Visible = False
cmdPrintThis.Visible = False
cmdPrintAll.Visible = True
cmdPrintBatch.Visible = False
cmdPrintSome.Visible = True
cmdCancelPrint.Visible = True
Call DisableMenu

End Sub

Public Sub DisableMenu()

mnuGenManual.Enabled = False
mnuDeleteDemand.Enabled = False
mnuEditDemands.Enabled = False
mnuPrint.Enabled = False
mnuReprint.Enabled = False
mnuPrintBatch.Enabled = False

End Sub

Public Sub EnableMenu()

mnuGenManual.Enabled = True
mnuDeleteDemand.Enabled = True
mnuEditDemands.Enabled = True
mnuPrint.Enabled = True
mnuReprint.Enabled = True
mnuPrintBatch.Enabled = True

End Sub

Public Sub ReprintDemands()

ReprintMode = True

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE IsPrinted = 'Y'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        rs!SendToPrint = "Y"
        rs.Update
        rs.MoveNext
    Wend
Else
    MsgBox "There are no demands to reprint.", vbOKOnly + vbInformation, "Reprint"
    rs.Close
    cnnDB.Close
    Set rs = Nothing
    Set cnnDB = Nothing
    ReprintMode = False
    Exit Sub
End If

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

Call EmptyBoxes
Call GetFirstDemand

chkPrint.Visible = True
cmdGenAll.Visible = False
cmdGenerateManual.Visible = False
cmdEdit.Visible = False
cmdDelete.Visible = False
cmdDeleteOld.Visible = False
cmdReprint.Visible = False
cmdPrint.Visible = False
cmdPrintThis.Visible = False
cmdReprintAll.Visible = True
cmdPrintBatch.Visible = False
cmdReprintSome.Visible = True
cmdCancelReprint.Visible = True
lbl1.Visible = True

Call DisableMenu

End Sub

Public Sub PrintBatch(BatchToPrint As Integer)

'Calls the end timeout
'Call CheckDateAndTimeoutFileNoKey

Conn.Connect = "DSN=" & Adsn & ";UID=;PWD="
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT * FROM Batches WHERE Batch = " & BatchToPrint
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

If Rst1.EOF = True And Rst1.EOF = True Then ' no record
    MsgBox "Invalid batch number entered.", "Invalid Batch"
    Exit Sub
End If

If MsgBox("Print Batch " & BatchToPrint & ": Demands " & Rst1!first & " to " & Rst1!last & " ?", vbYesNo, "Print Batch") = vbNo Then Exit Sub

Rst1.Close

SQLStr1 = "SELECT UniqueRefNumber, IsPrinted FROM DemandRecords WHERE Batch = " & BatchToPrint
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenDynamic, rdConcurRowVer)

If Rst1.EOF = False Then
    While Rst1.EOF = False
        Rst1.Edit
        'MsgBox Rst1!UniqueRefNumber
        Rst1!IsPrinted = "C"
        Rst1.Update
        Rst1.MoveNext
    Wend
End If
Rst1.Close

'send demands to crystal.
CR1.ReportFileName = App.Path & "\Demand" & SCID & ".rpt"
MsgBox CR1.ReportFileName
CR1.printReport

Call SetPrintedtoYes

End Sub


'Public Sub GetUniqueReferenceNumber()

'Dim Source As String
'Dim a As Integer
'Set cnnDB = New ADODB.Connection
'Set rs = New ADODB.Recordset

'cnnDB.Open "DSN=" & Adsn & ";UID=;PWD=;"

'Source = "SELECT UniqueRefNumber FROM DemandRecords"
'rs.Open Source, cnnDB, adOpenStatic, adLockReadOnly

'If rs.EOF = False Then
    'While rs.EOF = False
     '   If a < rs!UniqueRefNumber Then a = rs!UniqueRefNumber
    '    rs.MoveNext
   ' Wend
  '  rs.MoveFirst
 '   Call GetRecord
'End If

'NextUniqueRefNo = a + 1
'rs.Close
'cnnDB.Close
'Set cnnDB = Nothing

'End Sub

Public Sub GetFirstDemand()

    Dim a As Integer

    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset

    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD=;"

    strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords"
    If EditMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE ExportedToSage = 'N'"
    If PrintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'N'"
    If ReprintMode = True Then strSQLTitles = "SELECT UniqueRefNumber FROM DemandRecords WHERE IsPrinted = 'Y'"

    rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly
    
    If rs.EOF = False Then
        a = rs!UniqueRefNumber
        
        While rs.EOF = False
            If a > rs!UniqueRefNumber Then
                a = rs!UniqueRefNumber
            End If

            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        Set rs = New ADODB.Recordset
        strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & a
        rs.Open strSQLTitles, cnnDB, adOpenStatic, adLockReadOnly
        Call GetRecord
        cmdMoveFirst.Visible = True
        cmdMoveNext.Visible = True
        cmdMoveLast.Visible = True
        cmdMovePrevious.Visible = True
        rs.Close
        cnnDB.Close
        Set rs = Nothing
        Set cnnDB = Nothing
    Else
        MsgBox "There are no demands!", vbOKOnly + vbInformation, "No Demands"
        Exit Sub
    End If

End Sub

Public Sub UpdatePrint()

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text1.text)
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If chkPrint.Value = 0 Then ' user has selected to not print the current demand
    rs!SendToPrint = ""
    rs.Update
Else
    rs!SendToPrint = "Y"
    rs.Update
End If

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

Public Sub UnSelect()

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
strSQLTitles = "SELECT SendToPrint FROM DemandRecords WHERE SendToPrint = 'Y'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        rs!SendToPrint = ""
        rs.Update
        rs.MoveNext
    Wend
End If
rs.Close
cnnDB.Close
Set cnnDB = Nothing

End Sub

Public Sub SetPrintedtoYes()

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="

strSQLTitles = "SELECT IsPrinted FROM DemandRecords WHERE IsPrinted = 'C'"
rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        rs!IsPrinted = "Y"
        rs.Update
        rs.MoveNext
        DoEvents
    Wend
End If

rs.Close
cnnDB.Close
Set cnnDB = Nothing

End Sub

Public Sub PrintBatchSelected()

'to go to timeout
'Call CheckDateAndTimeoutFileNoKey

Dim Response
Dim batchnum As String
Dim match As Boolean
Dim NumOfBatches As Integer
Dim i As Integer
match = False

Conn.Connect = "DSN=" & Adsn & ";UID=PWD=;"
Conn.CursorDriver = rdUseIfNeeded
Conn.EstablishConnection rdDriverNoPrompt

SQLStr1 = "SELECT Batch FROM batches"
Set Rst1 = Conn.OpenResultset(SQLStr1, rdOpenStatic, rdConcurReadOnly)

If Rst1.EOF = False Then
    Rst1.MoveLast
    Rst1.MoveFirst
    NumOfBatches = Rst1.RowCount
    ReDim GetBatches(NumOfBatches) As Integer
    i = 1
    While Rst1.EOF = False
        GetBatches(i) = Rst1!batch
        Rst1.MoveNext
        i = i + 1
    Wend
End If
Rst1.Close
Conn.Close

ReenterBatch:
batchnum = InputBox("Enter batch to print: ", "Print Batch")

If IsNumeric(batchnum) = False Then
    While IsNumeric(batchnum) = False
        Response = MsgBox("You have entered an invalid batch number.", vbRetryCancel, "Incorrect Batch")
        If Response = vbCancel Then Exit Sub
        If Response = vbRetry Then batchnum = InputBox("Enter a batch to print: ", "Print Batch")
    Wend
End If

For i = 1 To NumOfBatches
    If GetBatches(i) = CInt(batchnum) Then match = True
Next i
If match = False Then 'not valid batchnumber
    Response = MsgBox("You have entered an invalid batch number.", vbRetryCancel, "Invalid Batch")
    If Response = vbCancel Then Exit Sub
    If Response = vbRetry Then GoTo ReenterBatch
Else 'valid batch number
    Call UnSelect
    Call PrintBatch(CInt(batchnum))
End If

End Sub

Public Sub SaveChanges()

If MsgBox("Save changes to current demand?", vbSystemModal + vbQuestion + vbYesNo, "Save Changes") = vbYes Then
    Set cnnDB = New ADODB.Connection
    Set rs = New ADODB.Recordset
    cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
    strSQLTitles = "SELECT * FROM DemandRecords WHERE UniqueRefNumber = " & CInt(Text1.text)
    rs.Open strSQLTitles, cnnDB, adOpenDynamic, adLockPessimistic
    If a = "I" Then rs!TransactionType = 4
    If a = "C" Then rs!TransactionType = 5
    rs!TotalAmount = txt9.text
    rs!VATAmount = txt8.text
    rs!Amount = txt7.text
    If txt4.text <> "" Then rs!IssueDate = Left(txt4.text, 6) + Right(txt4.text, 2)
    rs!DueDate = Left(txt5.text, 6) + Right(txt5.text, 2)
    rs!VATMonth = Month(txt4.text)
    If cboType.ListIndex = -1 Then
        rs!typeofdemand = typeofdemand
    Else
        rs!typeofdemand = cboType.ListIndex
    End If
    rs!Reference = txt6.text
    rs!text = txt11.text
    rs!description = txt10.text
    rs.Update
    rs.Close
    cnnDB.Close
    Set cnnDB = Nothing
End If

End Sub


Public Sub CancelPrint(a As Integer)

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
SQLStr1 = "SELECT IsPrinted FROM DemandRecords WHERE batch = " & a
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

While rs.EOF = False
    rs!IsPrinted = "N"
    rs.Update
    rs.MoveNext
Wend

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

Public Sub DeleteDemands()

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
SQLStr1 = "SELECT * FROM DemandRecords WHERE ExportedToSage = 'Y' AND ExportedToExcel = 'Y'"
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
End If

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

Public Sub CloseResultsets()

rs.Close
rs2.Close
Set rs = Nothing
Set rs2 = Nothing
cnnDB.Close
Set cnnDB = Nothing

End Sub

Private Sub cmdClearDemands_Click()
'user wants to delete all old demands - ones that have been printed, exported to sage and exported to excel
If MsgBox("Do you really want to clear down the demands table. All demands will be permanently erased?", vbYesNo + vbQuestion, "Delete Old Demands") = vbNo Then Exit Sub
MousePointer = vbHourglass

Call ClearDemands
'MsgBox "Old demands deleted successfully", vbOKOnly + vbInformation, "Deleted"
Call EmptyBoxes
Call GetFirstDemand
MsgBox "Old demands deleted successfully", vbOKOnly + vbInformation, "Deleted"
MousePointer = vbDefault

End Sub



Public Sub ClearDemands()

Set cnnDB = New ADODB.Connection
Set rs = New ADODB.Recordset

cnnDB.Open "DSN=" & Adsn & ";UID=;PWD="
SQLStr1 = "SELECT * FROM DemandRecords" 'Delete All this time. 'WHERE ExportedToSage = 'Y' AND ExportedToExcel = 'Y'"
rs.Open SQLStr1, cnnDB, adOpenDynamic, adLockPessimistic

If rs.EOF = False Then
    While rs.EOF = False
        rs.Delete
        rs.MoveNext
    Wend
End If

rs.Close
cnnDB.Close
Set rs = Nothing
Set cnnDB = Nothing

End Sub

