VERSION 5.00
Begin VB.Form frmPreSCBudVsAE 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Service Charge Budget Vs Actual Expenditure"
   ClientHeight    =   6030
   ClientLeft      =   1125
   ClientTop       =   1935
   ClientWidth     =   8925
   ClipControls    =   0   'False
   DrawMode        =   2  'Blackness
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreSCBudVsAE.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8925
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Year"
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Index           =   1
         Left            =   840
         TabIndex        =   6
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Index           =   0
         Left            =   840
         TabIndex        =   5
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   360
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1500
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   360
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1500
      Width           =   1335
   End
End
Attribute VB_Name = "frmPreSCBudVsAE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SEARCHPropertyMODE_ As Boolean
Dim LOOKUPCommand As String

Private Sub cmdRefreshData_Click()
   Me.MousePointer = vbHourglass
   
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString

   Call ExportData2NominalLedger(adoConn)

   adoConn.Close
   Set adoConn = Nothing
   Me.MousePointer = vbArrow
End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
   Dim adoConn    As New ADODB.Connection

   adoConn.Open getConnectionString

   If Not AreCA_Setup(adoConn) Then
      ShowMsgInTaskBar "Please setup control accounts for the client(s)", "Y", "N"
      adoConn.Close
      Set adoConn = Nothing
      Exit Sub
   End If

   adoConn.Close
   Set adoConn = Nothing

   If txtSCYRRStDt.text = "" Then
      txtSCYRRStDt.SetFocus
      Exit Sub
   End If
   If txtSCYRREnDt.text = "" Then
      txtSCYRREnDt.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\SCBudVsAE.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy"))
   Report.ParameterFields(2).AddCurrentValue CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy"))

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub Form_Load()
   Me.Height = 2460
   Me.Width = 5595
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = MODULEBACKCOLOR

   txtSCYRRStDt.text = "01/01/2000"
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")

   Dim szSQL      As String
   Dim adoConn    As New ADODB.Connection
   Dim TotalRow   As Integer, TotalCol As Integer
   Dim i          As Integer, j        As Integer

   adoConn.Open getConnectionString

   Call BankAndBalance(adoConn)

   adoConn.Close
   Set adoConn = Nothing
End Sub

Private Sub txtSCYRREnDt_Change()
    TextBoxChangeDate txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_GotFocus()
   If Len(txtSCYRREnDt.text) < 10 Then txtSCYRREnDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
    TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
End Sub

Private Sub txtSCYRREnDt_LostFocus()
    If txtSCYRREnDt.text <> "" Then TextBoxFormatDate txtSCYRREnDt
End Sub

Private Sub txtSCYRRStDt_Change()
    TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_GotFocus()
   If Len(txtSCYRRStDt.text) < 10 Then txtSCYRRStDt.text = Format(Date, "dd/mm/yyyy")
   SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
    TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
End Sub

Private Sub txtSCYRRStDt_LostFocus()
    If txtSCYRRStDt.text <> "" Then TextBoxFormatDate txtSCYRRStDt
End Sub

Private Function BankAndBalance(adoConn As ADODB.Connection) As String
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   szSQL = "SELECT CB.NominalCode AS BNC, CB.MY_ID AS ID, " & _
               "N.Name AS BNN, CB.CurrentBalance AS BAL, CB.CLIENT_ID " & _
           "FROM tlbClientBanks AS CB, NominalLedger AS N " & _
           "WHERE CB.NominalCode = N.Code AND CB.CLIENT_ID <> '' " & _
           "GROUP BY CB.NominalCode, CB.MY_ID, N.Name, CB.CurrentBalance, CB.CLIENT_ID;"
'Debug.Print szSQL
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      MsgBox "Please setup bank account for the client."
   Else
   End If

   ' Destroy Objects
   Set adoRst = Nothing

'   LoadAdoBank

   Exit Function

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Function
