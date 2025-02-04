VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmSCYRR 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Yearend Income Reconciliation Report"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSCYRR.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4335
   Begin VB.CommandButton cmdSCYRRClose 
      Caption         =   "&Close"
      Height          =   480
      Left            =   2880
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3855
      Width           =   1335
   End
   Begin VB.CommandButton cmdSCYRRPrint 
      Caption         =   "&Print"
      Default         =   -1  'True
      Height          =   480
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3855
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "Schedule of Yearly Income"
      Height          =   2175
      Left            =   120
      TabIndex        =   10
      Top             =   1560
      Width           =   4095
      Begin VB.OptionButton optSCYRRSngFnd 
         Caption         =   "Income by Fund - Single"
         Height          =   375
         Left            =   720
         TabIndex        =   11
         Top             =   1080
         Width           =   2175
      End
      Begin VB.OptionButton optSCYRRAllFnd 
         Caption         =   "Income by Fund - All"
         Height          =   375
         Left            =   720
         TabIndex        =   3
         Top             =   720
         Width           =   2175
      End
      Begin VB.OptionButton optSCYRRTotal 
         Caption         =   "Summary Report for all income"
         Height          =   375
         Left            =   720
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   2895
      End
      Begin MSForms.ComboBox cboSCYRRFnd 
         Height          =   315
         Left            =   720
         TabIndex        =   4
         Top             =   1440
         Width           =   2850
         VariousPropertyBits=   1753237531
         DisplayStyle    =   3
         Size            =   "5027;556"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   6
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "881"
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   2
         Height          =   1260
         Index           =   0
         Left            =   600
         Top             =   645
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   1260
         Index           =   1
         Left            =   600
         Top             =   645
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Year"
      Height          =   1215
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4095
      Begin VB.TextBox txtSCYRREnDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox txtSCYRRStDt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End Date"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Start Date"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmSCYRR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboSCYRRFnd_Click()
   optSCYRRSngFnd.Value = True
End Sub

Private Sub cmdSCYRRClose_Click()
   Unload Me
End Sub

Private Sub cmdSCYRRPrint_Click()
   If txtSCYRRStDt.text = "" Then
      txtSCYRRStDt.SetFocus
      Exit Sub
   End If
   If txtSCYRREnDt.text = "" Then
      txtSCYRREnDt.SetFocus
      Exit Sub
   End If
   If optSCYRRSngFnd.Value And cboSCYRRFnd.text = "" Then
      ShowMsgInTaskBar "Please select a fund."
      cboSCYRRFnd.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

'  All option selected
   If optSCYRRTotal.Value Then _
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\YrEndRecon.rpt")

   If optSCYRRAllFnd.Value Then _
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\YrEndIncRec.rpt")

   If optSCYRRSngFnd.Value Then _
      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\YrEndIncRec.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue CDate(txtSCYRRStDt.text)
   Report.ParameterFields(2).AddCurrentValue CDate(txtSCYRREnDt.text)

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub
'
'Private Function CalculateOpeningBalance(adoConn As ADODB.Connection) As Currency
'   Dim adoRst As New ADODB.Recordset
'   Dim szSQL As String
'
'   szSQL = "SELECT SUM(OSAmount)"
'
'End Function

Private Sub Form_Activate()
   If Not LoadCboSCYRRFnd Then Unload Me
End Sub

Private Sub Form_Load()
   Me.Top = 0
   Me.Left = 0
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor
   Frame2.BackColor = Me.BackColor
   optSCYRRTotal.BackColor = Me.BackColor
   optSCYRRAllFnd.BackColor = Me.BackColor
   optSCYRRSngFnd.BackColor = Me.BackColor

'  Set a set for Start and end date text boxes

   txtSCYRRStDt.text = "01/01/" & Format(Now, "yyyy")
   txtSCYRREnDt.text = Format(Now, "dd/mm/yyyy")
End Sub

Private Function LoadCboSCYRRFnd() As Boolean
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   LoadCboSCYRRFnd = True

   adoConn.Open getConnectionString
   szSQL = "SELECT FundID, FundName FROM FUND;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then GoTo NoRes

   Dim TotalRow As Integer, TotalCol As Integer
   Dim i As Integer, j As Integer

   TotalRow = adoRst.RecordCount
   TotalCol = adoRst.Fields.count - 1

   Dim Data() As String

   ReDim Data(TotalCol, TotalRow) As String

   For i = 0 To TotalRow - 1
       For j = 0 To TotalCol
           Data(j, i) = IIf(IsNull(adoRst.Fields(j).Value), "", adoRst.Fields(j).Value)
       Next j
       adoRst.MoveNext
       If adoRst.EOF Then Exit For
   Next i
   cboSCYRRFnd.Column() = Data()
   cboSCYRRFnd.ListIndex = -1
   
   adoRst.Close
   adoConn.Close
   Set adoConn = Nothing
   
   Exit Function

NoRes:
   Set adoRst = Nothing
   adoConn.Close
   Set adoConn = Nothing

   ShowMsgInTaskBar "You have not set up any fund.", , "N"

   LoadCboSCYRRFnd = False
   
   Exit Function

ErrorHandler:
   ShowMsgInTaskBar ERR.description & "::" & ERR.Number, , "N"

   Set adoRst = Nothing

   adoConn.Close
   Set adoConn = Nothing
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmCashbook.MousePointer = vbArrow
End Sub

Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmCashbook.MousePointer = vbArrow
End Sub

Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   frmCashbook.MousePointer = vbArrow
End Sub

Private Sub optSCYRRAllFnd_Click()
   cboSCYRRFnd.ListIndex = -1
End Sub

Private Sub optSCYRRSngFnd_Click()
   cboSCYRRFnd.SetFocus
End Sub

Private Sub optSCYRRTotal_Click()
   cboSCYRRFnd.ListIndex = -1
End Sub

Private Sub txtSCYRREnDt_Change()
   TextBoxChangeDate txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_GotFocus()
   SelTxtInCtrl txtSCYRREnDt
End Sub

Private Sub txtSCYRREnDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSCYRREnDt, KeyAscii
End Sub

Private Sub txtSCYRREnDt_LostFocus()
   TextBoxFormatDate txtSCYRREnDt
End Sub

Private Sub txtSCYRRStDt_Change()
   TextBoxChangeDate txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_GotFocus()
   SelTxtInCtrl txtSCYRRStDt
End Sub

Private Sub txtSCYRRStDt_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtSCYRRStDt, KeyAscii
End Sub

Private Sub txtSCYRRStDt_LostFocus()
   TextBoxFormatDate txtSCYRRStDt
End Sub
