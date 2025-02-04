VERSION 5.00
Begin VB.Form frmDtRange4BRec 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date Range for Bank Recon..."
   ClientHeight    =   2505
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3795
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDtRange4BRec.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtStatDate 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000014&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      BeginProperty Font 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   1680
      MaxLength       =   10
      TabIndex        =   2
      Text            =   "01/01/1980"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   7
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerateReportButton 
      Caption         =   "&Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Transaction Date - "
      Height          =   1095
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   3375
      Begin VB.TextBox txtToDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   1
         Top             =   660
         Width           =   1575
      End
      Begin VB.TextBox txtFromDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000014&
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   2057
            SubFormatType   =   3
         EndProperty
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   285
         Left            =   1440
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "01/01/1980"
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   660
         Width           =   210
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   300
         Width           =   390
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Statement Date:"
      Height          =   195
      Left            =   360
      TabIndex        =   8
      Top             =   1440
      Width           =   1140
   End
End
Attribute VB_Name = "frmDtRange4BRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
   Unload Me
End Sub

Private Sub cmdGenerateReportButton_Click()
   If txtFromDate.text = "" Then
      MsgBox "Please enter the FROM date?", vbExclamation + vbOKOnly, "Enter Date"
      txtFromDate.SetFocus
      Exit Sub
   End If

   If txtToDate.text = "" Then
      MsgBox "Please enter the TO date?", vbExclamation + vbOKOnly, "Enter Date"
      txtToDate.SetFocus
      Exit Sub
   End If

   If txtStatDate.text = "" Then
      MsgBox "Please enter the STATEMENT date?", vbExclamation + vbOKOnly, "Enter Date"
      txtStatDate.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CBBankRecon.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag
'   Bank Account Balance
   Report.ParameterFields(2).AddCurrentValue _
         frmCashbook.CalDrCrAcBalance(CDate(Format(txtFromDate.text, "dd mmmm yyyy")), CDate(Format(txtToDate.text, "dd mmmm yyyy")))
         'Val(frmCashbook.txtAcBal.text)
   Report.ParameterFields(3).AddCurrentValue CDate(txtFromDate.text)
   Report.ParameterFields(4).AddCurrentValue CDate(txtToDate.text)
   Report.ParameterFields(5).AddCurrentValue CDate(txtStatDate.text)
   Report.ParameterFields(6).AddCurrentValue CDate(frmCashbook.txtStatementDate.text)
   Report.ParameterFields(7).AddCurrentValue Format(frmCashbook.txtStOpenBal.text, "0.00")
'   Statement Closing Balance
'Below line modified by anol 18 Feb 2015
'Below is statement closing balance
'Issue 530
   Report.ParameterFields(8).AddCurrentValue Val(frmCashbook.txtStOpenBal.text) ' frmCashbook.CalculatedClosingBal(CDate(Format(txtToDate.text, "dd mmmm yyyy")))
   ' Val(frmCashbook.txtStOpenBal.text)
'         frmCashbook.CalculatedClosingBal (CDate(Format(txtToDate.text, "dd mmmm yyyy")))
'   Uncleared Balance
   Report.ParameterFields(9).AddCurrentValue _
         frmCashbook.CalculatedUnClearedBal(CDate(Format(txtFromDate.text, _
                                             "dd mmmm yyyy")), CDate(Format(txtToDate.text, _
                                             "dd mmmm yyyy")), CDate(Format(txtStatDate.text, _
                                             "dd mmmm yyyy")))

   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
   
   Unload Me
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   txtToDate.text = Format(Now, "dd/mm/yyyy")
   txtStatDate.text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
   frmCashbook.Enabled = True

   Unload Me
End Sub

Private Sub txtFromDate_Change()
   TextBoxChangeDate txtFromDate
End Sub

Private Sub txtFromDate_GotFocus()
   SelTxtInCtrl txtFromDate
End Sub

Private Sub txtFromDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtFromDate, KeyAscii
End Sub

Private Sub txtFromDate_LostFocus()
   TextBoxFormatDate txtFromDate
End Sub

Private Sub txtStatDate_Change()
   TextBoxChangeDate txtStatDate
End Sub

Private Sub txtStatDate_GotFocus()
   SelTxtInCtrl txtStatDate
End Sub

Private Sub txtStatDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtStatDate, KeyAscii
End Sub

Private Sub txtStatDate_LostFocus()
   TextBoxFormatDate txtStatDate
End Sub

Private Sub txtToDate_Change()
   TextBoxChangeDate txtToDate
End Sub

Private Sub txtToDate_GotFocus()
   SelTxtInCtrl txtToDate
End Sub

Private Sub txtToDate_KeyPress(KeyAscii As Integer)
   TextBoxKeyPrsDate txtToDate, KeyAscii
End Sub

Private Sub txtToDate_LostFocus()
   TextBoxFormatDate txtToDate
End Sub

