VERSION 5.00
Begin VB.Form frmDtRange4CB 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Date Range for Cashbook"
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
   Icon            =   "frmDtRange4CB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Close"
      Height          =   375
      Left            =   2040
      TabIndex        =   6
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerateReportButton 
      Caption         =   "&Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Statement Date Range"
      Height          =   1455
      Left            =   240
      TabIndex        =   3
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   1
         Top             =   960
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
         Left            =   1200
         MaxLength       =   10
         TabIndex        =   0
         Text            =   "01/01/2000"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmDtRange4CB"
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
      MsgBox "Please enter the From date?", vbExclamation + vbOKOnly, "Enter Date"
      txtFromDate.SetFocus
      Exit Sub
   End If

   If txtToDate.text = "" Then
      MsgBox "Please enter the To date?", vbExclamation + vbOKOnly, "Enter Date"
      txtToDate.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim rep As frmReport

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\CashBookRptPay.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   Report.ParameterFields(1).AddCurrentValue frmCashbook.txtBC.Tag
   Report.ParameterFields(2).AddCurrentValue CDate(Format(txtFromDate.text, "dd mmmm yyyy"))
   Report.ParameterFields(3).AddCurrentValue CDate(Format(txtToDate.text, "dd mmmm yyyy"))
   Report.ParameterFields(4).AddCurrentValue frmCashbook.CalDrCrAcBalance(CDate(Format(txtFromDate.text, "dd mmmm yyyy")), CDate(Format(txtToDate.text, "dd mmmm yyyy")))
   'resolved by BOSL
   'issue 523 cashbook report was not working properly
   'added by anol 19 Jan 2015
   Report.ParameterFields(5).AddCurrentValue frmCashbook.txtClientList.Tag
   Set rep = New frmReport
   Load rep
   rep.LoadReportViewer Report
   
   Unload Me
End Sub

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = MODULEBACKCOLOR
   txtToDate.text = Format(Now, "dd/mm/yyyy")
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

