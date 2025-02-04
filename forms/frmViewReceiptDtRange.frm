VERSION 5.00
Begin VB.Form frmViewReceiptDtRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Receipt Details"
   ClientHeight    =   2925
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
   Icon            =   "frmViewReceiptDtRange.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2925
   ScaleWidth      =   3795
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtReference 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2160
      MaxLength       =   8
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton GenerateReportButton 
      Caption         =   "Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Specify Date Range"
      Height          =   1455
      Left            =   240
      TabIndex        =   4
      Top             =   720
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
         TabIndex        =   2
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
         TabIndex        =   1
         Text            =   "01/01/1980"
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Enter the Reference"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "frmViewReceiptDtRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Me.BackColor = MODULEBACKCOLOR
   Frame2.BackColor = Me.BackColor

   txtToDate.text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub GenerateReportButton_Click()
   Dim datetype As Integer

   If txtFromDate.text = "" Then
      MsgBox "Please enter the from date?", vbExclamation + vbOKOnly, "Enter Date"
      txtFromDate.SetFocus
      Exit Sub
   End If

   If txtToDate.text = "" Then
      MsgBox "Please enter the to date?", vbExclamation + vbOKOnly, "Enter Date"
      txtToDate.SetFocus
      Exit Sub
   End If

   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PreViewReceiptExportedSAGE.rpt")

   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

   Report.EnableParameterPrompting = False
   Report.DiscardSavedData

   ' Passing the from and to date values to Crystal Reports
   Report.ParameterFields(1).AddCurrentValue CStr(txtReference.text)
   Report.ParameterFields(2).AddCurrentValue CDate(txtFromDate.text)
   Report.ParameterFields(3).AddCurrentValue CDate(txtToDate.text)
   Report.ParameterFields(4).AddCurrentValue IIf(txtReference.text = "", False, True)

   Load frmReport
   frmReport.LoadReportViewer Report
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

