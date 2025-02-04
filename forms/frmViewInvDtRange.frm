VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmViewInvDtRange 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Demand Analysis Report"
   ClientHeight    =   6675
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3885
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewInvDtRange.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   3885
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton GenerateReportButton 
      Caption         =   "Generate Report"
      Default         =   -1  'True
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   3480
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Date Type"
      Height          =   1455
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3375
      Begin MSForms.OptionButton optIssueDate 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   840
         Width           =   1935
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "3413;661"
         Value           =   "0"
         Caption         =   "Issue Date"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.OptionButton optDueDate 
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   360
         Width           =   1815
         VariousPropertyBits=   746588179
         BackColor       =   -2147483633
         ForeColor       =   -2147483630
         DisplayStyle    =   5
         Size            =   "3201;661"
         Value           =   "0"
         Caption         =   "Due Date"
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Specify Date Range"
      Height          =   1455
      Left            =   240
      TabIndex        =   0
      Top             =   1800
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   2
         Top             =   960
         Width           =   1815
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
         Left            =   960
         MaxLength       =   10
         TabIndex        =   1
         Text            =   "01/01/1980"
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   960
         Width           =   180
      End
      Begin VB.Label lblSpecifyDateRange 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   360
      End
   End
End
Attribute VB_Name = "frmViewInvDtRange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
   Me.Top = 50
   Me.Left = 50
   Me.Height = 4600
   Me.Width = 3975
   Me.BackColor = MODULEBACKCOLOR
   Frame1.BackColor = Me.BackColor
   Frame2.BackColor = Me.BackColor

   txtToDate.text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Me.MousePointer = vbArrow
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'frmMMain.fraCmdButton.Enabled = True
   Unload Me
End Sub

Private Sub GenerateReportButton_Click()
   Dim datetype As Integer

   If optDueDate.Value = False And optIssueDate.Value = False Then
      MsgBox "Please select the date type", vbExclamation + vbOKOnly, "Select Date Type"

   ElseIf txtFromDate.text = "" Or txtToDate.text = "" Then

      MsgBox "Please enter both the dates", vbExclamation + vbOKOnly, "Enter Both Dates"

   Else
      'If the Due Date is selected, then pass 0 as the datetype to Crystal Reports and
      'if the Invoice date is selected, then pass 1
      If optDueDate.Value = True Then
         datetype = 0
      Else: datetype = 1

      End If

      ' Passing the from and to date values to Crystal Reports
      Dim reportApp As New CRAXDRT.Application
      Dim Report As CRAXDRT.Report

      Set Report = reportApp.OpenReport(App.Path & szReportPath & "\SalesInvoicePreview2.rpt")
      Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws

      Report.EnableParameterPrompting = False
      Report.DiscardSavedData

      Report.ParameterFields(1).AddCurrentValue CDate(txtFromDate.text)
      Report.ParameterFields(2).AddCurrentValue CDate(txtToDate.text)
      Report.ParameterFields(3).AddCurrentValue datetype

      Load frmReport
      frmReport.LoadReportViewer Report
   End If
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
