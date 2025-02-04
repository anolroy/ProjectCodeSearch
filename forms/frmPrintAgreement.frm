VERSION 5.00
Begin VB.Form frmPrintAgreement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Agreement"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7140
   Icon            =   "frmPrintAgreement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7140
   Begin VB.Frame Frame2 
      Caption         =   "Print"
      Height          =   1140
      Left            =   45
      TabIndex        =   3
      Top             =   1935
      Width           =   7035
      Begin VB.CommandButton cmdPrintAgreement 
         Caption         =   "&Print Agreement"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1395
         TabIndex        =   5
         Top             =   360
         Width           =   1890
      End
      Begin VB.CommandButton cmdClose2 
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4050
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Print Selected"
      Height          =   1905
      Left            =   90
      TabIndex        =   0
      Top             =   45
      Width           =   6990
      Begin VB.OptionButton optAllProperty 
         Caption         =   "All Property"
         Height          =   465
         Left            =   1350
         TabIndex        =   2
         Top             =   855
         Width           =   1635
      End
      Begin VB.OptionButton optSelectedProperty 
         Caption         =   "Selected Property"
         Height          =   465
         Left            =   1350
         TabIndex        =   1
         Top             =   405
         Value           =   -1  'True
         Width           =   1635
      End
   End
End
Attribute VB_Name = "frmPrintAgreement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strClientID As String
Dim strPropertyID As String

Private Sub cmdClose2_Click()
    Unload Me
End Sub

Private Sub cmdPrintAgreement_Click()
   Dim reportApp As New CRAXDRT.Application
   Dim Report As CRAXDRT.Report
   Dim i As Integer

   Set Report = reportApp.OpenReport(App.Path & szReportPath & "\PrintAgreement.rpt")
   Report.Database.Tables(1).ConnectionProperties.Item("Database Password") = accessDBPws
    
   strPropertyID = "ALL"
    
   strClientID = frmClientNew4.txtClientID
   For i = 0 To frmClientNew4.flxPropertySelection1.Rows - 1
        If frmClientNew4.flxPropertySelection1.TextMatrix(i, 0) = "X" Then
            strPropertyID = frmClientNew4.flxPropertySelection1.TextMatrix(i, 1)
            Exit For
        End If
   Next
   If optSelectedProperty.Value = False Then
        strPropertyID = "ALL"
   End If
   Report.EnableParameterPrompting = False
   Report.DiscardSavedData
   Report.ParameterFields(1).AddCurrentValue strClientID
   Report.ParameterFields(2).AddCurrentValue strPropertyID


'   Report.ParameterFields(1).AddCurrentValue CDate(Format(txtSCYRRStDt.text, "dd mmmm yyyy"))
'   Report.ParameterFields(2).AddCurrentValue CDate(Format(txtSCYRREnDt.text, "dd mmmm yyyy"))

   Load frmReport
   frmReport.LoadReportViewer Report
End Sub

Private Sub Form_Load()
    Me.BackColor = MODULEBACKCOLOR
    Frame1.BackColor = MODULEBACKCOLOR
    Frame2.BackColor = MODULEBACKCOLOR
    optSelectedProperty.BackColor = MODULEBACKCOLOR
    optAllProperty.BackColor = MODULEBACKCOLOR
End Sub
