VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmStartUp 
   Caption         =   "Prestige Import Utility"
   ClientHeight    =   2205
   ClientLeft      =   6255
   ClientTop       =   6360
   ClientWidth     =   7725
   Icon            =   "frmStartUp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   7725
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   240
      TabIndex        =   3
      Top             =   120
      Width           =   7335
      Begin VB.CommandButton cmdImportData 
         Caption         =   "Import Standing Data >>"
         Height          =   375
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   5295
      End
      Begin VB.CommandButton cmdImportPI 
         Caption         =   "Import Transactions >>"
         Height          =   375
         Left            =   1680
         TabIndex        =   2
         Top             =   1200
         Width           =   5295
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E5E5E5&
         BackStyle       =   0  'Transparent
         Caption         =   "Property Company:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin MSForms.ComboBox cboShopCentre 
         Height          =   330
         Left            =   1680
         TabIndex        =   0
         Top             =   240
         Width           =   5295
         VariousPropertyBits=   679495707
         BackColor       =   15066597
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "9340;582"
         TextColumn      =   2
         ColumnCount     =   2
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0"
      End
   End
End
Attribute VB_Name = "frmStartUp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboShopCentre_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cboShopCentre_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
 If cboShopCentre.ListCount = 1 Then
      MsgBox "No company has been set up yet for the application. Please contact PCM.", vbCritical + vbOKOnly, "No company data"
   End If
End Sub

Private Sub cmdImportData_Click()
   If cboShopCentre.Text = "" Then
      cboShopCentre.SetFocus
      Exit Sub
   End If

   frmMain.Show
End Sub

Private Sub cmdImportPI_Click()
   If cboShopCentre.Text = "" Then
      cboShopCentre.SetFocus
      Exit Sub
   End If

   frmPurchaseInvoice.Show
   frmPurchaseInvoice.cmdBrowse.SetFocus
End Sub

Private Sub Form_Load()
   cboShopCentre.Clear
   cboShopCentre.Column() = modLists.loadCompanies

   If UCase(SystemUser) = "SAMRAT" And UCase(WS_Name) = "WS1" Then
      cboShopCentre.ListIndex = 0
      cmdImportData_Click
   End If
End Sub
'
'Private Function doLoadChild() As Boolean
' If Trim(cboShopCentre.Text) = "" Or Trim(cboShopCentre.Value) = "" Then
'      MsgBox ("Please select a company before continuing")
'      doLoadChild = False
'   Else
'      Me.Enabled = False
'      doLoadChild = True
' End If
'
'End Function
