VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAddBankTrans 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Bank Transactions"
   ClientHeight    =   10695
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   14220
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10695
   ScaleWidth      =   14220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraBkInput 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Index           =   0
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   12495
      Begin VB.CommandButton cmdUnitListBk 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   21
         Top             =   1116
         Width           =   255
      End
      Begin VB.TextBox txtUnitBk 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   1116
         Width           =   1065
      End
      Begin VB.TextBox txtUnitNameBk 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1116
         Width           =   2625
      End
      Begin VB.TextBox txtDeptBkName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   472
         Width           =   3045
      End
      Begin VB.TextBox txtBkAcName 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   3030
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   472
         Width           =   2625
      End
      Begin VB.TextBox txtNCNameBk 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   8670
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   120
         Width           =   3045
      End
      Begin VB.TextBox txtTotalBk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "0.00"
         Top             =   1440
         Width           =   1665
      End
      Begin VB.TextBox txtReference 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   3750
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1440
         Width           =   1905
      End
      Begin VB.TextBox txtBkAc 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   1680
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   472
         Width           =   1065
      End
      Begin VB.CommandButton cmdTaxListBk 
         Caption         =   "..."
         Height          =   285
         Index           =   0
         Left            =   11310
         TabIndex        =   12
         Top             =   1116
         Width           =   405
      End
      Begin VB.CommandButton cmdDeptBk 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   8385
         TabIndex        =   11
         Top             =   472
         Width           =   255
      End
      Begin VB.CommandButton cmdNCBk 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   8385
         TabIndex        =   10
         Top             =   120
         Width           =   255
      End
      Begin VB.CommandButton cmdBkList 
         Caption         =   "..."
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   2760
         TabIndex        =   9
         Top             =   472
         Width           =   255
      End
      Begin VB.CommandButton cmdUpdateBk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "&OK"
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   10740
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtVatBk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   10080
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1116
         Width           =   1185
      End
      Begin VB.TextBox txtDetailsBk 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   7320
         MaxLength       =   254
         TabIndex        =   6
         Top             =   794
         Width           =   4395
      End
      Begin VB.TextBox txtNetBk 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   7320
         TabIndex        =   5
         Top             =   1116
         Width           =   1665
      End
      Begin VB.TextBox txtDateBk 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   4
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox txtDeptBk 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   472
         Width           =   1050
      End
      Begin VB.TextBox txtNCBk 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   7320
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   1050
      End
      Begin VB.CommandButton cmdUpdateBk 
         BackColor       =   &H00FFC0C0&
         Caption         =   "C&lear"
         Height          =   285
         Index           =   1
         Left            =   9480
         MaskColor       =   &H00FFFFFF&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1440
         Width           =   975
      End
      Begin MSForms.ComboBox cboBRPProperty 
         Height          =   285
         Left            =   1680
         TabIndex        =   35
         Top             =   794
         Width           =   3975
         VariousPropertyBits=   1753237529
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7011;503"
         TextColumn      =   2
         ColumnCount     =   8
         ListRows        =   0
         cColumnInfo     =   8
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1411;3527;3527;0;0;0;0;0"
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Unit:"
         Height          =   195
         Index           =   8
         Left            =   840
         TabIndex        =   34
         Top             =   1110
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   33
         Top             =   795
         Width           =   645
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Details:"
         Height          =   195
         Index           =   11
         Left            =   6600
         TabIndex        =   32
         Top             =   794
         Width           =   540
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Total:"
         Height          =   195
         Index           =   4
         Left            =   6600
         TabIndex        =   31
         Top             =   1440
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Reference:"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   30
         Top             =   1440
         Width           =   750
      End
      Begin MSForms.ComboBox cboBRPClient 
         Height          =   285
         Left            =   1680
         TabIndex        =   29
         Top             =   120
         Width           =   3975
         VariousPropertyBits=   1753237529
         BorderStyle     =   1
         DisplayStyle    =   3
         Size            =   "7011;503"
         TextColumn      =   2
         ColumnCount     =   8
         ListRows        =   0
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontEffects     =   1073750016
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1763"
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   5
         Left            =   840
         TabIndex        =   28
         Top             =   120
         Width           =   435
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "VAT:"
         Height          =   195
         Index           =   12
         Left            =   9480
         TabIndex        =   27
         Top             =   1110
         Width           =   330
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Net:"
         Height          =   195
         Index           =   10
         Left            =   6600
         TabIndex        =   26
         Top             =   1116
         Width           =   300
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/C:"
         Height          =   195
         Index           =   7
         Left            =   6600
         TabIndex        =   25
         Top             =   120
         Width           =   315
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund:"
         Height          =   195
         Index           =   6
         Left            =   6600
         TabIndex        =   24
         Top             =   472
         Width           =   390
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         Height          =   195
         Index           =   3
         Left            =   840
         TabIndex        =   23
         Top             =   1440
         Width           =   375
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "BankA/C:"
         Height          =   195
         Index           =   2
         Left            =   840
         TabIndex        =   22
         Top             =   472
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmAddBankTrans"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
   Me.Top = (frmMMain.Height / 2) - (Me.Height / 2) - 400
   Me.Left = (frmMMain.Width / 2) - (Me.Width / 2) - 400
End Sub
