VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmReceiptEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Receipt"
   ClientHeight    =   10800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15150
   BeginProperty Font 
      Name            =   "Myriad Web"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmReceiptEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10800
   ScaleWidth      =   15150
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel1 
      Caption         =   "&Cancel"
      Height          =   300
      Left            =   5310
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2520
      Width           =   1155
   End
   Begin VB.PictureBox fraList 
      BackColor       =   &H80000004&
      Height          =   3870
      Index           =   0
      Left            =   8730
      ScaleHeight     =   3810
      ScaleWidth      =   5340
      TabIndex        =   45
      Top             =   1260
      Visible         =   0   'False
      Width           =   5400
      Begin VB.CommandButton cmdClose1 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4950
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   20
         Width           =   300
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxSupplier 
         Height          =   3255
         Index           =   0
         Left            =   15
         TabIndex        =   48
         Top             =   525
         Width           =   5280
         _ExtentX        =   9313
         _ExtentY        =   5741
         _Version        =   393216
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   -2147483634
         BackColorSel    =   12648447
         ForeColorSel    =   -2147483630
         BackColorBkg    =   16777215
         GridColor       =   14737632
         GridLinesFixed  =   1
         SelectionMode   =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label lblPayeeFlxConfigured 
         Caption         =   "NOT"
         Height          =   495
         Index           =   0
         Left            =   1515
         TabIndex        =   53
         Top             =   1680
         Width           =   1095
      End
      Begin MSForms.Label lblSearch0 
         Height          =   195
         Index           =   0
         Left            =   30
         TabIndex        =   52
         Top             =   0
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.Label lblSearch1 
         Height          =   195
         Index           =   0
         Left            =   1425
         TabIndex        =   51
         Top             =   15
         Width           =   735
         VariousPropertyBits=   8388627
         Caption         =   "dynamic"
         Size            =   "1296;353"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch1 
         Height          =   255
         Left            =   30
         TabIndex        =   50
         Top             =   240
         Width           =   1290
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2275;450"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearch2 
         Height          =   255
         Left            =   1365
         TabIndex        =   49
         Top             =   240
         Width           =   3510
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6191;450"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   2
         Left            =   0
         Top             =   30
         Width           =   4905
      End
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "&Delete"
      Height          =   435
      Left            =   4360
      TabIndex        =   11
      Top             =   4995
      Width           =   1335
   End
   Begin VB.Frame fraEdit 
      BorderStyle     =   0  'None
      Height          =   2850
      Left            =   840
      TabIndex        =   25
      Top             =   45
      Width           =   5775
      Begin VB.TextBox txtClient 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   80
         TabIndex        =   38
         Top             =   90
         Width           =   4455
      End
      Begin VB.TextBox txtRef 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   4
         Top             =   1800
         Width           =   4455
      End
      Begin VB.TextBox txtLesseeName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   435
         Width           =   2775
      End
      Begin VB.TextBox txtDate 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4005
         TabIndex        =   6
         Top             =   2145
         Width           =   1335
      End
      Begin VB.CommandButton cmdLesseeLookup 
         Caption         =   "v"
         Height          =   255
         Left            =   2505
         TabIndex        =   0
         Top             =   450
         Width           =   375
      End
      Begin VB.CommandButton cmdUpate 
         Caption         =   "&OK"
         Height          =   300
         Left            =   3330
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2475
         Width           =   1110
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   5
         Top             =   2145
         Width           =   2055
      End
      Begin VB.TextBox txtDetails 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         MaxLength       =   80
         TabIndex        =   3
         Top             =   1440
         Width           =   4455
      End
      Begin VB.TextBox txtLessee 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   435
         Width           =   1695
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   37
         Top             =   90
         Width           =   435
      End
      Begin MSForms.Label lblRptPostingDate 
         Height          =   285
         Left            =   5355
         TabIndex        =   36
         Top             =   2160
         Width           =   225
         ForeColor       =   8421504
         BackColor       =   16761024
         Caption         =   " P"
         Size            =   "397;503"
         FontName        =   "Myriad Web"
         FontEffects     =   1073741825
         FontHeight      =   195
         FontCharSet     =   0
         FontPitchAndFamily=   2
         FontWeight      =   700
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Reference"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   35
         Top             =   1755
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lessee"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   435
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Date"
         Height          =   255
         Index           =   1
         Left            =   3555
         TabIndex        =   33
         Top             =   2145
         Width           =   495
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Details"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   32
         Top             =   1395
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   31
         Top             =   2115
         Width           =   975
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fund"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   30
         Top             =   1095
         Width           =   360
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Bank A/C"
         Height          =   195
         Index           =   6
         Left            =   120
         TabIndex        =   29
         Top             =   795
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Receipt Type"
         Height          =   195
         Index           =   11
         Left            =   120
         TabIndex        =   28
         Top             =   2475
         Width           =   915
      End
      Begin MSForms.ComboBox cmbRptAmtType 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Top             =   2475
         Width           =   2055
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "3625;503"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "0"
      End
      Begin MSForms.ComboBox cmbBankAc 
         Height          =   285
         Left            =   1200
         TabIndex        =   1
         Top             =   765
         Width           =   4455
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "7858;503"
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "1058"
      End
      Begin MSForms.ComboBox cmbFund 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Top             =   1110
         Width           =   4455
         VariousPropertyBits=   1753237531
         BorderStyle     =   1
         DisplayStyle    =   7
         Size            =   "7858;503"
         BoundColumn     =   0
         TextColumn      =   2
         ColumnCount     =   2
         ListRows        =   20
         cColumnInfo     =   1
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
         Object.Width           =   "705"
      End
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit"
      Height          =   435
      Left            =   2240
      TabIndex        =   10
      Top             =   4995
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox fmeTenantLookup 
      Appearance      =   0  'Flat
      BackColor       =   &H00EFEFEF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   4830
      ScaleHeight     =   3195
      ScaleWidth      =   5430
      TabIndex        =   13
      Top             =   5700
      Visible         =   0   'False
      Width           =   5460
      Begin VB.CommandButton cmdClientList 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4725
         TabIndex        =   39
         Top             =   90
         Width           =   300
      End
      Begin VB.CommandButton cmdPropertyList 
         Caption         =   ".."
         BeginProperty Font 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4725
         TabIndex        =   40
         Top             =   420
         Width           =   300
      End
      Begin VB.CommandButton cmdGridTenantLookup 
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5115
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   25
         Width           =   285
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxTenant 
         Height          =   1905
         Left            =   30
         TabIndex        =   46
         Top             =   1275
         Width           =   5325
         _ExtentX        =   9393
         _ExtentY        =   3360
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         BackColorFixed  =   13553358
         ForeColorFixed  =   16777215
         BackColorSel    =   12648447
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         GridColor       =   -2147483638
         WordWrap        =   -1  'True
         HighLight       =   2
         GridLinesFixed  =   1
         ScrollBars      =   2
         SelectionMode   =   1
         Appearance      =   0
         BandDisplay     =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Myriad Web"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   9
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSForms.TextBox txtClientList 
         Height          =   285
         Left            =   765
         TabIndex        =   43
         Tag             =   "ALL"
         Top             =   90
         Width           =   4275
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "7541;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtPropertyList 
         Height          =   285
         Left            =   765
         TabIndex        =   41
         Tag             =   "ALL"
         Top             =   420
         Width           =   4005
         VariousPropertyBits=   679495711
         BorderStyle     =   1
         Size            =   "7064;503"
         Value           =   "ALL"
         SpecialEffect   =   0
         FontName        =   "Myriad Web"
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin MSForms.TextBox txtSearchName 
         Height          =   270
         Left            =   1500
         TabIndex        =   44
         Top             =   990
         Width           =   3855
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "6800;476"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label50 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client:"
         Height          =   195
         Index           =   13
         Left            =   30
         TabIndex        =   18
         Top             =   135
         Width           =   465
      End
      Begin VB.Label Label50 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Property:"
         Height          =   195
         Index           =   14
         Left            =   30
         TabIndex        =   17
         Top             =   450
         Width           =   645
      End
      Begin MSForms.TextBox txtSearchTenant 
         Height          =   270
         Left            =   120
         TabIndex        =   42
         Top             =   990
         Width           =   1350
         VariousPropertyBits=   679495707
         BorderStyle     =   1
         Size            =   "2381;476"
         SpecialEffect   =   0
         FontHeight      =   165
         FontCharSet     =   0
         FontPitchAndFamily=   2
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name"
         Height          =   195
         Index           =   2
         Left            =   1500
         TabIndex        =   16
         Top             =   780
         Width           =   405
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A/C"
         Height          =   195
         Index           =   1
         Left            =   75
         TabIndex        =   15
         Top             =   780
         Width           =   270
      End
      Begin VB.Shape Shape4 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   6  'Inside Solid
         FillColor       =   &H00C0FFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Index           =   6
         Left            =   30
         Top             =   780
         Width           =   5310
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   435
      Left            =   120
      TabIndex        =   9
      Top             =   4995
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cl&ose"
      Height          =   435
      Left            =   6480
      TabIndex        =   12
      Top             =   4995
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid flxReceiptSplit 
      Height          =   1665
      Left            =   120
      TabIndex        =   19
      Top             =   3255
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   2937
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      BackColorFixed  =   13553358
      ForeColorFixed  =   16777215
      BackColorSel    =   12648447
      ForeColorSel    =   0
      BackColorBkg    =   16777215
      GridColor       =   -2147483638
      WordWrap        =   -1  'True
      HighLight       =   2
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Myriad Web"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   9
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Amount"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   5
      Left            =   6480
      TabIndex        =   24
      Top             =   3000
      Width           =   555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Due Date"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   4
      Left            =   5400
      TabIndex        =   23
      Top             =   3000
      Width           =   660
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Details"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   3
      Left            =   2520
      TabIndex        =   22
      Top             =   3000
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fund"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   240
      TabIndex        =   21
      Top             =   3000
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "No"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   20
      Top             =   3000
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00E0E0E0&
      BorderColor     =   &H00C0C0C0&
      BorderStyle     =   6  'Inside Solid
      FillColor       =   &H00E0E0E0&
      FillStyle       =   0  'Solid
      Height          =   255
      Index           =   20
      Left            =   120
      Top             =   2985
      Width           =   7695
   End
End
Attribute VB_Name = "frmReceiptEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private iRecords As Integer
Dim szSel As String
Public TransactionID As Long
Public InvoiceNO As String
Public BoolReconciled As Boolean
Private Sub cboClientList_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then fmeTenantLookup.Visible = False
End Sub

Private Sub cboPropertyList_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then fmeTenantLookup.Visible = False
End Sub

Private Sub cmbBankAc_Change()
    'added by anol 21 July 2015
    If txtAmount.Locked Then
      cmbBankAc.Locked = True
   End If
End Sub

Private Sub cmbBankAc_Click()
    'added by anol 21 July 2015
    If txtAmount.Locked Then
      cmbBankAc.Locked = True
   End If
End Sub

Private Sub cmbBankAc_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then Unload Me
   If KeyAscii = 13 Then
        FocusControl cmbFund
   End If
End Sub

Private Sub cmbFund_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then Unload Me
   If KeyAscii = 13 Then
        FocusControl txtDetails
   End If
End Sub

Private Sub cmbRptAmtType_KeyPress(KeyAscii As MSForms.ReturnInteger)
    If KeyAscii = 13 Then
        FocusControl cmdUpate
    End If
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub cmdCancel1_Click()
    Dim adoConn As New ADODB.Connection
    adoConn.Open getConnectionString
    ConfigFlxReceiptSplit
    LoadFlxReceiptSplit adoConn
    adoConn.Close
    fraEdit.Enabled = True
    cmdUpate.Enabled = True
End Sub

Private Sub cmdClientList_Click()
    fmeTenantLookup.Enabled = False
    LoadflxClient

   'tabTenant.Enabled = False 'it is already false by other lessee grid
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   'cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   'Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = fmeTenantLookup.Left + 500 'tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = fmeTenantLookup.Top + 200 'tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   FocusControl txtSearch1
   szSel = "Client"
End Sub
Private Sub LoadflxClient()
   flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 70
   flxSupplier(0).ColWidth(1) = 1500
   flxSupplier(0).ColWidth(2) = 3300
   flxSupplier(0).ColAlignment = vbLeftJustify


   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(1)

   lblSearch0(0).Caption = "Client ID"
   lblSearch1(0).Caption = "Client Name"
   
   
   flxSupplier(0).RowHeight(0) = 0


   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

    szSQL = "SELECT CLIENTID, CLIENTNAME, CT FROM   CLIENT ORDER BY CLIENTID;"
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Dim iRows As Integer
   flxSupplier(0).Rows = 2
   iRows = 1
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = "ALL"
      flxSupplier(0).TextMatrix(iRows, 2) = "ALL"
      flxSupplier(0).AddItem ""
   iRows = 2
   While Not adoRst.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = adoRst.Fields.Item("CLIENTID").Value
      flxSupplier(0).TextMatrix(iRows, 2) = adoRst.Fields.Item("CLIENTNAME").Value
      If Not adoRst.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRst.MoveNext
   Wend
 
   Set adoRst = Nothing
   Set adoConn = Nothing
   Exit Sub

Error_Handler:
  
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdClose1_Click()
     fraList(0).Visible = False
     fmeTenantLookup.Enabled = True
End Sub

Private Sub cmdDelete_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String
   'issue 571 Note 1153
   'Modified by anol 30 Aug 2015
    Dim iSelection As Integer
   If txtAmount.Locked Then
      ShowMsgInTaskBar "The transactions has been reconciled and you cannot delete the line", "Y", "N"
      Exit Sub
   End If
    'iRecords is keeping the record cound whenever you are load the split grid.
    If frmDemands3.tabPayment.Tab = 0 Then
                If iRecords < 2 Then 'when there is only one split
           
                If MsgBox("Do you wish to delete the split line?", vbQuestion + vbYesNo, "Delete a split line") = vbNo Then Exit Sub
                MsgBox "It is not possible to delete the last line of a lessee receipt. The receipt amount will be automatically set to zero."
                    If flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 0) = "" Then
                        MsgBox "Please select a transaction to delete", vbInformation, "Warning"
                        Exit Sub
                    End If
                adoConn.Open getConnectionString
                adoConn.BeginTrans
                    szSQL = "UPDATE tlbReceiptSplit SET Amount = 0, OSAmount = 0 " & _
                         "WHERE TransactionID = '" & flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 0) & "';"
                     adoConn.Execute szSQL
                If frmDemands3.sEditSRR = 2 Then 'when clicking lower grid on SI form
                    'here the type is 3,4
                     szSQL = "UPDATE tlbReceipt " & _
                         "SET    Amount   = 0, " & _
                                "OSAmount = 0, NLPost =false " & _
                         "WHERE TransactionID = " & _
                                TransactionID & ";"
                     adoConn.Execute szSQL
                     szSQL = "UPDATE NLPosting AS N " & _
                         "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & TransactionID & "' AND " & _
                         " (N.TRANSACTION_TYPE = 3 OR  N.TRANSACTION_TYPE = 4)"
                         'checked
                     adoConn.Execute szSQL
                     
                Else   'when clicking upper grid on SI form
                     'here the Type is 23
                    
        
                     szSQL = "UPDATE tlbReceipt " & _
                         "SET    Amount   = 0, " & _
                                "OSAmount = 0, NLPost =false " & _
                         "WHERE TransactionID = " & _
                                TransactionID & ";"
                     adoConn.Execute szSQL
                    
                     ''For Type 23
        ''Fixed by anol 20170801 Issue related to 422
        ''.Fields.Item("PARENT_RECORD").Value = receipt("TransactionID").Value
        '' .Fields.Item("TRANS_ID").Value = receipt("TransactionID").Value
        ''
        ''For 3,4 This is posting from Demand receipt table
        '' .Fields.Item("PARENT_RECORD").Value = ReceiptSplit("TransactionID").Value
        '' .Fields.Item("TRANS_ID").Value = receipt("TransactionID").Value
        
                     szSQL = "UPDATE NLPosting AS N SET    N.DeleteFlag = TRUE " & _
                         " WHERE  N.TRANS_ID = '" & TransactionID & "' AND " & _
                         " N.TRANSACTION_TYPE = 23 ;"
                          
                     adoConn.Execute szSQL
                End If
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit."
                     Exit Sub
                Else
                     adoConn.CommitTrans
                       If frmDemands3.sEditSRR = 2 Then
                            frmDemands3.ConfigFlxCrPoA
                            frmDemands3.LoadFlxCrPoA adoConn
                        Else
                            frmDemands3.ConfigFlxSPayment
                            frmDemands3.LoadFlxSPayment adoConn
                            'need to check if
                        End If
                        cmdDelete.Enabled = False
                        txtAmount.text = "0.00"
                        LoadFlxReceiptSplit adoConn
                        HighLightRowFlxGrid flxReceiptSplit, 1
                        adoConn.Close
                        Set adoConn = Nothing
                End If
                '--------------------------------------------------------------------------------------------
                '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
                   adoConn.Open getConnectionString
                   Export_SRnSRR_2_NL adoConn
                   adoConn.Close
                   frmDemands3.cmdEditReceipt.Enabled = False
                   Exit Sub
           End If
        ' This part is for if there is more than one split line
                If flxReceiptSplit.row <= 0 Then
                   ShowMsgInTaskBar "Select a split to delete."
                   Exit Sub
                End If
                iSelection = flxReceiptSplit.row
                If MsgBox("Do you wish to delete the split line?", vbQuestion + vbYesNo, "Delete a split line") = vbNo Then Exit Sub
                adoConn.Open getConnectionString
                adoConn.BeginTrans
                szSQL = "DELETE * " & _
                        "FROM tlbReceiptSplit " & _
                        "WHERE TransactionID = '" & _
                            flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 0) & "';"
                adoConn.Execute szSQL
                If frmDemands3.sEditSRR = 2 Then 'lower credit grid
                'Type 3,4
                     szSQL = "UPDATE tlbReceipt " & _
                             "SET Amount = Amount - " & Val(txtAmount.text) & ", " & _
                                    "OSAmount = OSAmount - " & Val(txtAmount.text) & " , NLPost = false " & _
                             "WHERE TransactionID = " & _
                                    TransactionID & ";"
                     adoConn.Execute szSQL
                     
                     
                     szSQL = "UPDATE NLPosting AS N " & _
                              "SET    N.DeleteFlag = TRUE " & _
                              "WHERE  N.TRANS_ID = '" & TransactionID & "' AND " & _
                               " (N.TRANSACTION_TYPE = 3 OR  N.TRANSACTION_TYPE = 4) "
                            
                          adoConn.Execute szSQL
                Else
                      szSQL = "UPDATE tlbReceipt " & _
                             "SET    Amount   = Amount   - " & Val(txtAmount.text) & ", " & _
                                    "OSAmount = OSAmount - " & Val(txtAmount.text) & " , NLPost = false " & _
                             "WHERE TransactionID = " & _
                                    TransactionID & ";"
                     adoConn.Execute szSQL
                     
                     
        
                     szSQL = "UPDATE NLPosting AS N " & _
                              "SET    N.DeleteFlag = TRUE " & _
                              " WHERE  N.TRANS_ID = '" & TransactionID & "' AND " & _
                              " N.TRANSACTION_TYPE = 23"
                              
                     adoConn.Execute szSQL
                End If
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit."
                     Exit Sub
                Else
                    adoConn.CommitTrans
                    LoadFlxReceiptSplit adoConn
                    If frmDemands3.sEditSRR = 2 Then
                        frmDemands3.ConfigFlxCrPoA
                        frmDemands3.LoadFlxCrPoA adoConn
                    Else
                        frmDemands3.ConfigFlxSPayment
                        frmDemands3.LoadFlxSPayment adoConn
                    End If
                    HighLightRowFlxGrid flxReceiptSplit, 1
                    adoConn.Close
                    Set adoConn = Nothing
                End If
                
        
        '--------------------------------------------------------------------------------------------
        '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
           adoConn.Open getConnectionString
           Export_SRnSRR_2_NL adoConn
           adoConn.Close
           frmDemands3.cmdEditReceipt.Enabled = False
           ShowMsgInTaskBar "The split line has been deleted.", "Y", "P"
        
          
           If iSelection <= iRecords Then
                flxReceiptSplit.row = iSelection
                HighLightRowFlxGrid flxReceiptSplit, flxReceiptSplit.row
           Else
                flxReceiptSplit.row = iRecords
                HighLightRowFlxGrid flxReceiptSplit, iRecords
           End If
           txtAmount.text = Format(flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 6), "0.00")
'           adoConn.Close
'           Set adoConn = Nothing
   End If
   If frmDemands3.tabPayment.Tab = 1 Then
         If iRecords < 2 Then 'when there is only one split
           
                If MsgBox("Do you wish to delete the split line?", vbQuestion + vbYesNo, "Delete a split line") = vbNo Then Exit Sub
                MsgBox "It is not possible to delete the last line of a lessee receipt. The receipt amount will be automatically set to zero."
                    If flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 0) = "" Then
                        MsgBox "Please select a transaction to delete", vbInformation, "Warning"
                        Exit Sub
                    End If
                    adoConn.Open getConnectionString
                    adoConn.BeginTrans
                    szSQL = "UPDATE tlbReceiptSplit SET Amount = 0, OSAmount = 0 " & _
                         "WHERE TransactionID = '" & flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 0) & "';"
                     adoConn.Execute szSQL
              
                    'here the type is 3,4,23
                     szSQL = "UPDATE tlbReceipt " & _
                         "SET    Amount   = 0, " & _
                                "OSAmount = 0, NLPost =false " & _
                         "WHERE TransactionID = " & _
                                TransactionID & ";"
                     adoConn.Execute szSQL
                     szSQL = "UPDATE NLPosting AS N " & _
                         "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & TransactionID & "' AND " & _
                         " (N.TRANSACTION_TYPE = 3 OR  N.TRANSACTION_TYPE = 4 OR N.TRANSACTION_TYPE = 23 )"
                         'checked
                     adoConn.Execute szSQL
                
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit."
                     Exit Sub
                Else
                     adoConn.CommitTrans
                     'frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 7) = Format(frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 7) - Val(txtAmount.text), "0.00")
                     'because you are deleting all splits
                     frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 7) = "0.00"
                     cmdDelete.Enabled = False
                     txtAmount.text = "0.00"
                     LoadFlxReceiptSplit adoConn
                     HighLightRowFlxGrid flxReceiptSplit, 1
                     adoConn.Close
                     Set adoConn = Nothing
                End If
                '--------------------------------------------------------------------------------------------
                '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
                   adoConn.Open getConnectionString
                   Export_SRnSRR_2_NL adoConn
                   adoConn.Close
                   frmDemands3.cmdEditReceipt.Enabled = False
                   Exit Sub
           End If
        ' This part is for if there is more than one split line
        'This part is never reaching code because you are starting with zero values for receipt tab1
                If flxReceiptSplit.row <= 0 Then
                   ShowMsgInTaskBar "Select a split to delete."
                   Exit Sub
                End If
                iSelection = flxReceiptSplit.row
                If MsgBox("Do you wish to delete the split line?", vbQuestion + vbYesNo, "Delete a split line") = vbNo Then Exit Sub
                adoConn.Open getConnectionString
                
                szSQL = "DELETE * " & _
                        "FROM tlbReceiptSplit " & _
                        "WHERE TransactionID = '" & _
                            flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 0) & "';"
                adoConn.Execute szSQL
                'lower credit grid
                'Type 3,4,23
                     szSQL = "UPDATE tlbReceipt " & _
                             "SET Amount = Amount - " & Val(txtAmount.text) & ", " & _
                                    "OSAmount = OSAmount - " & Val(txtAmount.text) & " , NLPost = false " & _
                             "WHERE TransactionID = " & _
                                    TransactionID & ";"
                     adoConn.Execute szSQL
                     
                     
                     szSQL = "UPDATE NLPosting AS N " & _
                              "SET    N.DeleteFlag = TRUE " & _
                              "WHERE  N.TRANS_ID = '" & TransactionID & "' AND " & _
                               " (N.TRANSACTION_TYPE = 3 OR  N.TRANSACTION_TYPE = 4 OR N.TRANSACTION_TYPE = 23) "
                            
                          adoConn.Execute szSQL
               
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit."
                     Exit Sub
                Else
                    adoConn.CommitTrans
                    frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 7) = Format(frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 7) - Val(txtAmount.text), "0.00")
                    LoadFlxReceiptSplit adoConn
                    
                    frmDemands3.ConfigFlxCrPoA
                    frmDemands3.LoadFlxCrPoA adoConn
                    
                    frmDemands3.ConfigFlxSPayment
                    frmDemands3.LoadFlxSPayment adoConn
                    HighLightRowFlxGrid flxReceiptSplit, 1
                    adoConn.Close
                    Set adoConn = Nothing
                End If
                
        
                '--------------------------------------------------------------------------------------------
                '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
                   adoConn.Open getConnectionString
                   Export_SRnSRR_2_NL adoConn
                   adoConn.Close
                   frmDemands3.cmdEditReceipt.Enabled = False
                   ShowMsgInTaskBar "The split line has been deleted.", "Y", "P"
        
          
           If iSelection <= iRecords Then
                flxReceiptSplit.row = iSelection
                HighLightRowFlxGrid flxReceiptSplit, flxReceiptSplit.row
           Else
                flxReceiptSplit.row = iRecords
                HighLightRowFlxGrid flxReceiptSplit, iRecords
           End If
           txtAmount.text = Format(flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 6), "0.00")
           adoConn.Close
           Set adoConn = Nothing
   End If
End Sub

Private Sub cmdEdit_Click()
   cmdSave.Enabled = False

   If flxReceiptSplit.row <= 0 Then
      ShowMsgInTaskBar "Select a split to edit."
      Exit Sub
   End If

   If cmdEdit.Caption = "&Cancel" Then
      fraEdit.Enabled = False
      flxReceiptSplit.Enabled = True
      cmdSave.Enabled = False
      cmdEdit.Caption = "&Edit"
      Exit Sub
   End If

   fraEdit.Enabled = True
   flxReceiptSplit.Enabled = False

   cmdEdit.Caption = "&Cancel"
End Sub

Private Sub cmdGridTenantLookup_Click()
   fmeTenantLookup.Visible = False
   fraEdit.Enabled = True
End Sub

Private Sub cmdPropertyList_Click()
    fmeTenantLookup.Enabled = False
    LoadflxProperty

   'tabTenant.Enabled = False 'it is already false by other lessee grid
   txtSearch1.Visible = True
   txtSearch2.Visible = True

   txtSearch1.text = ""
   txtSearch2.text = ""

   'fraList(0).Width = 5115
   'Picture1.Width = 5815
   'cmdGridUnitLookup(0).Left = fraList(0).Width - cmdGridUnitLookup(0).Width
   'Shape4(2).Width = fraList(0).Width - cmdGridUnitLookup(0).Width - 50
  ' flxSupplier(0).Width = 4695
   fraList(0).Left = fmeTenantLookup.Left + 500 'tabTenant.Left + txtDNC(1).Left
   fraList(0).Top = fmeTenantLookup.Top + 200 'tabTenant.Top + txtDNC(1).Top
   fraList(0).Visible = True
   fraList(0).ZOrder 0
   FocusControl txtSearch1
   szSel = "Property"
End Sub
Private Sub LoadflxProperty()
    flxSupplier(0).Cols = 3
   flxSupplier(0).ColWidth(0) = 70
   flxSupplier(0).ColWidth(1) = 1500
   flxSupplier(0).ColWidth(2) = 3300
   flxSupplier(0).ColAlignment = vbLeftJustify


   lblSearch0(0).Width = 1400
   lblSearch0(0).Left = 50
   lblSearch1(0).Width = 2600
   lblSearch1(0).Left = lblSearch0(0).Left + flxSupplier(0).ColWidth(1)

   lblSearch0(0).Caption = "Property ID"
   lblSearch1(0).Caption = "Property Name"
   
   
   flxSupplier(0).RowHeight(0) = 0


   On Error GoTo Error_Handler

   Dim adoConn As ADODB.Connection
   Dim rRow As Integer, iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String

   Set adoConn = New ADODB.Connection
   adoConn.Open getConnectionString

   If txtClientList.Tag = "ALL" Then
      szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "ORDER BY PropertyID;"
   Else
      szSQL = "SELECT PropertyID, PropertyName, ProAddressLine1, ProPostCode " & _
              "FROM Property " & _
              "WHERE ClientID = '" & txtClientList.Tag & "' " & _
              "ORDER BY PropertyID;"
   End If
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   Dim iRows As Integer
   flxSupplier(0).Rows = 2
   iRows = 1
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = "ALL"
      flxSupplier(0).TextMatrix(iRows, 2) = "ALL"
      flxSupplier(0).AddItem ""
   iRows = 2
   While Not adoRst.EOF
      flxSupplier(0).TextMatrix(iRows, 0) = ""
      flxSupplier(0).TextMatrix(iRows, 1) = adoRst.Fields.Item("PropertyID").Value
      flxSupplier(0).TextMatrix(iRows, 2) = adoRst.Fields.Item("PropertyName").Value
      If Not adoRst.EOF Then flxSupplier(0).AddItem ""
      iRows = iRows + 1
      adoRst.MoveNext
   Wend
 
   Set adoRst = Nothing
   Set adoConn = Nothing
   Exit Sub

Error_Handler:
  
   Set adoRst = Nothing
   Set adoConn = Nothing
End Sub
Private Sub cmdUpate_Click()
    If Trim(txtDate.text) = "" Then
        MsgBox "Please enter a valid date (dd/mm/yyyy).", vbInformation, "Invalid date!"
        FocusControl txtDate
        Exit Sub
    End If
     If lblRptPostingDate.ToolTipText = "" Then
        MsgBox "Please enter a valid posting date", vbInformation, "Warning!!"
        Exit Sub
    End If
   If cmbBankAc.text = "" Then
      ShowMsgInTaskBar "Please select bank account.", "Y", "N"
      FocusControl cmbBankAc
      Exit Sub
   End If
   If cmbRptAmtType.text = "" Then
      ShowMsgInTaskBar "Please select Receipt Type.", "Y", "N"
      FocusControl cmbRptAmtType
      Exit Sub
   End If
   If cmbFund.ListIndex = "-1" Then
      ShowMsgInTaskBar "Please select a fund.", "Y", "N"
      FocusControl cmbFund
      Exit Sub
   End If
   With flxReceiptSplit
        If .row = 0 Then
            .row = 1
        End If
      .TextMatrix(.row, 1) = cmbFund.Column(0)
      .TextMatrix(.row, 3) = cmbFund.text
      .TextMatrix(.row, 4) = txtDetails.text
      .TextMatrix(.row, 5) = txtDate.text
      .TextMatrix(.row, 6) = Format(Val(txtAmount.text), "0.00")
   End With
   
   fraEdit.Enabled = False
   cmdSave.Enabled = True
   FocusControl cmdSave
   flxReceiptSplit.Enabled = True
   flxReceiptSplit.row = 0
   cmdUpate.Enabled = False
   FocusControl cmdSave
End Sub



Private Sub flxReceiptSplit_RowColChange()
   If flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 1) = "" Then
      flxReceiptSplit.row = flxReceiptSplit.row - 1
   End If

   HighLightRowFlxGrid flxReceiptSplit, flxReceiptSplit.row

   cmbFund.ListIndex = FindComboIndex(cmbFund, flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 1), 0)
   txtDetails.text = flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 4)
   txtAmount.text = Format(flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 6), "0.00")

'   If flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 8) <> "" Then
'      txtAmount.Locked = True
'      'added by anol 21 July 2015
'      cmbBankAc.Locked = True
'
'   Else
'      txtAmount.Locked = False
'      'added by anol 21 July 2015
'      cmbBankAc.Locked = False
'   End If
   cmdDelete.Enabled = True
   fraEdit.Enabled = True
   cmdUpate.Enabled = True
End Sub

Private Sub flxSupplier_Click(index As Integer)
   fraList(0).Visible = False
   fmeTenantLookup.Enabled = True
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   If szSel = "Client" Then
        txtClientList.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtClientList.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
        'tabTenant.Enabled = True
        txtPropertyList.Tag = "ALL"
        txtPropertyList.text = "ALL"
        FilterTenantsList
        FocusControl cmdPropertyList
   End If
    If szSel = "Property" Then
        txtPropertyList.Tag = flxSupplier(0).TextMatrix(flxSupplier(0).row, 1)
        txtPropertyList.text = flxSupplier(0).TextMatrix(flxSupplier(0).row, 2)
        FilterTenantsList
        FocusControl txtSearchTenant
   End If
   flxSupplier(0).Clear
   adoConn.Close
End Sub

Private Sub flxTenant_Click()
   fmeTenantLookup.Visible = False

   txtLessee.text = flxTenant.TextMatrix(flxTenant.row, 1)
   txtLesseeName.text = flxTenant.TextMatrix(flxTenant.row, 2)
   txtClient.text = flxTenant.TextMatrix(flxTenant.row, 3)
   'if you change lessee then clinet may change then bank may change so nee dto load the bank again
   Dim adoConn As New ADODB.Connection
   adoConn.Open getConnectionString
   LoadBankAgain adoConn
   adoConn.Close
   fraEdit.Enabled = True
   FocusControl cmbBankAc
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdLesseeLookup_Click()
   fmeTenantLookup.Left = txtLessee.Left
   fmeTenantLookup.Top = txtLessee.Top + txtLessee.Height
   fmeTenantLookup.Visible = True
   fmeTenantLookup.ZOrder 0
   flxTenant.Visible = True
   FocusControl cmdClientList
   txtSearchTenant.text = ""
   txtSearchName.text = ""
   FilterTenantsList
   FocusControl txtSearchTenant
   fraEdit.Enabled = False
End Sub
Private Function FilterTenantsList() As String
   
   Dim szWhere As String
   Dim szOrderBy As String
   ConfigflxTenant
   szOrderBy = "LeaseDetails.SageAccountNumber ASC"
   
   If txtClientList.text = "ALL" And txtPropertyList.text = "ALL" Then _
      szWhere = ""
      
   If txtClientList.text <> "ALL" And txtPropertyList.text = "ALL" Then _
      szWhere = "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
      
   If txtClientList.text = "ALL" And txtPropertyList.text <> "ALL" Then _
      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' "
      
   If txtClientList.text <> "ALL" And txtPropertyList.text <> "ALL" Then _
      szWhere = "AND PROPERTY.PROPERTYID = '" & txtPropertyList.Tag & "' " & _
                         "AND CLIENT.CLIENTID = '" & txtClientList.Tag & "' "
                         
   Dim Filter As String
   'Wild card search has been implemented by anol
   'issue 0000445: Searching issues found through out Prestige
   'Date 22 Feb 2015
   If Len(txtSearchTenant.text) > 0 Then
      txtSearchName.text = ""
      Filter = " SageAccountNumber LIKE '%" + UCase(txtSearchTenant.text) + "*'"
   End If
   
   If Len(txtSearchName.text) > 0 Then
      txtSearchName.text = ""
      Filter = " CompanyName LIKE '%" + UCase(txtSearchName.text) + "*'"
   End If
   
   Dim szSQL As String
   Dim adoConn As New ADODB.Connection

   adoConn.Open getConnectionString
   szSQL = "SELECT Tenants.SageAccountNumber, Tenants.Name, " & _
               "Tenants.CompanyName, UnitName, LeaseDetails.UnitNumber, LeaseDetails.Usage, " & _
               "ClientName, PropertyName, Property.PropertyID,Client.ClientID  " & _
           "FROM LeaseDetails, Units, Property, Client, Tenants  " & _
           "WHERE LeaseDetails.UnitNumber = Units.UnitNumber And " & _
               "LeaseDetails.Status = True And " & _
               "Units.PropertyId = Property.PropertyID And " & _
               "Property.ClientID = Client.ClientID AND " & _
               "LeaseDetails.SageAccountNumber = Tenants.SageAccountNumber " & _
               "" & szWhere & " " & _
           "ORDER BY " & szOrderBy & ";"
   
   Dim adoRst As New ADODB.Recordset
   
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   adoRst.Filter = Filter

   flxTenant.Clear
   flxTenant.Rows = adoRst.RecordCount + 1

  Dim iRow As Integer
  iRow = 1
  While Not adoRst.EOF
      flxTenant.TextMatrix(iRow, 1) = adoRst!SageAccountNumber
      flxTenant.TextMatrix(iRow, 2) = adoRst!Name
      flxTenant.TextMatrix(iRow, 3) = adoRst!ClientID
      adoRst.MoveNext
      iRow = iRow + 1
   Wend

   adoRst.Close
   Set adoRst = Nothing
   

   adoConn.Close
   Set adoConn = Nothing

End Function


Private Sub ConfigflxTenant()
   fmeTenantLookup.Visible = True
   flxTenant.Visible = True

'   flxTenant.RowHeight(0) = 350
   flxTenant.Cols = 4
   flxTenant.RowHeight(0) = 0
   flxTenant.row = 0

   flxTenant.ColWidth(0) = 100
   flxTenant.ColWidth(1) = Label20(2).Left - Label20(1).Left
   flxTenant.TextMatrix(0, 1) = "Sage A/C"
   flxTenant.ColAlignment(0) = vbLeftJustify

   flxTenant.ColWidth(2) = 3600
   flxTenant.TextMatrix(0, 2) = "Name"
   flxTenant.ColAlignment(1) = vbLeftJustify
   flxTenant.ColWidth(3) = 0 'client ID
End Sub


Private Function SiPi_Check1(adoConn As ADODB.Connection, szSiPi As String) As Boolean
   Dim szSQL      As String
   Dim adoRst     As New ADODB.Recordset
   Dim szTran2Fix As String

   If szSiPi = "PI" Then
      szSQL = "SELECT  P.TransactionID " & _
               "FROM tlbPayment AS P, (" & _
                     "SELECT PayHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                     "From tlbPaymentSplit " & _
                     "Group by PayHeader " & _
                     ") AS Q " & _
               "WHERE P.TransactionID = Q.PayHeader AND P.Amount <> P.OSAmount AND " & _
                     "ROUND(P.Amount - P.OSAmount, 2) <> Q.T;"
'Debug.Print szSQL
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRst.EOF
         szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("TransactionID").Value)

         adoRst.MoveNext
      Wend

      adoRst.Close
   End If

   If szSiPi = "SI" Then
      szSQL = "SELECT  R.TransactionID " & _
               "FROM tlbReceipt AS R, (" & _
                     "SELECT RptHeader, ROUND(Sum(Amount) - Sum(OSAmount), 2) AS T " & _
                     "From tlbReceiptSplit " & _
                     "Group by RptHeader " & _
                     ") AS Q " & _
               "WHERE R.TransactionID = Q.RptHeader AND R.Amount <> R.OSAmount AND " & _
                     "ROUND(R.Amount - R.OSAmount, 2) <> Q.T;"
'Debug.Print szSQL
      adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

      While Not adoRst.EOF
         szTran2Fix = szTran2Fix + ", " + CStr(adoRst.Fields.Item("TransactionID").Value)

         adoRst.MoveNext
      Wend

      adoRst.Close
   End If

   Set adoRst = Nothing

   If Len(szTran2Fix) > 0 Then szTran2Fix = Mid(szTran2Fix, 3)

   If Len(szTran2Fix) > 0 Then
        SiPi_Check1 = False
   Else
        SiPi_Check1 = True
        'MsgBox "HI"
   End If
      
End Function
Private Sub cmdSave_Click()
   Dim adoConn As New ADODB.Connection
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, dAmt As Double, i As Integer
'    If IsNull(cboClient.Column(0)) Then
'        MsgBox "Please select a proper client"
'        Exit Sub
'    End If
    If Trim(txtDate.text) = "" Then
        MsgBox "Please enter a valid date.", vbInformation, "Warning!!"
        FocusControl txtDate
        Exit Sub
    End If
     If lblRptPostingDate.ToolTipText = "" Then
        MsgBox "Please enter a valid posting date", vbInformation, "Warning!!"
        Exit Sub
    End If
    If DateDiff("d", lblRptPostingDate.ToolTipText, txtDate.text) > 0 Then
          MsgBox "Posting date cannot be before the transaction date", vbInformation, "Posting Date"
          Exit Sub
    End If
        adoConn.Open getConnectionString
        If IsPeriodStatus(lblRptPostingDate.ToolTipText, txtClient.text, adoConn) = 0 Then
           ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(lblRptPostingDate.ToolTipText, txtClient.text, adoConn) = 9 Then
           ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        End If
        If adoConn.State = 1 Then
                adoConn.Close
        End If
   Dim szTransID As String
   If frmDemands3.tabPayment.Tab = 0 Then
           adoConn.Open getConnectionString
           If frmDemands3.sEditSRR = 1 Then
        '--------------------------------For Type 23-----flxSPayment------------------------------------------------
                 adoConn.BeginTrans
                szSQL = "SELECT * " & _
                        "FROM tlbReceiptSplit " & _
                        "WHERE RptHeader = " & _
                            TransactionID & ";"
                
                adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                dAmt = 0
                For i = 1 To flxReceiptSplit.Rows - 1
                          'issue 357 receipt split was not found and issue 380
                    If flxReceiptSplit.TextMatrix(i, 1) <> "" Then
                         If adoRst.EOF Then
                '                if msgbox "Do you want to add this "
                             adoRst.AddNew
                             adoRst.Fields.Item("RptHeader").Value = Trim(TransactionID)
                             adoRst.Fields.Item("TransactionID").Value = UniqueID() 'New ID '
                         End If
                         adoRst.Fields.Item("SplitID").Value = i
                         adoRst.Fields.Item("FundID").Value = flxReceiptSplit.TextMatrix(i, 1)
                         adoRst.Fields.Item("Amount").Value = flxReceiptSplit.TextMatrix(i, 6)
                         adoRst.Fields.Item("OSAmount").Value = flxReceiptSplit.TextMatrix(i, 6)
                         dAmt = dAmt + flxReceiptSplit.TextMatrix(i, 6)
                         adoRst.Fields.Item("Description").Value = flxReceiptSplit.TextMatrix(i, 4)
                         adoRst.Update
                    End If
                Next i
                adoRst.Close
                
                '  updating the header record
                szSQL = "SELECT * " & _
                        "FROM tlbReceipt " & _
                        "WHERE TransactionID = " & _
                            TransactionID & ";"
                adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                '  Resolved By BOSL. Issue: 0000503. Modified by Asif
                '   Dim szTransID As String
                szTransID = TransactionID
                '   DeleteJournalNLPosting adoConn, szTransID
                
                ''For Type 23
                ''Fixed by anol 20170801 Issue related to 422
                ''.Fields.Item("PARENT_RECORD").Value = receipt("TransactionID").Value
                '' .Fields.Item("TRANS_ID").Value = receipt("TransactionID").Value
                ''
                ''For 3,4 This is posting from Demand receipt table
                '' .Fields.Item("PARENT_RECORD").Value = ReceiptSplit("TransactionID").Value
                '' .Fields.Item("TRANS_ID").Value = receipt("TransactionID").Value
                
                '    adoConn.Execute "UPDATE NLPosting AS N " & _
                '            "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & szTransID & "';"
                
                adoConn.Execute "UPDATE NLPosting AS N " & _
                         "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & szTransID & "' AND N.TRANSACTION_TYPE = 23;"
                '''
                
                With adoRst.Fields
                   .Item("SageAccountNumber").Value = txtLessee.text
                   .Item("BankCode").Value = cmbBankAc.Value
                   'added by anol 20170424 issue 361 . The list of bank accounts is not filtering correctly and the wrong bank account is being posted to in the nominal ledger.
                   .Item("NominalCode").Value = cmbBankAc.Value
                   .Item("RptAmtType").Value = cmbRptAmtType.Value
                   .Item("Amount").Value = txtAmount.text
                   .Item("OSAmount").Value = .Item("Amount").Value
                   .Item("RDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                   .Item("ExtRef").Value = txtRef.text
                   'Resolved by BOSL
                   'Below line added by anol 29 Mar 2015
                   'issue 549: Demand receipts not working note 3
                   .Item("PostingDate").Value = Format(lblRptPostingDate.ToolTipText, "dd mmmm yyyy")
                   .Item("FundID").Value = cmbFund.Column(0)
                   .Item("Details").Value = txtDetails.text
                   
                   '  Resolved By BOSL. Issue: 0000503. Modified by Asif
                   .Item("NLPost").Value = False
                    '  Issue: 0000571 note 1148. Modified by anol 25 aug 2015
                   .Item("ClientID").Value = txtClient.text
                End With
                adoRst.Update
                adoRst.Close
                      
                Set adoRst = Nothing
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit."
                     Exit Sub
                Else
                     adoConn.CommitTrans
                     frmMMain.Leasee1_LesseList_isUptoDate = False
                     frmMMain.Leasee4_LesseList_isUptoDate = False
                     frmMMain.frmDemand3_LesseList_isUptoDate = False
                End If
                
               
                
                cmdEdit.Caption = "&Edit"
                frmDemands3.AfterEditReceipt adoConn
                cmdCancel_Click
                 
                adoConn.Close
                Set adoConn = Nothing
                '--------------------------------------------------------------------------------------------
                '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
                adoConn.Open getConnectionString
                Export_SRnSRR_2_NL adoConn
                adoConn.Close
                frmDemands3.cmdEditReceipt.Enabled = False
          End If
        '------------------------------------For type 3,4------------From:flxCrPoA-------------------------------------
         If frmDemands3.sEditSRR = 2 Then
                adoConn.BeginTrans
                '  updating all splits
                szSQL = "SELECT * " & _
                        "FROM tlbReceiptSplit " & _
                        "WHERE RptHeader = " & _
                            TransactionID & ";"
                adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                '  Resolved By BOSL. Issue: 0000503. Modified by Asif
                szTransID = TransactionID
                
                dAmt = 0
                For i = 1 To flxReceiptSplit.Rows - 1
                '      If Left(frmDemands3.flxCrPoA.TextMatrix(frmDemands3.flxCrPoA.row, 0), 2) <> "SA" Then
                    adoRst.Find "TransactionID = '" & flxReceiptSplit.TextMatrix(i, 0) & "'", , , 1
                    If flxReceiptSplit.TextMatrix(i, 1) <> "" Then
                    'below clause is added by anol 20170422
                    'issue 357 receipt split was not found and issue 380
                         If adoRst.EOF Then
                '                if msgbox "Do you want to add this "
                             adoRst.AddNew
                             adoRst.Fields.Item("RptHeader").Value = Trim(TransactionID)
                             adoRst.Fields.Item("TransactionID").Value = UniqueID() 'New ID '
                         End If
                         adoRst.Fields.Item("SplitID").Value = i
                         adoRst.Fields.Item("FundID").Value = flxReceiptSplit.TextMatrix(i, 1)
                         adoRst.Fields.Item("Amount").Value = flxReceiptSplit.TextMatrix(i, 6)
                         adoRst.Fields.Item("OSAmount").Value = flxReceiptSplit.TextMatrix(i, 6)
                         dAmt = dAmt + flxReceiptSplit.TextMatrix(i, 6)
                         adoRst.Fields.Item("Description").Value = flxReceiptSplit.TextMatrix(i, 4)
                         adoRst.Update
                        
                   End If
                Next i
                adoRst.Close
                'added by anol 20170801 issue 422
                         adoConn.Execute "UPDATE NLPosting AS N " & _
                         "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & szTransID & "' AND (N.TRANSACTION_TYPE = 3 OR  N.TRANSACTION_TYPE = 4);"
                
                '   DeleteJournalNLPosting adoConn, szTransID
                
                '
                '   adoConn.Execute "UPDATE NLPosting AS N " & _
                '            "SET    N.DeleteFlag = TRUE " & _
                '            "WHERE  N.TRANS_ID = '" & szTransID & "';"
                         
                '''
                
                '  updating the header record
                szSQL = "SELECT * " & _
                        "FROM tlbReceipt " & _
                        "WHERE TransactionID = " & _
                            TransactionID & ";"
                adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                With adoRst.Fields
                   .Item("SageAccountNumber").Value = txtLessee.text
                   .Item("BankCode").Value = cmbBankAc.Value
                     'added by anol 20170424 issue 361 . The list of bank accounts is not filtering correctly and the wrong bank account is being posted to in the nominal ledger.
                   .Item("NominalCode").Value = cmbBankAc.Value
                   .Item("RptAmtType").Value = cmbRptAmtType.Value
                   .Item("Amount").Value = AllSplitTotal
                   .Item("OSAmount").Value = .Item("Amount").Value
                   .Item("RDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                   .Item("ExtRef").Value = txtRef.text
                   'Resolved by BOSL
                   'Below line added by anol 29 Mar 2015
                   'issue 549: Demand receipts not working note 3
                   .Item("PostingDate").Value = Format(lblRptPostingDate.ToolTipText, "dd mmmm yyyy")
'                   If Left(frmDemands3.flxCrPoA.TextMatrix(frmDemands3.flxCrPoA.row, 0), 2) = "SA" Then
                      .Item("FundID").Value = cmbFund.Column(0)
                      .Item("Details").Value = txtDetails.text
'                   End If
                   
                   '  Resolved By BOSL. Issue: 0000503. Modified by Asif
                   .Item("NLPost").Value = False
                    '  Issue: 0000571 note 1148. Modified by anol 25 aug 2015
                   .Item("ClientID").Value = txtClient.text
                End With
                adoRst.Update
                adoRst.Close
                
                Set adoRst = Nothing
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit Error! On Type  3,4"
                     Exit Sub
                Else
                     adoConn.CommitTrans
                     frmMMain.Leasee1_LesseList_isUptoDate = False
                     frmMMain.Leasee4_LesseList_isUptoDate = False
                     frmMMain.frmDemand3_LesseList_isUptoDate = False
                End If
                
               
                
                cmdEdit.Caption = "&Edit"
                frmDemands3.AfterEditReceipt adoConn
                cmdCancel_Click
                
                adoConn.Close
                Set adoConn = Nothing
                
                '--------------------------------------------------------------------------------------------
                '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
                adoConn.Open getConnectionString
                Export_SRnSRR_2_NL adoConn
                adoConn.Close
           End If
           frmDemands3.cmdEditReceipt.Enabled = False
    End If
    If frmDemands3.tabPayment.Tab = 1 Then
         adoConn.Open getConnectionString
           
        '--------------------------------For Type 23-----flxSPayment------------------------------------------------
                 adoConn.BeginTrans
                szSQL = "SELECT * " & _
                        "FROM tlbReceiptSplit " & _
                        "WHERE RptHeader = " & _
                            TransactionID & ";"
                
                adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                dAmt = 0
                For i = 1 To flxReceiptSplit.Rows - 1
                          'issue 357 receipt split was not found and issue 380
                    If flxReceiptSplit.TextMatrix(i, 1) <> "" Then
                         If adoRst.EOF Then
                '                if msgbox "Do you want to add this "
                             adoRst.AddNew
                             adoRst.Fields.Item("RptHeader").Value = Trim(TransactionID)
                             adoRst.Fields.Item("TransactionID").Value = UniqueID() 'New ID '
                         End If
                         adoRst.Fields.Item("SplitID").Value = i
                         adoRst.Fields.Item("FundID").Value = flxReceiptSplit.TextMatrix(i, 1)
                         adoRst.Fields.Item("Amount").Value = flxReceiptSplit.TextMatrix(i, 6)
                         adoRst.Fields.Item("OSAmount").Value = flxReceiptSplit.TextMatrix(i, 6)
                         dAmt = dAmt + flxReceiptSplit.TextMatrix(i, 6)
                         adoRst.Fields.Item("Description").Value = flxReceiptSplit.TextMatrix(i, 4)
                         adoRst.Update
                    End If
                Next i
                adoRst.Close
                
                '  updating the header record
                szSQL = "SELECT * " & _
                        "FROM tlbReceipt " & _
                        "WHERE TransactionID = " & _
                            TransactionID & ";"
                adoRst.Open szSQL, adoConn, adOpenDynamic, adLockOptimistic
                
                '  Resolved By BOSL. Issue: 0000503. Modified by Asif
                '   Dim szTransID As String
                szTransID = TransactionID
                '   DeleteJournalNLPosting adoConn, szTransID
                
                ''For Type 23
                ''Fixed by anol 20170801 Issue related to 422
                ''.Fields.Item("PARENT_RECORD").Value = receipt("TransactionID").Value
                '' .Fields.Item("TRANS_ID").Value = receipt("TransactionID").Value
                ''
                ''For 3,4 This is posting from Demand receipt table
                '' .Fields.Item("PARENT_RECORD").Value = ReceiptSplit("TransactionID").Value
                '' .Fields.Item("TRANS_ID").Value = receipt("TransactionID").Value
                
                '    adoConn.Execute "UPDATE NLPosting AS N " & _
                '            "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & szTransID & "';"
                
                adoConn.Execute "UPDATE NLPosting AS N " & _
                         "SET    N.DeleteFlag = TRUE " & _
                         "WHERE  N.TRANS_ID = '" & szTransID & "' AND (N.TRANSACTION_TYPE = 3 OR  N.TRANSACTION_TYPE = 4 OR N.TRANSACTION_TYPE = 23);"
                '''
                
                With adoRst.Fields
                   .Item("SageAccountNumber").Value = txtLessee.text
                   .Item("BankCode").Value = cmbBankAc.Value
                   'added by anol 20170424 issue 361 . The list of bank accounts is not filtering correctly and the wrong bank account is being posted to in the nominal ledger.
                   .Item("NominalCode").Value = cmbBankAc.Value
                   .Item("RptAmtType").Value = cmbRptAmtType.Value
                   .Item("Amount").Value = txtAmount.text
                   .Item("OSAmount").Value = .Item("Amount").Value
                   .Item("RDate").Value = Format(txtDate.text, "dd mmmm yyyy")
                   .Item("ExtRef").Value = txtRef.text
                   'Resolved by BOSL
                   'Below line added by anol 29 Mar 2015
                   'issue 549: Demand receipts not working note 3
                   .Item("PostingDate").Value = Format(lblRptPostingDate.ToolTipText, "dd mmmm yyyy")
                   .Item("FundID").Value = cmbFund.Column(0)
                   .Item("Details").Value = txtDetails.text
                   
                   '  Resolved By BOSL. Issue: 0000503. Modified by Asif
                   .Item("NLPost").Value = False
                    '  Issue: 0000571 note 1148. Modified by anol 25 aug 2015
                   .Item("ClientID").Value = txtClient.text
                End With
                adoRst.Update
                adoRst.Close
                      
                Set adoRst = Nothing
                If SiPi_Check1(adoConn, "SI") = False Then
                     adoConn.RollbackTrans
                     adoConn.Close
                     MsgBox "Could not save the transaction, Please contact with PCM consulting Ltd.", vbInformation, "Receipt Edit."
                     Exit Sub
                Else
                     adoConn.CommitTrans
                     'frmDemands3.flxSPayment.Clear
                     'frmDemands3.flxCrPoA.Clear
                     frmDemands3.flxReceiptHistory.TextMatrix(frmDemands3.flxReceiptHistory.row, 7) = Format(dAmt, "0.00")
                     frmMMain.Leasee1_LesseList_isUptoDate = False
                     frmMMain.Leasee4_LesseList_isUptoDate = False
                     frmMMain.frmDemand3_LesseList_isUptoDate = False
                End If
               
                
                cmdEdit.Caption = "&Edit"
                frmDemands3.AfterEditReceipt adoConn 'reload the grid
                cmdCancel_Click
                 
                adoConn.Close
                Set adoConn = Nothing
                '--------------------------------------------------------------------------------------------
                '  Export Sales receipt and sales receipt refund Transactions to Nominal Ledger (NLPosting table)
                adoConn.Open getConnectionString
                Export_SRnSRR_2_NL adoConn
                adoConn.Close
                frmDemands3.cmdEditReceipt.Enabled = False
        
    End If
   
End Sub

Private Function AllSplitTotal() As Currency
   If flxReceiptSplit.Rows = 2 Then
      AllSplitTotal = CCur(Val(txtAmount.text))
      Exit Function
   End If
   
   Dim iRow As Integer

   For iRow = 1 To flxReceiptSplit.Rows - 1
      AllSplitTotal = AllSplitTotal + CCur(flxReceiptSplit.TextMatrix(iRow, 6))
   Next iRow
End Function

Private Sub flxTenant_GotFocus()
   If flxTenant.row = 0 Then flxTenant.row = 1
End Sub

Private Sub flxTenant_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then fmeTenantLookup.Visible = False
   If KeyAscii = 13 Then flxTenant_Click
End Sub

Private Sub ConfigFlxReceiptSplit()
   Dim szHeader As String
   
   With flxReceiptSplit
      .Clear
      .FormatString = szHeader$
      .Rows = 2
      .RowHeight(0) = 0
      szHeader$ = "TransactionID|<FundID|<No|<FundName|<Description|<DueDate|>Amount||ReconNow"
'                       0           1     2      3           4          5        6   7    8
      .Cols = 10
      .ColWidth(0) = 0
      .ColWidth(1) = 0
      .ColWidth(2) = 0                                      'No
      .ColWidth(3) = Label2(3).Left - Label2(2).Left        'Fund Name
      .ColWidth(4) = Label2(4).Left - Label2(3).Left        'Description
      .ColWidth(5) = Label2(5).Left - Label2(4).Left        'Due Date
      .ColWidth(6) = .Left + .Width - Label2(5).Left - 300  'Amount
      .ColWidth(7) = 0                                      'Flag for edit
      .ColWidth(8) = 0                                      'Reconciled?
      .ColWidth(9) = 0                                      'posting date
   End With
End Sub

Public Sub LoadFlxReceiptSplit(adoConn As ADODB.Connection)
    Dim szSQL As String
    If BoolReconciled Then 'if transaction is reconciled only some part of this line can be edited issue 515
      txtAmount.Locked = True
      cmbBankAc.Locked = True
      txtClient.Locked = True
      cmdLesseeLookup.Enabled = False
      cmbFund.Locked = False
      txtDate.Locked = True
   Else
      txtAmount.Locked = False
      cmbBankAc.Locked = False
      txtClient.Locked = False
      cmdLesseeLookup.Enabled = True
      cmbFund.Locked = False
      txtDate.Locked = False
   End If
    If frmDemands3.tabPayment.Tab = 0 Then
        If frmDemands3.sEditSRR = 1 Then
            szSQL = "SELECT R.TransactionID, R.FundID, '' AS No, F.FundName, R.Details AS Description, " & _
                "R.DDate AS DueDate, FORMAT(R.Amount,'0.00') AS AMOUNT, R.ReconNow, R.POSTINGDATE " & _
            "FROM tlbReceipt AS R INNER JOIN Fund AS F ON R.FundID = F.FundID " & _
            "WHERE R.TransactionID = " & TransactionID & " AND " & _
                "R.Type = 23 " & _
            "ORDER BY R.TransactionID;"
            iRecords = populateGridDefinedHeader(adoConn, szSQL, flxReceiptSplit)
            If iRecords > 0 Then
                cmbFund.ListIndex = FindComboIndex(cmbFund, flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 1), 0)
                txtAmount.text = Format(Val(flxReceiptSplit.TextMatrix(1, 6)), "0.00")
                txtDetails.text = flxReceiptSplit.TextMatrix(1, 4)
                'lblRptPostingDate.ToolTipText = flxReceiptSplit.TextMatrix(1, 9)
            End If
        End If
        If frmDemands3.sEditSRR = 2 Then
            szSQL = "SELECT S.TransactionID, S.FundID, S.SplitID AS No, F.FundName, S.Description, " & _
                "S.DueDate, FORMAT(S.Amount,'0.00') AS AMOUNT, '', R.ReconNow, R.POSTINGDATE " & _
            "FROM (tlbReceiptSplit AS S INNER JOIN Fund AS F ON S.FundID = F.FundID) INNER JOIN " & _
                "tlbReceipt AS R ON S.RptHeader = R.TransactionID " & _
            "WHERE R.TransactionID = " & TransactionID & " AND " & _
                "(R.Type = 3 OR R.Type = 4) " & _
            "ORDER BY S.SplitID;"
            iRecords = populateGridDefinedHeader(adoConn, szSQL, flxReceiptSplit)
            If iRecords > 0 Then
                cmbFund.ListIndex = FindComboIndex(cmbFund, flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 1), 0)
                txtAmount.text = Format(Val(flxReceiptSplit.TextMatrix(1, 6)), "0.00")
                txtDetails.text = flxReceiptSplit.TextMatrix(1, 4)
                'lblRptPostingDate.ToolTipText = flxReceiptSplit.TextMatrix(1, 9)
            End If
        End If
    ElseIf frmDemands3.tabPayment.Tab = 1 Then
        szSQL = "SELECT S.TransactionID, S.FundID, S.SplitID AS No, F.FundName, S.Description, " & _
                "S.DueDate, FORMAT(S.Amount,'0.00') AS AMOUNT, '', R.ReconNow, R.POSTINGDATE " & _
            "FROM (tlbReceiptSplit AS S INNER JOIN Fund AS F ON S.FundID = F.FundID) INNER JOIN " & _
                "tlbReceipt AS R ON S.RptHeader = R.TransactionID " & _
            "WHERE R.TransactionID = " & TransactionID & " AND " & _
                "(R.Type = 3 OR R.Type = 4 OR R.Type = 23) " & _
            "ORDER BY S.SplitID;"
            iRecords = populateGridDefinedHeader(adoConn, szSQL, flxReceiptSplit)
            If iRecords > 0 Then
                cmbFund.ListIndex = FindComboIndex(cmbFund, flxReceiptSplit.TextMatrix(flxReceiptSplit.row, 1), 0)
                txtAmount.text = Format(Val(flxReceiptSplit.TextMatrix(1, 6)), "0.00")
                txtDetails.text = flxReceiptSplit.TextMatrix(1, 4)
                'lblRptPostingDate.ToolTipText = flxReceiptSplit.TextMatrix(1, 9)
            End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
   Me.Height = 6075
   Me.Width = 7980
   Me.BackColor = MODULEBACKCOLOR
   Label1(1).BackColor = MODULEBACKCOLOR
   frmMMain.Arrange vbCascade
   Me.ZOrder 0
   fraEdit.BackColor = Me.BackColor
   ConfigFlxReceiptSplit
    
   Dim adoConn As New ADODB.Connection


   adoConn.Open getConnectionString
   'Below line added by anol 25 aug 2015
   'issue 571 note 1148
'    PrepareCboClient adoConn
    'End of Modi
'   PrepareList adoConn             'prepare the list of clients and properties in dwopdown comboes
   'LoadBank_Fund adoConn

   'LoadFlxReceiptSplit adoConn
   
   adoConn.Close
   Set adoConn = Nothing
   'flxReceiptSplit.row = 1
   'flxReceiptSplit_RowColChange
   'Call WheelHook(Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   frmDemands3.Enabled = True
   frmDemands3.editexception = False
   'added by anol 02 Sep 2015
   'frmDemands3.ConfigFlxReceiptHistory
'   Call WheelUnHook(Me.hWnd)
End Sub

Private Sub fraEdit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MousePointer = vbArrow
End Sub

Private Sub lblRptPostingDate_DblClick(Cancel As MSForms.ReturnBoolean)
    'Resolved by BOSL
    'Below line added by anol 29 Mar 2015
    'issue 549: Demand receipts not working note 3
     DispayCalendar Me, lblRptPostingDate.ToolTipText, txtDate.text, frmDemands3.txtRptClientList.text
End Sub

Private Sub txtAmount_Click()
        SelTxtInCtrl txtAmount
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
   If KeyAscii = 45 Then
        KeyAscii = 0
   End If

   If txtAmount.Locked Then
      ShowMsgInTaskBar "The transactions has been reconciled and the amount cannot be edited", "Y", "N"
      Exit Sub
   End If
    If KeyAscii = 13 Then
        FocusControl txtDate
   End If
   If KeyAscii = 27 And txtAmount.text <> "" Then
      txtAmount.text = ""
   End If
   If KeyAscii = 27 And txtAmount.text = "" Then Unload Me
   DigitTextKeyPress txtAmount, KeyAscii
End Sub

Private Sub txtAmount_LostFocus()
    txtAmount.text = Format(Val(txtAmount.text), "0.00")
End Sub

Private Sub txtDate_Change()
   TextBoxChangeDate txtDate
    'Resolved by BOSL
    'Added by anol 29 Mar 2015
    'issue 549: Demand receipts not working
   lblRptPostingDate.ToolTipText = txtDate.text
End Sub

Private Sub txtDate_GotFocus()
   SelTxtInCtrl txtDate
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
    If txtDate.Locked Then
      ShowMsgInTaskBar "The transactions has been reconciled and the date cannot be edited", "Y", "N"
      Exit Sub
   End If
   If KeyAscii = 27 And txtDate.text <> "" Then txtDate.text = ""
   If KeyAscii = 27 And txtDate.text = "" Then Unload Me
   If KeyAscii = 13 Then
        FocusControl cmbRptAmtType
   End If
   TextBoxKeyPrsDate txtDate, KeyAscii
End Sub

Private Sub txtDate_LostFocus()
   TextBoxFormatDate txtDate
    'Resolved by BOSL
    'Below line added by anol 29 Mar 2015
    'issue 549: Demand receipts not working note 3
    If txtDate.text <> "" Then
      If TextBoxFormatDate(txtDate) Then
         If lblRptPostingDate.ToolTipText = "" And lblRptPostingDate.ToolTipText <> txtDate.text Then
            lblRptPostingDate.ToolTipText = txtDate.text
          End If
         
        If frmMMain.IsRibbonVersion And lblRptPostingDate.ToolTipText <> "" Then
        Dim adoConn As New ADODB.Connection
        Dim szSQL As String
        
        adoConn.Open getConnectionString
        'If IsNull(frmDemands3.txtRptClientList.text) Then Exit Sub
        If IsPeriodStatus(lblRptPostingDate.ToolTipText, txtClient.text, adoConn) = 0 Then
           ShowMsgInTaskBar "The posting date cannot fall within a closed financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        ElseIf IsPeriodStatus(lblRptPostingDate.ToolTipText, txtClient.text, adoConn) = 9 Then
           ShowMsgInTaskBar "The posting date does not fall in any existing financial period", "Y", "N"
           adoConn.Close
           Set adoConn = Nothing
           Exit Sub
        End If
        adoConn.Close
        Set adoConn = Nothing
        End If
      End If
   End If
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
   If KeyAscii = 27 Then Unload Me
   If KeyAscii = 13 Then
        FocusControl txtRef
   End If
End Sub

Private Sub txtLessee_GotFocus()
   FocusControl cmdLesseeLookup
End Sub

Public Sub LoadBank_Fund_amtType(adoConn As ADODB.Connection)
   ' Error Handler
   On Error GoTo Error_Handler

   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

   szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
               "tlbClientBanks.Bank_AC_Name " & _
           "FROM tlbClientBanks;"
'Debug.Print szSQL

'frmDemands3.cboRptClientList.value

'   szSQL = "SELECT CB.NominalCode AS BNC,NL.Name as Bank_AC_Name " & _
'   "FROM tlbClientBanks AS CB, NominalLedger as NL where  NL.ClientID = CB.CLIENT_ID AND NL.ClientID ='" & frmDemands3.cboRptClientList.Value & "' AND NL.Code =CB.NominalCode order by CB.NominalCode "
 szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
                  "NominalLedger.Name AS Bank_AC_Name, tlbClientBanks.CurrentBalance AS BAL " & _
              "FROM tlbClientBanks, NominalLedger " & _
              "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
                  "tlbClientBanks.CLIENT_ID = '" & txtClient.text & "' AND " & _
                  "NominalLedger.ClientID = '" & txtClient.text & "' " & _
              "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance;"
 
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   ReDim szaData(1, adoRst.RecordCount - 1) As String

   While Not adoRst.EOF
      szaData(0, iRec) = adoRst.Fields.Item("BNC").Value
      szaData(1, iRec) = adoRst.Fields.Item("Bank_AC_Name").Value
      iRec = iRec + 1
      adoRst.MoveNext
   Wend

   cmbBankAc.Clear
   cmbBankAc.Column() = szaData()
   adoRst.Close

   szSQL = "SELECT FundID, FundName FROM FUND;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   ReDim szaData(1, adoRst.RecordCount - 1) As String

   iRec = 0
   While Not adoRst.EOF
      szaData(0, iRec) = adoRst.Fields.Item("FundID").Value
      szaData(1, iRec) = adoRst.Fields.Item("FundName").Value
      iRec = iRec + 1
      adoRst.MoveNext
   Wend
   adoRst.Close

   cmbFund.Clear
   cmbFund.Column() = szaData()

   szSQL = "SELECT SecondaryCode.Code as C, SecondaryCode.Value as V " & _
           "FROM PrimaryCode, SecondaryCode " & _
           "WHERE PrimaryCode.Value = 'RECEIPT AMOUNT TYPE' AND " & _
                 "PrimaryCode.CODE = SecondaryCode.PrimaryCode " & _
           "ORDER BY SecondaryCode.Value;"

   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly

   If adoRst.EOF Then
      adoRst.Close
      Set adoRst = Nothing
      Exit Sub
   End If

   ReDim szaData(1, adoRst.RecordCount - 1) As String

   iRec = 0
   While Not adoRst.EOF
      szaData(0, iRec) = adoRst!c
      szaData(1, iRec) = adoRst!V
      adoRst.MoveNext
      iRec = iRec + 1
   Wend
   adoRst.Close

   cmbRptAmtType.Clear
   cmbRptAmtType.Column() = szaData()

   ' Destroy Objects
   Set adoRst = Nothing

   Exit Sub

   ' Error Handling Code
Error_Handler:
   ' Destroy Objects
   Set adoRst = Nothing
End Sub
Private Sub LoadBankAgain(adoConn As ADODB.Connection)
   
   Dim iRec As Integer
   Dim adoRst As New ADODB.Recordset
   Dim szSQL As String, szaData() As String

    szSQL = "SELECT tlbClientBanks.NominalCode AS BNC, " & _
                  "NominalLedger.Name AS Bank_AC_Name, tlbClientBanks.CurrentBalance AS BAL " & _
              "FROM tlbClientBanks, NominalLedger " & _
              "WHERE tlbClientBanks.NominalCode = NominalLedger.Code AND " & _
                  "tlbClientBanks.CLIENT_ID = '" & txtClient.text & "' AND " & _
                  "NominalLedger.ClientID = '" & txtClient.text & "' " & _
              "GROUP BY tlbClientBanks.NominalCode, NominalLedger.Name, tlbClientBanks.CurrentBalance;"
 
   adoRst.Open szSQL, adoConn, adOpenStatic, adLockReadOnly
   
   ReDim szaData(1, adoRst.RecordCount - 1) As String

   While Not adoRst.EOF
      szaData(0, iRec) = adoRst.Fields.Item("BNC").Value
      szaData(1, iRec) = adoRst.Fields.Item("Bank_AC_Name").Value
      iRec = iRec + 1
      adoRst.MoveNext
   Wend

   cmbBankAc.Clear
   cmbBankAc.Column() = szaData()
   If cmbBankAc.ListCount > 0 Then
        cmbBankAc.ListIndex = 0
   End If
   adoRst.Close
   Set adoRst = Nothing
End Sub

Private Sub txtRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
         FocusControl txtAmount
    End If
End Sub

Private Sub txtSearchName_KeyPress(KeyAscii As MSForms.ReturnInteger)
   If KeyAscii = 27 Then fmeTenantLookup.Visible = False
End Sub

Private Sub txtSearchTenant_Change()
   FilterTenantsList
End Sub

Private Sub txtSearchTenant_KeyPress(KeyAscii As MSForms.ReturnInteger)
    
   If KeyAscii = 27 Then fmeTenantLookup.Visible = False
   If KeyAscii = 13 And Len(txtSearchTenant.text) > 0 Then
        FocusControl flxTenant
   ElseIf KeyAscii = 13 And Len(txtSearchName.text) > 0 Then
        FocusControl txtSearchName
   End If
End Sub

' Here you can add scrolling support to controls that don't normally respond.
' This Sub could always be moved to a module to make scrollwheel behaviour
' generic across forms.
' ===========================================================================
Public Sub MouseWheel(ByVal MouseKeys As Long, ByVal Rotation As Long, ByVal Xpos As Long, ByVal Ypos As Long)
  Dim ctl As Control
  Dim bHandled As Boolean
  Dim bOver As Boolean

  For Each ctl In Controls
    ' Is the mouse over the control
    On Error Resume Next
    bOver = (ctl.Visible And IsOver(ctl.hWnd, Xpos, Ypos))
    On Error GoTo 0

    If bOver Then
      ' If so, respond accordingly
      bHandled = True
      Select Case True

        Case TypeOf ctl Is MSHFlexGrid
          FlexGridScroll ctl, MouseKeys, Rotation, Xpos, Ypos

        Case TypeOf ctl Is PictureBox
          'PictureBoxZoom ctl, MouseKeys, Rotation, Xpos, Ypos
            bHandled = False
        Case TypeOf ctl Is ListBox, TypeOf ctl Is TextBox, TypeOf ctl Is ComboBox
          ' These controls already handle the mousewheel themselves, so allow them to:
          If ctl.Enabled Then ctl.SetFocus

        Case Else
          bHandled = False

      End Select
      If bHandled Then Exit Sub
    End If
    bOver = False
  Next ctl
'
'  ' Scroll was not handled by any controls, so treat as a general message send to the form
'  Me.Caption = "Form Scroll " & IIf(Rotation < 0, "Down", "Up")
End Sub


